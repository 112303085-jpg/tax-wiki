"""簡報流程（graph）的程式節點。規則見 AGENTS.md 第 7.3.1 節。

用法：python tools/deck.py <指令> slides/<主題>
  status           目前走到流程哪一步、下一步做什麼
  inventory        依 00_需求.md 盤點 wiki，產生 01_素材.md
  check-outline    檢查 02_大綱.md
  approve-outline  記錄使用者核准（只在使用者於對話中明確核准後執行）
  check-slides     檢查 03_逐頁.md（對照大綱、欄位、字數、圖貼、數字出處）
  render           由 03_逐頁.md 產生 <主題>.html 與 <主題>.pptx（排版在 deck_render.py、deck_pptx.ps1）
  check-render     以 Edge 無頭模式檢查 HTML 版面並逐頁截圖；讀取 PowerPoint 文字框檢查結果
  finalize         判定草稿／定稿
  ci               唯讀檢查（GitHub Actions 用）：內容重檢、成品是否過期、草稿／定稿是否與審閱狀態一致

只用 Python 標準函式庫；PowerPoint 需本機安裝 Office，HTML 檢查需 Microsoft Edge。
"""
import datetime
import hashlib
import json
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
WIKI = ROOT / "wiki"

LAYOUTS = ["封面", "目錄", "分隔頁", "條列", "兩欄比較", "大數字", "表格", "圖片", "例題", "常見誤解", "免責聲明"]
NO_SOURCE = {"封面", "目錄", "分隔頁", "免責聲明"}   # 不含事實、不需引用頁的版型
MAX_RETRY = 3
TITLE_MAX = 18      # 標題字數上限
MESSAGE_MAX = 45    # 主訊息字數上限
PAGE_RANGE = {"入門": (8, 25), "進階": (10, 35)}
NEED_KEYS = ["主題", "層級", "稅目", "格式", "時長分鐘"]
OUTLINE_COLS = ["頁", "版型", "標題", "主訊息", "引用頁"]
SKIP_WIKI = {"index.md", "log.md", "overview.md"}


# ---------- 共用 ----------

def roc_today():
    d = datetime.date.today()
    return f"{d.year - 1911}-{d.month:02d}-{d.day:02d}"


TEXT_SUFFIX = {".md", ".html", ".json", ".svg", ".ps1", ".py", ".txt", ".css", ".js"}


def fingerprint(path):
    """檔案指紋。文字檔先把 CRLF 統一成 LF，讓 Windows 本機與 GitHub（Linux）算出相同結果。"""
    data = path.read_bytes()
    if path.suffix.lower() in TEXT_SUFFIX:
        data = data.replace(b"\r\n", b"\n")
    return hashlib.sha256(data).hexdigest()[:12]


def frontmatter(text):
    m = re.match(r"﻿?---\r?\n(.*?)\r?\n---\r?\n", text, re.S)
    if not m:
        return {}, text
    data = {}
    for line in m.group(1).splitlines():
        if ":" not in line or line[:1].isspace():
            continue
        key, value = line.split(":", 1)
        value = re.split(r"\s+#", value, maxsplit=1)[0].strip()
        if value.startswith("[") and value.endswith("]"):
            value = [x.strip().strip('"').strip("'") for x in value[1:-1].split(",") if x.strip()]
        data[key.strip()] = value
    return data, text[m.end():]


def wiki_pages():
    pages = {}
    for p in WIKI.rglob("*.md"):
        if p.name in SKIP_WIKI:
            continue
        pages[p.relative_to(WIKI).with_suffix("").as_posix()] = p
    return pages


def resolve(name, pages):
    """[[頁名]] 或 [[資料夾/頁名]] → 檔案路徑；找不到或有歧義回傳 None。"""
    name = name.strip().strip("[]").split("|")[0].strip()
    if name in pages:
        return name
    hits = [rel for rel in pages if rel.split("/")[-1] == name]
    return hits[0] if len(hits) == 1 else None


def page_info(path):
    fm, body = frontmatter(path.read_text(encoding="utf-8"))
    todo = []
    m = re.search(r"^## 待查證\s*\n(.*?)(?=^## |\Z)", body, re.S | re.M)
    if m:
        for ln in m.group(1).splitlines():
            if re.match(r"^>\s*\[!", ln):          # callout 標題列
                continue
            item = re.sub(r"^>?\s*(\d+\.|[-*])?\s*", "", ln).strip()
            if item and item not in ("（無）",):
                todo.append(item)
    return {
        "類型": fm.get("類型", ""),
        "稅目": fm.get("稅目", []) if isinstance(fm.get("稅目"), list) else [fm.get("稅目", "")],
        "查證狀態": fm.get("查證狀態", ""),
        "審閱": fm.get("審閱", ""),
        "可用於簡報": fm.get("可用於簡報", ""),
        "待查證": todo,
    }


def usability(info):
    if info["可用於簡報"] == "是":
        return "可用（定稿）"
    if info["查證狀態"] == "已查證":
        return "可用（僅草稿）"
    return "不可用"


def load_state(deck):
    p = deck / "state.json"
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8-sig"))
    return {"retries": {}}


def save_state(deck, state):
    (deck / "state.json").write_text(json.dumps(state, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def report(deck, title, lines):
    p = deck / "04_檢查報告.md"
    head = "" if p.exists() else "# 檢查報告\n\n> deck.py 自動產生，只增不改。\n"
    stamp = datetime.datetime.now().strftime("%H:%M")
    with p.open("a", encoding="utf-8") as f:
        f.write(head + f"\n## [{roc_today()} {stamp}] {title}\n" + "\n".join(lines) + "\n")


def read_need(deck):
    p = deck / "00_需求.md"
    if not p.exists():
        return None, ["找不到 00_需求.md"]
    fm, _ = frontmatter(p.read_text(encoding="utf-8"))
    errs = [f"00_需求.md 缺少欄位：{k}" for k in NEED_KEYS if not fm.get(k)]
    if fm.get("層級") and fm["層級"] not in PAGE_RANGE:
        errs.append(f"00_需求.md 層級須為 {'或'.join(PAGE_RANGE)}")
    if isinstance(fm.get("稅目"), str):
        fm["稅目"] = [fm["稅目"]]
    return fm, errs


# ---------- 節點：inventory ----------

def cmd_inventory(deck):
    need, errs = read_need(deck)
    if errs:
        return fail(errs)
    pages = wiki_pages()
    rows, gaps = [], []
    for rel, p in sorted(pages.items()):
        info = page_info(p)
        if not set(info["稅目"]) & set(need["稅目"]):
            continue
        use = usability(info)
        rows.append(f"| [[{rel}]] | {info['類型']} | {info['查證狀態']} | {info['審閱']} | {use} |")
        if use == "不可用":
            items = "；".join(info["待查證"][:3]) or "（頁面未列出待查證項目）"
            gaps.append(f"- [[{rel}]]（{info['查證狀態']}）：{items}")
    out = [
        "---", f"主題: {need['主題']}", f"產生日: {roc_today()}", "產生者: tools/deck.py inventory", "---", "",
        f"# {need['主題']}｜素材盤點", "",
        f"稅目：{'、'.join(need['稅目'])}。只有「可用」的頁可以寫進大綱；草稿可用「可用（僅草稿）」的頁。", "",
        "| 頁面 | 類型 | 查證狀態 | 審閱 | 本次可否使用 |", "|---|---|---|---|---|", *rows, "",
        "## 缺口（不可用的頁與原因）", "", *(gaps or ["（無）"]), "",
    ]
    (deck / "01_素材.md").write_text("\n".join(out), encoding="utf-8")
    state = load_state(deck)
    state["inventory"] = {"date": roc_today(), "need": fingerprint(deck / "00_需求.md")}
    save_state(deck, state)
    usable = sum("不可用" not in r for r in rows)
    print(f"已產生 01_素材.md：相關頁 {len(rows)}，可用 {usable}，不可用 {len(rows) - usable}")
    return 0


# ---------- 節點：check-outline ----------

def parse_outline(text):
    fm, body = frontmatter(text)
    rows = []
    for line in body.splitlines():
        s = line.strip()
        if not s.startswith("|"):
            continue
        cells = [c.strip() for c in s.strip("|").split("|")]
        if cells and (cells[0] == "頁" or set("".join(cells)) <= set("-: ")):
            continue
        rows.append(cells)
    return fm, rows


def cmd_check_outline(deck):
    state = load_state(deck)
    need, errs = read_need(deck)
    outline_path = deck / "02_大綱.md"
    if not outline_path.exists():
        errs.append("找不到 02_大綱.md")
    if "inventory" not in state or not (deck / "01_素材.md").exists():
        errs.append("尚未執行 inventory")
    elif need and state["inventory"]["need"] != fingerprint(deck / "00_需求.md"):
        errs.append("00_需求.md 在盤點後被修改，請重新執行 inventory")
    if errs:
        return fail(errs)

    fm, rows = parse_outline(outline_path.read_text(encoding="utf-8"))
    pages = wiki_pages()
    warns = []
    if fm.get("主題") != need["主題"]:
        errs.append(f"大綱主題「{fm.get('主題')}」與需求「{need['主題']}」不一致")
    if fm.get("層級") != need["層級"]:
        errs.append(f"大綱層級「{fm.get('層級')}」與需求「{need['層級']}」不一致")
    lo, hi = PAGE_RANGE[need["層級"]]
    if not lo <= len(rows) <= hi:
        errs.append(f"頁數 {len(rows)} 不在{need['層級']}場的範圍 {lo}～{hi} 頁")

    cited, titles = set(), {}
    for i, cells in enumerate(rows, 1):
        tag = f"第{i}列"
        if len(cells) != len(OUTLINE_COLS):
            errs.append(f"{tag}：欄數 {len(cells)}，應為 {len(OUTLINE_COLS)}（{'、'.join(OUTLINE_COLS)}）")
            continue
        no, layout, title, message, refs = cells
        if no != str(i):
            errs.append(f"{tag}：頁碼「{no}」應為 {i}")
        if layout not in LAYOUTS:
            errs.append(f"{tag}：版型「{layout}」不在允許清單（{'、'.join(LAYOUTS)}）")
        if not title:
            errs.append(f"{tag}：缺標題")
        elif len(title) > TITLE_MAX:
            errs.append(f"{tag}：標題 {len(title)} 字，超過 {TITLE_MAX} 字")
        if title in titles:
            errs.append(f"{tag}：標題與第{titles[title]}頁重複")
        titles.setdefault(title, i)
        if layout not in NO_SOURCE:
            if not message or message == "—":
                errs.append(f"{tag}：缺主訊息")
            elif len(message) > MESSAGE_MAX:
                errs.append(f"{tag}：主訊息 {len(message)} 字，超過 {MESSAGE_MAX} 字")
            names = re.findall(r"\[\[([^\]]+)\]\]", refs)
            if not names:
                errs.append(f"{tag}：「{layout}」版型須引用至少一個 wiki 頁")
            for n in names:
                rel = resolve(n, pages)
                if not rel:
                    errs.append(f"{tag}：找不到 wiki 頁 [[{n}]]（或名稱有歧義）")
                    continue
                info = page_info(pages[rel])
                if not set(info["稅目"]) & set(need["稅目"]):
                    errs.append(f"{tag}：[[{n}]] 的稅目不在需求範圍內")
                if usability(info) == "不可用":
                    errs.append(f"{tag}：[[{n}]] 為「{info['查證狀態']}」，不可用於簡報")
                else:
                    cited.add(rel)
    if rows:
        if rows[0][1:2] != ["封面"]:
            errs.append("第1頁版型須為「封面」")
        if rows[-1][1:2] != ["免責聲明"]:
            errs.append("最後一頁版型須為「免責聲明」")
    unreviewed = sorted(r for r in cited if usability(page_info(pages[r])) != "可用（定稿）")
    if unreviewed:
        warns.append(f"引用頁中有 {len(unreviewed)} 頁尚未審閱，成品將為草稿：" + "、".join(unreviewed))

    fp = fingerprint(outline_path)
    ok = not errs
    retries = state.setdefault("retries", {})
    retries["outline"] = 0 if ok else retries.get("outline", 0) + 1
    state["outline_check"] = {"ok": ok, "fingerprint": fp, "date": roc_today(), "cited": sorted(cited)}
    save_state(deck, state)
    lines = [f"- 大綱指紋 {fp}，{len(rows)} 頁，結果：{'通過' if ok else '未通過'}"]
    lines += [f"- 錯誤：{e}" for e in errs] + [f"- 提醒：{w}" for w in warns]
    report(deck, "check-outline", lines)
    print("\n".join(lines))
    if not ok and retries["outline"] >= MAX_RETRY:
        print(f"已連續 {retries['outline']} 次未通過，依規則停止，請向使用者回報。")
    return 0 if ok else 1


# ---------- 節點：approve-outline ----------

def cmd_approve_outline(deck):
    state = load_state(deck)
    chk = state.get("outline_check")
    outline_path = deck / "02_大綱.md"
    if not chk or not chk["ok"]:
        return fail(["大綱尚未通過 check-outline，不能核准"])
    if chk["fingerprint"] != fingerprint(outline_path):
        return fail(["大綱在檢查後被修改，請重新執行 check-outline"])
    pages = wiki_pages()
    state["outline_approved"] = {
        "date": roc_today(),
        "fingerprint": chk["fingerprint"],
        "pages": {rel: fingerprint(pages[rel]) for rel in chk["cited"]},
    }
    save_state(deck, state)
    report(deck, "approve-outline", [f"- 使用者核准大綱（指紋 {chk['fingerprint']}），記錄引用頁 {len(chk['cited'])} 頁的版本"])
    print(f"已記錄核准：大綱指紋 {chk['fingerprint']}，引用頁 {len(chk['cited'])} 頁")
    return 0


# ---------- 節點：check-slides ----------

# 版型 → (必填欄位, 選填欄位)；「主訊息」與「備註」另外處理
LAYOUT_FIELDS = {
    "封面": ({"副標", "圖貼"}, set()),
    "目錄": ({"項目", "圖貼"}, set()),
    "分隔頁": ({"小標", "圖貼"}, set()),
    "條列": ({"項目", "圖貼", "出處", "引用頁"}, {"說明"}),
    "兩欄比較": ({"左標", "左項目", "右標", "右項目", "圖貼", "出處", "引用頁"}, {"補充"}),
    "大數字": ({"數字", "圖貼", "出處", "引用頁"}, {"前言", "後語", "項目"}),
    "表格": ({"表格", "圖貼", "出處", "引用頁"}, {"說明", "附註"}),
    "圖片": ({"圖", "圖說", "圖貼", "出處", "引用頁"}, {"項目"}),
    "例題": ({"題目", "答案", "圖貼", "出處", "引用頁"}, {"步驟", "算式", "答案標籤", "說明"}),
    "常見誤解": ({"表格", "圖貼", "出處", "引用頁"}, set()),
    "免責聲明": ({"圖貼"}, set()),
}
LIST_FIELDS = {"項目", "左項目", "右項目", "補充", "圖貼"}
TABLE_FIELDS = {"表格", "步驟"}
LIST_MAX = {"項目": 5, "左項目": 3, "右項目": 3, "補充": 2, "圖貼": 4}
TEXT_MAX = {"說明": 60, "附註": 70, "題目": 60, "圖說": 80, "出處": 70, "數字": 12, "答案": 10}
TEXT_DEFAULT_MAX = 40
ITEM_MAX = 40
TABLE_ROWS_MAX = 6
CELL_MAX = 30
# 圖貼不得使用含佔位文字的素材（AGENTS.md 第 7.3 節）
ASSET_BANNED = ("assets/callouts/", "assets/backgrounds/")
ASSET_DIAGRAM_OK = {"assets/diagrams/01-arrow-right.svg", "assets/diagrams/02-arrow-curved.svg"}
NUM_RE = re.compile(r"(\d[\d,_]*(?:\.\d+)?)\s*(萬|億)?")
UNIT = {"萬": 10_000, "億": 100_000_000}


def numbers(text, keep_raw=False):
    """抽出文字中的數字（去掉千分位與底線）。帶「萬／億」者換算成完整數值；
    keep_raw=True 時同時保留換算前的寫法（用於 wiki 端，擴大比對範圍）。"""
    out = set()
    for n, unit in NUM_RE.findall(text):
        n = n.replace(",", "").replace("_", "").rstrip(".")
        if unit:
            v = float(n) * UNIT[unit]
            out.add(str(int(v)) if v == int(v) else str(v))
            if keep_raw:
                out.add(n)
        else:
            out.add(n)
    return out


def parse_slides(text):
    fm, body = frontmatter(text)
    slides, cur, field = [], None, None
    for raw in body.splitlines():
        m = re.match(r"^## (\d+)｜(.+?)｜(.+?)\s*$", raw)
        if m:
            cur = {"no": int(m.group(1)), "layout": m.group(2), "title": m.group(3), "fields": {}, "order": []}
            slides.append(cur)
            field = None
            continue
        if cur is None:
            continue
        m = re.match(r"^- ([^:：\s]+)[:：][ \t]*(.*)$", raw)
        if m:
            field = m.group(1)
            cur["order"].append(field)
            cur["fields"][field] = m.group(2).strip() if m.group(2).strip() else None
            continue
        m = re.match(r"^\s{2,}- (.*)$", raw)
        if m and field:
            val = cur["fields"].get(field)
            cur["fields"][field] = (val if isinstance(val, list) else []) + [m.group(1).strip()]
            continue
        m = re.match(r"^\s{2,}\|(.*)\|\s*$", raw)
        if m and field:
            cells = [c.strip() for c in m.group(1).split("|")]
            if set("".join(cells)) <= set("-: "):
                continue
            val = cur["fields"].get(field)
            cur["fields"][field] = (val if isinstance(val, dict) else {"rows": []})
            cur["fields"][field]["rows"].append(cells)
    return fm, slides


def outline_by_page(deck):
    _, rows = parse_outline((deck / "02_大綱.md").read_text(encoding="utf-8"))
    return {int(r[0]): {"layout": r[1], "title": r[2], "message": r[3],
                        "refs": re.findall(r"\[\[([^\]]+)\]\]", r[4])} for r in rows if len(r) == 5}


def stale_pages(state):
    appr = state.get("outline_approved") or {}
    pages = wiki_pages()
    return [r for r, fp in appr.get("pages", {}).items() if r not in pages or fingerprint(pages[r]) != fp]


def approval_problems(deck, state, allow_stale=False):
    """allow_stale=True 用於 check-slides：大綱沒改、只有引用頁被修改時，允許重新檢查。"""
    appr = state.get("outline_approved")
    if not appr:
        return ["大綱尚未核准"]
    if appr["fingerprint"] != fingerprint(deck / "02_大綱.md"):
        return ["大綱在核准後被修改，核准已失效"]
    stale = stale_pages(state)
    if stale and not allow_stale:
        return [f"核准後 wiki 頁被修改，須重新執行 check-slides：{'、'.join(stale)}"]
    return []


def slide_text(slide):
    """投影片上會出現的所有文字（不含圖貼路徑與自製圖路徑）。"""
    parts = [slide["title"]]
    for k, v in slide["fields"].items():
        if k in ("圖貼", "圖", "引用頁"):
            continue
        if isinstance(v, list):
            parts += v
        elif isinstance(v, dict):
            parts += [c for row in v["rows"] for c in row]
        elif v:
            parts.append(v)
    return "\n".join(parts)


def check_one_slide(deck, s, plan, pages, all_refs):
    errs = []
    f = s["fields"]
    layout = s["layout"]
    if not plan:
        return ["大綱中沒有這一頁"]
    if layout != plan["layout"]:
        errs.append(f"版型「{layout}」與大綱「{plan['layout']}」不一致")
    if s["title"] != plan["title"]:
        errs.append(f"標題「{s['title']}」與大綱「{plan['title']}」不一致")
    if layout not in LAYOUT_FIELDS:
        return errs + [f"未知版型「{layout}」"]
    req, opt = LAYOUT_FIELDS[layout]
    if plan["message"] not in ("", "—"):
        req = req | {"主訊息"}
        if f.get("主訊息") != plan["message"]:
            errs.append("主訊息與大綱不一致")
    allowed = req | opt | {"主訊息", "備註"}
    for k in req:
        if not f.get(k):
            errs.append(f"缺少欄位「{k}」")
    for k in f:
        if k not in allowed:
            errs.append(f"「{layout}」版型不使用欄位「{k}」")
    for k, v in f.items():
        if v is None:
            errs.append(f"欄位「{k}」是空的")
        elif k in LIST_FIELDS:
            if not isinstance(v, list):
                errs.append(f"欄位「{k}」應為清單（每項縮排兩格 `- `）")
                continue
            if len(v) > LIST_MAX.get(k, 5):
                errs.append(f"「{k}」{len(v)} 項，超過 {LIST_MAX.get(k, 5)} 項")
            if k != "圖貼":
                errs += [f"「{k}」有一項 {len(x)} 字，超過 {ITEM_MAX} 字：{x[:12]}…" for x in v if len(x) > ITEM_MAX]
        elif k in TABLE_FIELDS:
            if not isinstance(v, dict):
                errs.append(f"欄位「{k}」應為表格")
                continue
            rows = v["rows"]
            if len(rows) - 1 > TABLE_ROWS_MAX:
                errs.append(f"「{k}」{len(rows) - 1} 列，超過 {TABLE_ROWS_MAX} 列")
            if len({len(r) for r in rows}) > 1:
                errs.append(f"「{k}」各列欄數不一致")
            errs += [f"「{k}」有一格 {len(c)} 字，超過 {CELL_MAX} 字：{c[:12]}…" for r in rows for c in r if len(c) > CELL_MAX]
        elif k not in ("備註", "引用頁") and isinstance(v, str) and len(v) > TEXT_MAX.get(k, TEXT_DEFAULT_MAX):
            errs.append(f"「{k}」{len(v)} 字，超過 {TEXT_MAX.get(k, TEXT_DEFAULT_MAX)} 字")
    # 引用頁：只能用大綱這一頁列出的頁
    refs = re.findall(r"\[\[([^\]]+)\]\]", f.get("引用頁") or "")
    extra = [r for r in refs if r not in plan["refs"]]
    if extra:
        errs.append(f"引用了大綱這一頁沒有的 wiki 頁：{'、'.join(extra)}")
    # 圖貼與圖
    for a in f.get("圖貼") or []:
        if not isinstance(f.get("圖貼"), list):
            break
        if not (ROOT / a).is_file():
            errs.append(f"圖貼不存在：{a}")
        elif a.startswith(ASSET_BANNED) or (a.startswith("assets/diagrams/") and a not in ASSET_DIAGRAM_OK):
            errs.append(f"圖貼含佔位文字或不適合直接使用：{a}")
    if f.get("圖") and not (deck / f["圖"]).is_file():
        errs.append(f"圖不存在：{f['圖']}")
    # 數字：每個數字都要出現在引用頁（無引用頁的版型，對照整份大綱的引用頁）
    src = [pages[resolve(r, pages)] for r in (refs or all_refs) if resolve(r, pages)]
    known = set()
    for p in src:
        known |= numbers(p.read_text(encoding="utf-8"), keep_raw=True)
    missing = sorted(numbers(slide_text(s)) - known, key=lambda x: (len(x), x))
    if missing:
        errs.append(f"這些數字在引用頁中找不到：{'、'.join(missing)}")
    return errs


def cmd_check_slides(deck):
    state = load_state(deck)
    problems = approval_problems(deck, state, allow_stale=True)
    stale = stale_pages(state)
    path = deck / "03_逐頁.md"
    if not path.exists():
        problems.append("找不到 03_逐頁.md")
    if problems:
        return fail(problems)
    _, slides = parse_slides(path.read_text(encoding="utf-8"))
    plan = outline_by_page(deck)
    pages = wiki_pages()
    all_refs = sorted({r for p in plan.values() for r in p["refs"]})
    per_slide, errs_global = {}, []
    nos = [s["no"] for s in slides]
    if nos != sorted(plan):
        errs_global.append(f"頁碼 {nos} 與大綱 {sorted(plan)} 不一致（不得新增、刪除或調換頁）")
    for s in slides:
        e = check_one_slide(deck, s, plan.get(s["no"]), pages, all_refs)
        if e:
            per_slide[s["no"]] = e
    ok = not per_slide and not errs_global
    retries = state.setdefault("retries", {}).setdefault("slides", {})
    for no in list(retries):
        if int(no) not in per_slide:
            retries[no] = 0
    for no in per_slide:
        retries[str(no)] = retries.get(str(no), 0) + 1
    state["slides_check"] = {"ok": ok, "fingerprint": fingerprint(path), "date": roc_today(),
                             "failing": sorted(per_slide)}
    if ok and stale:      # 引用頁被修改但逐頁內容仍通過：更新記錄的引用頁版本（大綱未改，不需重新核准）
        for r in stale:
            if r in pages:
                state["outline_approved"]["pages"][r] = fingerprint(pages[r])
    save_state(deck, state)
    lines = [f"- 逐頁內容指紋 {fingerprint(path)}，{len(slides)} 頁，結果：{'通過' if ok else '未通過'}"]
    if stale:
        lines.append(f"- 引用頁在核准後被修改，本次已重新檢查{'並更新版本記錄' if ok else ''}：{'、'.join(stale)}")
    lines += [f"- 錯誤：{e}" for e in errs_global]
    for no, es in sorted(per_slide.items()):
        lines += [f"- 第{no}頁：{e}" for e in es]
    report(deck, "check-slides", lines)
    print("\n".join(lines))
    stuck = [no for no, n in retries.items() if n >= MAX_RETRY]
    if stuck:
        print(f"第{'、'.join(stuck)}頁已連續 {MAX_RETRY} 次未通過，依規則停止，請向使用者回報。")
    return 0 if ok else 1


# ---------- 節點：render ----------

import shutil
import subprocess
import tempfile

EDGE = pathlib.Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
PPTX_SCRIPT = ROOT / "tools" / "deck_pptx.ps1"


def recycle(path):
    """移到資源回收筒（可還原），不直接刪除。"""
    import os
    cmd = ("Add-Type -AssemblyName Microsoft.VisualBasic; "
           "[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($env:DECK_RECYCLE, 'OnlyErrorDialogs', 'SendToRecycleBin')")
    subprocess.run(["powershell", "-NoProfile", "-Command", cmd], check=True,
                   env={**os.environ, "DECK_RECYCLE": str(path)})


def slides_ready(deck, state):
    problems = approval_problems(deck, state)
    chk = state.get("slides_check")
    path = deck / "03_逐頁.md"
    if not chk or not chk["ok"]:
        problems.append("逐頁內容尚未通過 check-slides")
    elif not path.exists() or chk["fingerprint"] != fingerprint(path):
        problems.append("03_逐頁.md 在檢查後被修改，請重新執行 check-slides")
    return problems


def is_draft(slides, pages):
    refs = {resolve(r, pages) for s in slides for r in re.findall(r"\[\[([^\]]+)\]\]", s["fields"].get("引用頁") or "")}
    refs.discard(None)
    return any(usability(page_info(pages[r])) != "可用（定稿）" for r in refs), sorted(refs)


def cmd_render(deck):
    import deck_render
    state = load_state(deck)
    problems = slides_ready(deck, state)
    if problems:
        return fail(problems)
    need, _ = read_need(deck)
    _, slides = parse_slides((deck / "03_逐頁.md").read_text(encoding="utf-8"))
    pages = wiki_pages()
    draft, refs = is_draft(slides, pages)
    topic = need["主題"]
    ctx = {"topic": topic, "level": need["層級"], "draft": draft, "refs": refs, "root": ROOT, "deck": deck,
           "date_long": deck_render.roc_long(roc_today())}
    out_html, out_pptx = deck / f"{topic}.html", deck / f"{topic}.pptx"
    prev = state.get("render", {})
    recycled = []
    for p, key in ((out_html, "html"), (out_pptx, "pptx")):
        if p.exists() and prev.get(key) != fingerprint(p):
            recycle(p)                      # 不是本流程產生的舊檔（或被手改過）：先移到資源回收筒
            recycled.append(p.name)
    fmts = need["格式"] if isinstance(need["格式"], list) else [need["格式"]]
    check_dir = deck / "_check"
    check_dir.mkdir(exist_ok=True)
    result = {"date": roc_today(), "slides": state["slides_check"]["fingerprint"], "draft": draft}
    if "html" in fmts:
        out_html.write_text(deck_render.build_html(slides, ctx), encoding="utf-8")
        result["html"] = fingerprint(out_html)
    if "pptx" in fmts:
        spec_path, rep_path = check_dir / "pptx-spec.json", check_dir / "pptx-report.json"
        spec_path.write_text(json.dumps(deck_render.build_pptx_spec(slides, ctx), ensure_ascii=False), encoding="utf-8")
        run = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(PPTX_SCRIPT),
                              "-Spec", str(spec_path), "-Out", str(out_pptx), "-Report", str(rep_path),
                              "-PngDir", str(check_dir)], capture_output=True, text=True, encoding="utf-8", errors="replace")
        m = re.search(r"OK slides=(\d+)", run.stdout)
        if run.returncode != 0 or not out_pptx.exists() or not m:
            return fail([f"PowerPoint 產生失敗：{(run.stdout + run.stderr).strip()[-400:]}"])
        if int(m.group(1)) != len(slides):
            return fail([f"PowerPoint 只畫了 {m.group(1)} 張，應為 {len(slides)} 張"])
        pngs = len(list(check_dir.glob("pptx-*.png")))
        if pngs != len(slides):
            return fail([f"PowerPoint 匯出圖片 {pngs} 張，應為 {len(slides)} 張"])
        result["pptx"] = fingerprint(out_pptx)
    state["render"] = result
    state.pop("render_check", None)
    save_state(deck, state)
    lines = [f"- 由逐頁內容 {result['slides']} 產生：{'、'.join(p.name for p in (out_html, out_pptx) if p.exists())}；狀態：{'草稿' if draft else '定稿'}"]
    if recycled:
        lines.append(f"- 舊檔已移到資源回收筒：{'、'.join(recycled)}")
    report(deck, "render", lines)
    print("\n".join(lines))
    return 0


# ---------- 節點：check-render ----------

def edge(args, timeout=120):
    prof = tempfile.mkdtemp(prefix="deck-edge-")
    try:
        return subprocess.run([str(EDGE), "--headless=new", "--disable-gpu", "--no-first-run", "--hide-scrollbars",
                               f"--user-data-dir={prof}", "--window-size=1280,720", *args],
                              capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout)
    finally:
        shutil.rmtree(prof, ignore_errors=True)


def cmd_check_render(deck):
    import html as htmllib
    state = load_state(deck)
    need, _ = read_need(deck)
    topic = need["主題"]
    rend = state.get("render")
    out_html, out_pptx = deck / f"{topic}.html", deck / f"{topic}.pptx"
    problems = slides_ready(deck, state)
    if not rend or rend.get("slides") != state.get("slides_check", {}).get("fingerprint"):
        problems.append("成品不是由目前的逐頁內容產生，請重新 render")
    for p, key in ((out_html, "html"), (out_pptx, "pptx")):
        if key in rend if rend else False:
            if not p.exists() or fingerprint(p) != rend[key]:
                problems.append(f"{p.name} 在 render 後被修改或不存在，請重新 render")
    if problems:
        return fail(problems)
    check_dir = deck / "_check"
    per, warns = {}, []
    if "html" in rend:
        if not EDGE.exists():
            return fail([f"找不到 Edge（{EDGE}），無法檢查 HTML"])
        url = out_html.as_uri() + "?check"
        run = edge(["--virtual-time-budget=20000", "--dump-dom", url])
        m = re.search(r'<pre id="deck-check">(.*?)</pre>', run.stdout, re.S)
        if not m or not m.group(1).strip():
            return fail(["HTML 檢查沒有產生結果（Edge 無法執行頁面上的檢查程式）"])
        res = json.loads(htmllib.unescape(m.group(1)))
        if not res.get("fontLoaded"):
            warns.append("檢查時 Noto Sans TC 未載入（可能沒有網路），HTML 以備用字型檢查")
        for s in res["slides"]:
            for i in s["issues"]:
                per.setdefault(s["no"], []).append("HTML：" + i)
        for old in check_dir.glob("html-*.png"):
            old.unlink()
        n = len(res["slides"])
        for i in range(1, n + 1):
            edge(["--virtual-time-budget=8000", f"--screenshot={check_dir / f'html-{i:02d}.png'}", out_html.as_uri() + f"?shot#{i}"])
        shots = len(list(check_dir.glob("html-*.png")))
        if shots != n:
            per.setdefault(0, []).append(f"HTML 截圖只產生 {shots}/{n} 張")
    if "pptx" in rend:
        rep = check_dir / "pptx-report.json"
        items = json.loads(rep.read_text(encoding="utf-8-sig") or "[]") if rep.exists() else None
        if items is None:
            per.setdefault(0, []).append("找不到 PowerPoint 檢查結果，請重新 render")
        else:
            for it in items if isinstance(items, list) else [items]:
                per.setdefault(it["no"], []).append("PowerPoint：" + it["issue"])
    ok = not per
    retries = state.setdefault("retries", {})
    retries["render"] = 0 if ok else retries.get("render", 0) + 1
    state["render_check"] = {"ok": ok, "date": roc_today(), "html": rend.get("html"), "pptx": rend.get("pptx")}
    save_state(deck, state)
    lines = [f"- 結果：{'通過' if ok else '未通過'}；截圖與 PowerPoint 匯出圖在 _check/"]
    lines += [f"- 提醒：{w}" for w in warns]
    for no, es in sorted(per.items()):
        lines += [f"- 第{no}頁：{e}" for e in es]
    report(deck, "check-render", lines)
    print("\n".join(lines))
    if not ok and retries["render"] >= MAX_RETRY:
        print(f"成品檢查已連續 {retries['render']} 次未通過，依規則停止，請向使用者回報。")
    return 0 if ok else 1


# ---------- 節點：finalize ----------

def cmd_finalize(deck):
    state = load_state(deck)
    rc, rend = state.get("render_check"), state.get("render")
    need, _ = read_need(deck)
    topic = need["主題"]
    problems = slides_ready(deck, state)
    if not rc or not rc["ok"]:
        problems.append("成品尚未通過 check-render")
    else:
        for key, p in (("html", deck / f"{topic}.html"), ("pptx", deck / f"{topic}.pptx")):
            if rc.get(key) and (not p.exists() or fingerprint(p) != rc[key]):
                problems.append(f"{p.name} 在檢查後被修改，請重新 render 與 check-render")
    if problems:
        return fail(problems)
    _, slides = parse_slides((deck / "03_逐頁.md").read_text(encoding="utf-8"))
    pages = wiki_pages()
    draft, refs = is_draft(slides, pages)
    if draft != rend["draft"]:
        return fail(["引用頁的審閱狀態在 render 後有變動，請重新 render（草稿／定稿標示會改變）"])
    unreviewed = [r for r in refs if usability(page_info(pages[r])) != "可用（定稿）"]
    state["final"] = {"date": roc_today(), "status": "草稿" if draft else "定稿", "html": rc.get("html"), "pptx": rc.get("pptx")}
    save_state(deck, state)
    lines = [f"- 狀態：{'草稿' if draft else '定稿'}；成品：{topic}.html、{topic}.pptx"]
    if unreviewed:
        lines.append(f"- 定稿前須審閱的頁（{len(unreviewed)}）：{'、'.join(unreviewed)}")
    report(deck, "finalize", lines)
    print("\n".join(lines))
    print("請在 wiki/log.md 記一筆 slides 紀錄。")
    return 0


# ---------- ci：給 GitHub Actions 用的唯讀檢查 ----------

def cmd_ci(deck):
    """不改動簡報資料夾：在暫存複本上重跑 check-outline 與 check-slides，
    並確認成品是由目前的逐頁內容產生、草稿／定稿標示與目前審閱狀態一致。
    不需要 PowerPoint 或 Edge，可在 Linux 上執行。"""
    import io
    import contextlib
    state = load_state(deck)
    need, errs = read_need(deck)
    if errs:
        return fail(errs)
    topic = need["主題"]
    if not state.get("outline_approved"):
        print(f"{deck.name}：大綱尚未核准，流程進行中，略過成品檢查")
    stale = stale_pages(state)
    if stale:
        errs.append(f"引用的 wiki 頁在上次檢查後被修改，請在本機重跑 check-slides、render、check-render、finalize：{'、'.join(stale)}")
    tmp = pathlib.Path(tempfile.mkdtemp(prefix="deck-ci-"))
    try:
        copy = tmp / deck.name
        shutil.copytree(deck, copy, ignore=shutil.ignore_patterns("_check"))
        for name, fn in (("check-outline", cmd_check_outline), ("check-slides", cmd_check_slides)):
            if name == "check-slides" and not (copy / "03_逐頁.md").exists():
                continue
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf):
                rc = fn(copy)
            if rc != 0:
                errs.append(f"{name} 未通過：\n" + buf.getvalue().strip())
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    rend, fin = state.get("render"), state.get("final")
    sp = deck / "03_逐頁.md"
    if rend and sp.exists():
        if rend.get("slides") != fingerprint(sp):
            errs.append("03_逐頁.md 在 render 之後被修改，成品已過期，請在本機重跑 render、check-render、finalize")
        for key in ("html", "pptx"):
            p = deck / f"{topic}.{key}"
            if rend.get(key) and (not p.exists() or fingerprint(p) != rend[key]):
                errs.append(f"{p.name} 與 render 紀錄不符（被手改、遺失或過期），請在本機重新 render")
    if fin and sp.exists():
        _, slides = parse_slides(sp.read_text(encoding="utf-8"))
        draft, _ = is_draft(slides, wiki_pages())
        now = "草稿" if draft else "定稿"
        if fin.get("status") != now:
            errs.append(f"簡報標為「{fin.get('status')}」，但依目前的審閱狀態應為「{now}」，請在本機重跑 render、check-render、finalize")
    if errs:
        print(f"✗ {deck.name}")
        for e in errs:
            print(f"錯誤：{e}")
        return 1
    print(f"✓ {deck.name}：內容檢查通過，成品與紀錄一致（{fin.get('status') if fin else '尚未收尾'}）")
    return 0


# ---------- 節點：status ----------

def cmd_status(deck):
    state = load_state(deck)
    pages = wiki_pages()
    steps = []

    def step(name, ok, note=""):
        steps.append((name, ok, note))

    need_ok = (deck / "00_需求.md").exists()
    step("00 需求", need_ok, "" if need_ok else "請建立 00_需求.md")
    inv = state.get("inventory")
    inv_ok = bool(inv) and need_ok and inv["need"] == fingerprint(deck / "00_需求.md")
    step("01 素材盤點", inv_ok, "" if inv_ok else "執行 inventory")
    outline = deck / "02_大綱.md"
    step("02 撰寫大綱", outline.exists(), "" if outline.exists() else "agent 撰寫 02_大綱.md")
    chk = state.get("outline_check")
    chk_ok = bool(chk) and outline.exists() and chk["ok"] and chk["fingerprint"] == fingerprint(outline)
    step("02 檢查大綱", chk_ok, "" if chk_ok else "執行 check-outline")
    appr = state.get("outline_approved")
    appr_ok, note = False, "等待使用者在對話中核准"
    if appr and outline.exists():
        if appr["fingerprint"] != fingerprint(outline):
            note = "大綱在核准後被修改，核准已失效"
        else:
            stale = [r for r, fp in appr["pages"].items() if r not in pages or fingerprint(pages[r]) != fp]
            appr_ok = not stale
            note = "" if appr_ok else "核准後這些 wiki 頁被修改，須重新執行 check-slides：" + "、".join(stale)
    step("02 使用者核准大綱", appr_ok, note)
    sp = deck / "03_逐頁.md"
    step("03 撰寫逐頁內容", sp.exists(), "" if sp.exists() else "agent 撰寫 03_逐頁.md")
    sc = state.get("slides_check")
    sc_ok = bool(sc) and sp.exists() and sc["ok"] and sc["fingerprint"] == fingerprint(sp)
    step("03 檢查逐頁（check-slides）", sc_ok, "" if sc_ok else "執行 check-slides")
    rend = state.get("render")
    r_ok = bool(rend) and sc_ok and rend.get("slides") == sc["fingerprint"]
    step("05 產出（render）", r_ok, "" if r_ok else "執行 render")
    rc = state.get("render_check")
    rc_ok = bool(rc) and r_ok and rc["ok"] and rc.get("html") == rend.get("html") and rc.get("pptx") == rend.get("pptx")
    step("05 檢查成品（check-render）", rc_ok, "" if rc_ok else "執行 check-render")
    fin = state.get("final")
    f_ok = bool(fin) and rc_ok and fin.get("html") == rc.get("html") and fin.get("pptx") == rc.get("pptx")
    step("06 收尾（finalize）", f_ok, (fin["status"] if f_ok else "執行 finalize"))

    for name, ok, note in steps:
        print(f"{'✓' if ok else '·'} {name}" + (f" — {note}" if note else ""))
    nxt = next((s for s in steps if not s[1]), None)
    print(f"\n下一步：{nxt[0]}" + (f"（{nxt[2]}）" if nxt and nxt[2] else "") if nxt else "\n全部完成")
    return 0


def not_yet(_deck):
    return fail(["此節點在大綱核准後才建立，目前尚未實作"])


def fail(errs):
    for e in errs:
        print(f"錯誤：{e}")
    return 1


COMMANDS = {
    "status": cmd_status,
    "inventory": cmd_inventory,
    "check-outline": cmd_check_outline,
    "approve-outline": cmd_approve_outline,
    "check-slides": cmd_check_slides,
    "render": cmd_render,
    "check-render": cmd_check_render,
    "finalize": cmd_finalize,
    "ci": cmd_ci,
}

if __name__ == "__main__":
    sys.path.insert(0, str(ROOT / "tools"))
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) != 3 or sys.argv[1] not in COMMANDS:
        print(__doc__)
        sys.exit(2)
    deck = (ROOT / sys.argv[2]) if not pathlib.Path(sys.argv[2]).is_absolute() else pathlib.Path(sys.argv[2])
    if not deck.is_dir():
        print(f"錯誤：找不到資料夾 {deck}")
        sys.exit(2)
    sys.exit(COMMANDS[sys.argv[1]](deck))
