"""產生 GitHub Pages 網站（AGENTS.md 第 9 節）。

只公開「定稿」的簡報：slides/<主題>/state.json 的 final.status 為「定稿」，且 deck.py ci 通過。
草稿一律不放上網站（AGENTS.md 第 7.3 節：草稿不得對外使用）。

用法：python tools/build_site.py <輸出資料夾>（預設 _site）
網站結構：
  index.html                     簡報目錄
  assets/                        通用素材（簡報 HTML 以 ../../assets/ 引用）
  slides/<主題>/<主題>.html       簡報
  slides/<主題>/<主題>.pptx       PowerPoint 下載
  slides/<主題>/images/          該簡報的自製圖
"""
import html
import json
import pathlib
import shutil
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import deck  # noqa: E402

DISCLAIMER = "本網站內容為教學整理，僅供學習參考，不構成個案稅務或法律意見。法規與數額可能修正，實際適用請以財政部及主管機關最新公告為準，或洽詢專業人士。"


def main(out):
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    shutil.copytree(ROOT / "assets", out / "assets",
                    ignore=shutil.ignore_patterns("_generator.py", "preview.html"))
    published, skipped = [], []
    for d in sorted((ROOT / "slides").iterdir()):
        state_file = d / "state.json"
        if not d.is_dir() or not state_file.exists():
            continue
        state = json.loads(state_file.read_text(encoding="utf-8-sig"))
        need, _ = deck.read_need(d)
        topic = need["主題"]
        final = state.get("final") or {}
        if final.get("status") != "定稿":
            skipped.append(f"{d.name}（{final.get('status', '尚未收尾')}）")
            continue
        if deck.cmd_ci(d) != 0:
            sys.exit(f"{d.name} 未通過 deck.py ci，停止產生網站")
        dst = out / "slides" / d.name
        dst.mkdir(parents=True)
        for name in (f"{topic}.html", f"{topic}.pptx"):
            if (d / name).exists():
                shutil.copy2(d / name, dst / name)
        if (d / "images").is_dir():
            shutil.copytree(d / "images", dst / "images")
        published.append({
            "topic": topic, "dir": d.name, "level": need.get("層級", ""),
            "minutes": need.get("時長分鐘", ""), "date": final.get("date", ""),
            "pptx": (d / f"{topic}.pptx").exists(),
        })
    (out / "index.html").write_text(index_page(published), encoding="utf-8")
    (out / ".nojekyll").write_text("", encoding="utf-8")
    print(f"網站已產生：{out}；公開 {len(published)} 份" + (f"；略過（非定稿）：{'、'.join(skipped)}" if skipped else ""))


def index_page(items):
    rows = []
    for it in items:
        base = f"slides/{it['dir']}/{it['topic']}"
        pptx = f'<a class="ghost" href="{html.escape(base)}.pptx">下載 PowerPoint</a>' if it["pptx"] else ""
        rows.append(f"""<li class="card">
  <div class="meta">{html.escape(str(it['level']))}場｜約 {html.escape(str(it['minutes']))} 分鐘｜定稿日 {html.escape(str(it['date']))}</div>
  <h2>{html.escape(it['topic'])}</h2>
  <div class="actions"><a class="main" href="{html.escape(base)}.html">播放簡報</a>{pptx}</div>
</li>""")
    body = "\n".join(rows) or '<li class="card"><h2>目前沒有定稿的簡報</h2></li>'
    return f"""<!doctype html>
<html lang="zh-Hant-TW">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>台灣稅務教學簡報</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;700;900&display=swap" rel="stylesheet">
<style>
:root{{--ink:#111;--paper:#fff;--muted:#555;--accent:#7B61FF;--orange:#FF8A3D}}
*{{box-sizing:border-box}}
body{{margin:0;background:var(--paper);color:var(--ink);font-family:'Noto Sans TC','Microsoft JhengHei',sans-serif;line-height:1.6}}
main{{max-width:880px;margin:0 auto;padding:48px max(16px,env(safe-area-inset-left,0px)) 64px}}
h1{{font-size:clamp(34px,6vw,52px);font-weight:900;margin:0 0 6px;display:inline-block;transform:skewX(-6deg);text-shadow:5px 5px 0 var(--accent)}}
.lead{{color:var(--muted);font-weight:700;margin:0 0 36px}}
ul{{list-style:none;padding:0;margin:0;display:grid;gap:28px}}
.card{{border:4px solid var(--ink);box-shadow:8px 8px 0 var(--ink);padding:22px 26px}}
.card h2{{margin:4px 0 16px;font-size:28px;font-weight:900}}
.meta{{color:var(--muted);font-size:15px;font-variant-numeric:tabular-nums}}
.actions{{display:flex;flex-wrap:wrap;gap:14px}}
.actions a{{display:inline-block;padding:8px 18px;font-weight:900;text-decoration:none;border:3px solid var(--ink);color:var(--ink)}}
.actions a.main{{background:var(--orange);box-shadow:4px 4px 0 var(--ink)}}
.actions a:focus-visible{{outline:3px solid var(--accent);outline-offset:3px}}
footer{{margin-top:48px;color:var(--muted);font-size:14px}}
</style>
</head>
<body>
<main>
<h1>台灣稅務教學簡報</h1>
<p class="lead">綜合所得稅與房地合一，依法規原文與財政部公告整理，數字皆經程式驗算。</p>
<ul>
{body}
</ul>
<footer>{html.escape(DISCLAIMER)}</footer>
</main>
</body>
</html>
"""


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main(ROOT / (sys.argv[1] if len(sys.argv) > 1 else "_site"))
