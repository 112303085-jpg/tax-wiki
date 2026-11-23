"""Wiki helper: print or embed verbatim article text from raw/ law files.

Usage:
  python tools/law.py show <raw檔> <條號>...        印出條文原文，例：python tools/law.py show raw/A_法規/1150911_所得稅法.md 17 14-4
  python tools/law.py expand <wiki頁>...             將頁面中的 <!-- law: <raw檔> | <條號> --> 標記展開成條文原文 callout
  python tools/law.py check  <wiki頁>...             檢查頁面中已展開的條文是否仍與 raw 原文逐字一致

展開後的格式（標記保留，可重複執行；check 會比對其內容）：
<!-- law: raw/A_法規/1150911_所得稅法.md | 14-4 -->
> [!quote] 法條原文｜所得稅法第14條之4（版本 115-09-11）
> ...
<!-- /law -->
"""
import re, sys, pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent


def load(raw_rel):
    text = (ROOT / raw_rel).read_text(encoding="utf-8")
    meta = dict(re.findall(r"^(\S+?):\s*(.*)$", text.split("---", 2)[1], re.M))
    arts, cur = {}, None
    for line in text.split("---", 2)[2].splitlines():
        m = re.match(r"^#### 第 (\S+) 條\s*$", line)
        if m:
            cur = m.group(1); arts[cur] = []; continue
        if line.startswith("#"):
            cur = None; continue
        if cur is not None:
            arts[cur].append(line.rstrip())
    return meta, {k: "\n".join(v).strip() for k, v in arts.items()}


def label(no):
    a, _, b = no.partition("-")
    return f"第{a}條" + (f"之{b}" if b else "")


def callout(raw_rel, no):
    meta, arts = load(raw_rel)
    if no not in arts:
        sys.exit(f"找不到條號 {no}：{raw_rel}")
    body = [l for l in arts[no].splitlines() if l.strip()]
    head = f"> [!quote] 法條原文｜{meta.get('標題','')}{label(no)}（版本 {meta.get('版本/修正日','')}）"
    return "\n".join([head] + ["> " + l.rstrip() for l in body])


def raw_corpus():
    return "".join(re.sub(r"\s", "", f.read_text(encoding="utf-8")) for f in (ROOT / "raw").rglob("*.md"))


def check_info_quotes(path, s, corpus):
    """[!info] 官方解釋、[!quote] 法規沿革 callout 的每一行都必須逐字出現在 raw/ 中（標註「要旨」者除外）。"""
    bad = 0
    for block in re.findall(r"^> \[!(?:info\]|quote\] 法規沿革)[^\n]*\n((?:>[^\n]*\n?)*)", s, re.M):
        if "要旨" in block: continue
        for line in block.splitlines():
            t = re.sub(r"\s", "", line.lstrip(">"))
            if t and t not in corpus:
                bad += 1; print(f"引文找不到原文（[!info]／法規沿革）：{path}：{line[:60]}")
    return bad


PAT = re.compile(r"<!-- law: (.+?) \| (\S+?) -->\n(?:.*?\n)?<!-- /law -->|<!-- law: (.+?) \| (\S+?) -->(?!\n>)", re.S)


def expand(path):
    p = pathlib.Path(path); s = p.read_text(encoding="utf-8")
    def rep(m):
        raw_rel = (m.group(1) or m.group(3)).strip(); no = (m.group(2) or m.group(4)).strip()
        return f"<!-- law: {raw_rel} | {no} -->\n{callout(raw_rel, no)}\n<!-- /law -->"
    s2 = PAT.sub(rep, s)
    p.write_text(s2, encoding="utf-8")
    print(f"{path}: {len(PAT.findall(s))} 處")


def check(path, corpus=None):
    s = pathlib.Path(path).read_text(encoding="utf-8"); bad = 0
    for m in re.finditer(r"<!-- law: (.+?) \| (\S+?) -->\n(.*?)\n<!-- /law -->", s, re.S):
        if m.group(3) != callout(m.group(1).strip(), m.group(2).strip()):
            bad += 1; print(f"不一致：{path} {m.group(2)}")
    if re.search(r"<!-- law: [^\n]+ -->(?!\n>)", s):
        bad += 1; print(f"{path}: 有尚未展開的條文標記")
    bad += check_info_quotes(path, s, corpus if corpus is not None else raw_corpus())
    print(f"{path}: {'OK' if not bad else f'{bad} 處不一致'}")
    return bad


if __name__ == "__main__":
    cmd, *args = sys.argv[1:]
    if cmd == "show":
        meta, arts = load(args[0])
        for no in args[1:]:
            print(f"=== {label(no)} ===\n{arts.get(no, '（找不到）')}\n")
    elif cmd == "expand":
        for a in args: expand(a)
    elif cmd == "check":
        corpus = raw_corpus()
        sys.exit(1 if sum(check(a, corpus) for a in args) else 0)
