"""Wiki helper: run every ```python example in wiki pages and write the real output below it.

Usage:
  python tools/calc.py run   <wiki頁>...   執行頁面中的 python 例題，將結果寫入其後的「執行結果」區塊
  python tools/calc.py check <wiki頁>...   重新執行並比對，若頁面上的結果與實際輸出不同則回報

頁面寫法（執行結果區塊由本工具產生，不要手寫）：
```python
print(1 + 1)
```
執行結果：
```text
2
```
"""
import io, re, sys, pathlib, contextlib

BLOCK = re.compile(r"```python\n(.*?)```\n(?:執行結果：\n```text\n.*?```\n)?", re.S)


def execute(code):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        exec(compile(code, "<wiki>", "exec"), {})
    return buf.getvalue().rstrip("\n")


def render(code):
    return f"```python\n{code}```\n執行結果：\n```text\n{execute(code)}\n```\n"


def run(path):
    p = pathlib.Path(path); s = p.read_text(encoding="utf-8")
    n = len(BLOCK.findall(s))
    p.write_text(BLOCK.sub(lambda m: render(m.group(1)), s), encoding="utf-8")
    print(f"{path}: {n} 個例題已執行")


def check(path):
    s = pathlib.Path(path).read_text(encoding="utf-8")
    bad = 0
    for m in re.finditer(r"```python\n(.*?)```\n執行結果：\n```text\n(.*?)\n```\n", s, re.S):
        if execute(m.group(1)) != m.group(2):
            bad += 1; print(f"結果不符：{path}\n{m.group(1)[:80]}")
    missing = len(BLOCK.findall(s)) - len(re.findall(r"```python\n.*?```\n執行結果：", s, re.S))
    if missing: bad += missing; print(f"{path}: {missing} 個例題尚未執行")
    print(f"{path}: {'OK' if not bad else f'{bad} 處問題'}")
    return bad


if __name__ == "__main__":
    cmd, *args = sys.argv[1:]
    if cmd == "run":
        for a in args: run(a)
    elif cmd == "check":
        sys.exit(1 if sum(check(a) for a in args) else 0)
