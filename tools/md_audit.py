#!/usr/bin/env python3
"""按 MD格式约束.md 对项目全部 md 做合规扫描。
用法: python3 tools/md_audit.py [文件...]  # 缺省扫描 md2docx.sh 全量范围
"""
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# 与转换管线一致的范围：build/ 下的全部 Markdown。
FIND_CMD = [
    "find", "build", "-name", "*.md",
]
MD格式约束_EXEMPT = set()  # README/MD格式约束 在仓库根,不在 build/ 扫描范围,天然豁免


def list_mds(explicit):
    if explicit:
        return explicit
    out = subprocess.run(FIND_CMD, cwd=ROOT, capture_output=True, text=True).stdout
    return sorted(l for l in out.splitlines() if l.strip())


def parse_struct(lines):
    """返回 [(lineno, kind, text)];kind: blank/h1/hN/table/hr/text/fence_open/fence_close
    围栏闭合须用同字符且长度不小于开启长度(CommonMark 规则)。"""
    struct, in_fence, marker_re = [], False, None
    for i, ln in enumerate(lines, 1):
        if in_fence:
            if marker_re.match(ln):
                in_fence = False
                struct.append((i, "fence_close", ln))
            else:
                struct.append((i, "fence_line", ln))
            continue
        m = re.match(r"^ {0,3}(`{3,}|~{3,})", ln)
        if m:
            in_fence = True
            tok = m.group(1)
            marker_re = re.compile(r"^ {0,3}%s{%d,}\s*$" % (re.escape(tok[0]), len(tok)))
            struct.append((i, "fence_open", ln))
            continue
        if not ln.strip():
            k = "blank"
        elif re.match(r"^# [^#]", ln):
            k = "h1"
        elif re.match(r"^#+\s", ln):
            k = "h%d" % (len(ln) - len(ln.lstrip("#")))
        elif ln.lstrip().startswith("|"):
            k = "table"
        elif re.match(r"^-{3,}$", ln.strip()):
            k = "hr"
        else:
            k = "text"
        struct.append((i, k, ln))
    return struct


def audit(path):
    lines = Path(path).read_text(encoding="utf-8").splitlines()
    issues = []

    def add(lineno, code, msg):
        issues.append((lineno, code, msg))

    struct = parse_struct(lines)
    raw = lines

    # 1) 唯一一级标题开头
    h1s = [(i, t) for i, k, t in struct if k == "h1"]
    if len(h1s) != 1 or (h1s and h1s[0][0] != 1):
        add(h1s[0][0] if h1s else 1, "H1", f"一级标题数量={len(h1s)} 应为文档第 1 行")

    # 2) 标题跳级
    prev = 1
    for i, k, t in struct:
        if k.startswith("h") and k[1:].isdigit():
            lv = int(k[1:])
            if lv > prev + 1:
                add(i, "SKIP", f"标题跳级 {prev}->{lv}: {t.strip()[:40]}")
            prev = lv

    # 3) 表格块表题
    blocks, idx = [], 0
    while idx < len(struct):
        if struct[idx][1] == "table":
            s = idx
            while idx < len(struct) and struct[idx][1] == "table":
                idx += 1
            blocks.append((s, idx - 1))
        else:
            idx += 1
    bold_re = re.compile(r"^\*\*[^*]+\*\*$")
    for bi, (s, e) in enumerate(blocks):
        ln = struct[s][0]
        j = s - 1
        while j >= 0 and struct[j][1] == "blank":
            j -= 1
        pk, pt = (struct[j][1], struct[j][2]) if j >= 0 else ("", "")
        if bi == 0 and pk == "h1":
            continue  # 封面信息表
        if not (pk == "text" and bold_re.match(pt.strip())):
            add(ln, "CAP", f"表格缺加粗表题 (前文[{pk}]: {pt.strip()[:40]})")
        elif re.search(r"[表图]\s*\d+[-–—]\d+", pt):
            add(ln, "CAPNUM", f"表题含手写编号: {pt.strip()[:40]}")

    # 4) 代码块语言标识 / mermaid fig-cap —— 正确配对(围栏允许 ≤3 空格缩进)
    in_fence = False
    for i, ln in enumerate(lines):
        if not re.match(r"^ {0,3}`{3,}", ln):
            continue
        if not in_fence:
            in_fence = True
            lang = re.match(r"^ {0,3}`{3,}(\S*)", ln).group(1)
            if not lang:
                add(i + 1, "LANG", "代码块缺语言标识")
            elif lang == "mermaid":
                j = i + 1
                while j < len(lines) and not lines[j].strip():
                    j += 1
                if j >= len(lines) or "fig-cap:" not in lines[j]:
                    add(i + 1, "FIGCAP", "mermaid 块缺 fig-cap")
        else:
            in_fence = False

    # 5) Emoji
    for i, k, t in struct:
        if k in ("text", "table"):
            for e in "⚠✅❌❗":
                if e in t:
                    add(i, "EMOJI", f"Emoji「{e}」: {t.strip()[:40]}")
                    break

    # 6) 正文、表格单元格与标题箭头
    for i, k, t in struct:
        if k in ("text", "table") or (k.startswith("h") and k[1:].isdigit()):
            if re.search(r"[→⇒➜]", t):
                add(i, "ARROW", f"正文箭头: {t.strip()[:40]}")

    # 7) 不可剥离的手写标题编号
    for i, k, t in struct:
        if k.startswith("h") and k[1:].isdigit():
            body = re.sub(r"^#+\s*", "", t)
            if re.match(r"^(第[一二三四五六七八九十]+章|[一二三四五六七八九十]+、|（[一二三四五六七八九十]+）|\d+、)", body):
                add(i, "NUMFMT", f"不可剥离编号: {body[:30]}")

    # 8) 手写图表编号(正文含 表 N-M / 图 N-M)
    for i, k, t in struct:
        if k == "text" and re.search(r"[表图]\s*\d+[-–—]\d+", t):
            add(i, "REFNUM", f"正文疑似手写图表编号: {t.strip()[:40]}")

    # 9) ASCII 框线字符(约束 5:禁止 ┌ ┐ └ ┘ │ ─ 等拼接框图,含代码块内)
    box_re = re.compile(r"[┌┐└┘├┤┬┴┼╔╗╚╝║═╠╣]")
    for i, ln in enumerate(lines, 1):
        if box_re.search(ln):
            add(i, "ASCII", f"ASCII 框线字符: {ln.strip()[:40]}")

    return issues


def main():
    explicit = sys.argv[1:]
    files = list_mds(explicit)
    total = 0
    for f in files:
        if f in MD格式约束_EXEMPT:
            continue
        issues = audit(ROOT / f)
        if issues:
            print(f"\n=== {f} ({len(issues)}) ===")
            for ln, code, msg in issues:
                print(f"  L{ln} [{code}] {msg}")
            total += len(issues)
    print(f"\n扫描 {len(files)} 个文件,共 {total} 个问题")


if __name__ == "__main__":
    main()
