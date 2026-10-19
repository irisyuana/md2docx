#!/usr/bin/env python3
"""机械类合规修复:H1 降级 / EMOJI 替换 / LANG 补语言标识。
用法:
  python3 tools/md_fix_mech.py --dry   # 预览
  python3 tools/md_fix_mech.py --apply # 应用
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
from md_audit import audit, list_mds, MD格式约束_EXEMPT  # noqa: E402

DRY = "--apply" not in sys.argv


def fix_h1(lines):
    """第一个 # 在第 1 行则保留为封面,其余标题层级 +1(跳过代码块)。"""
    h1_idx = [i for i, l in enumerate(lines) if re.match(r"^ {0,3}# [^#]", l)]
    if len(h1_idx) < 2 or h1_idx[0] != 0:
        return lines, 0
    out, in_f, n = [], False, 0
    for i, l in enumerate(lines):
        if re.match(r"^ {0,3}(`{3,}|~{3,})", l):
            in_f = not in_f
            out.append(l)
            continue
        if i > 0 and not in_f and re.match(r"^ {0,3}#{1,6}\s", l):
            l = "#" + l
            n += 1
        out.append(l)
    return out, n


def fix_emoji_line(line):
    if not any(e in line for e in "✅❌⚠"):
        return line
    # 独立单元格先替换(| ✅ | / | ❌ |)
    line = re.sub(r"(?<=\|)\s*✅\s*(?=\|)", " 是 ", line)
    line = re.sub(r"(?<=\|)\s*❌\s*(?=\|)", " 否 ", line)
    # ✅/⚠️+文字 → 去掉 Emoji 保留文字(lookahead 排除单元格竖线)
    line = re.sub(r"✅\s*(?=[^\s|])", "", line)
    line = re.sub(r"⚠️?\s*(?=[^\s|])", "", line)
    # ❌ (说明) → 否（说明）;❌+文字 → 否 文字
    line = re.sub(r"❌\s*\(([^)]*)\)", r"否（\1）", line)
    line = re.sub(r"❌\s*(?=[^\s|])", "否 ", line)
    # 独立单元格兜底
    line = re.sub(r"(?<=\|)\s*✅\s*(?=\|)", " 是 ", line)
    line = re.sub(r"(?<=\|)\s*❌\s*(?=\|)", " 否 ", line)
    # 兜底
    return line.replace("✅", "是").replace("❌", "否").replace("⚠️", "").replace("⚠", "")


def fix_lang(lines):
    """裸 ``` 开启的代码块补语言标识(目录树/纯文本 → text)。"""
    out, in_f, n = [], False, 0
    for l in lines:
        m = re.match(r"^ {0,3}(`{3,})(\S*)\s*$", l)
        if m:
            if not in_f:
                in_f = True
                if not m.group(2):
                    l = m.group(1) + "text"
                    n += 1
            else:
                in_f = False
        out.append(l)
    return out, n


def main():
    files = [f for f in list_mds([]) if f not in MD格式约束_EXEMPT]
    stats = {"H1": 0, "EMOJI": 0, "LANG": 0, "files": 0}
    for f in files:
        p = ROOT / f
        lines = p.read_text(encoding="utf-8").splitlines()
        issues = {c for _, c, _ in audit(p)}
        if not (issues & {"H1", "EMOJI", "LANG"}):
            continue
        orig = list(lines)
        log = []
        if "H1" in issues:
            lines, n = fix_h1(lines)
            if n:
                stats["H1"] += n
                log.append(f"H1 降级 {n} 个标题")
        if "EMOJI" in issues:
            n = sum(1 for a, b in zip(lines, map(fix_emoji_line, lines)) if a != b)
            lines = [fix_emoji_line(l) for l in lines]
            if n:
                stats["EMOJI"] += n
                log.append(f"EMOJI 清理 {n} 行")
        if "LANG" in issues:
            new, n = fix_lang(lines)
            n2 = sum(1 for a, b in zip(lines, new) if a != b)
            lines = new
            if n2:
                stats["LANG"] += n2
                log.append(f"LANG 补标识 {n2} 个")
        if lines == orig:
            continue
        stats["files"] += 1
        print(f"{f}: {'; '.join(log)}")
        if not DRY:
            p.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"\n[{'APPLY' if not DRY else 'DRY'}] 修改 {stats['files']} 文件:"
          f" H1={stats['H1']} EMOJI={stats['EMOJI']} LANG={stats['LANG']}")


if __name__ == "__main__":
    main()
