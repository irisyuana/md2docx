#!/usr/bin/env python3
"""批量插入表题:按 audit 定位 CAP 表格块,倒序插入 **表题** 行。
用法: python3 tools/md_insert_caps.py <jobs.json>
jobs.json: {"文件路径": ["表题1", "表题2", ...]}  # 按文档中出现顺序
"""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
from md_audit import audit, parse_struct  # noqa: E402


def insert_caps(rel, captions):
    p = ROOT / rel
    lines = p.read_text(encoding="utf-8").splitlines()
    cap_lines = {ln for ln, c, _ in audit(p) if c == "CAP"}
    struct = parse_struct(lines)
    blocks, i = [], 0
    while i < len(struct):
        if struct[i][1] == "table":
            s = i
            while i < len(struct) and struct[i][1] == "table":
                i += 1
            blocks.append((s, i - 1))
        else:
            i += 1
    targets = [s for s, _ in blocks if struct[s][0] in cap_lines]
    if len(targets) != len(captions):
        print(f"!! {rel}: CAP 块 {len(targets)} 个 != 表题 {len(captions)} 个,跳过")
        return 0
    for ln, cap in sorted(zip(targets, captions), reverse=True):
        idx = next(j for j, (l, k, t) in enumerate(struct) if l == ln)
        lines[idx:idx] = [f"**{cap}**", ""]
    p.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return len(captions)


def main():
    jobs = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    total = 0
    for rel, caps in jobs.items():
        n = insert_caps(rel, caps)
        total += n
        print(f"{n:3d} {rel}")
    print(f"共插入 {total} 个表题")


if __name__ == "__main__":
    main()
