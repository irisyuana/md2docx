#!/usr/bin/env python3
"""mermaid 转图尺寸/字号回归审计。

对 build/ 下所有 md 中的 mermaid 代码块逐块走真实转换链
（pandoc + reference/diagram.lua + reference/layout.lua），从产物 docx
读取嵌入 PNG 的像素尺寸与 wp:extent 显示尺寸，检查：

- 渲染是否成功（失败时 mermaid 源码会以代码块形式落进 docx）；
- 显示尺寸是否不超过 A4 正文区（上限与 reference/postprocess.py 一致）；
- A4 适配后的有效字号（渲染字号 17px = 12.75pt × 适配率；五号 = 10.5pt）。

用法：
  python3 tools/mermaid_fit_audit.py                 # 审计全部 md
  python3 tools/mermaid_fit_audit.py 路径/xx.md ...  # 只审计指定 md
报告输出到 tools/mermaid_fit_report.md。
"""
import json
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
BUILD = ROOT / "build"
REPORT = Path(__file__).resolve().parent / "mermaid_fit_report.md"

RENDER_SCALE = 5                 # diagram.lua 中 mmdc 的 --scale
CSS_PX_EMU = 9525                # 1/96 英寸的 EMU 数
FONT_PT_AT_100 = 12.75           # 17px 渲染字号在 100% 适配下的磅值
A4_W_EMU, A4_H_EMU = 5_270_500, 7_968_480

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def find_md_files():
    return sorted(BUILD.rglob("*.md"))


def extract_blocks(md_path):
    src = md_path.read_text(encoding="utf-8")
    blocks = []
    for m in re.finditer(r"```mermaid\n(.*?)```", src, re.S):
        code = m.group(1)
        cap = re.match(r"%%\|\s*fig-cap:\s*(.+)", code)
        plain = re.sub(r"^%%\|.*$", "", code, flags=re.M)
        mtype = re.match(r"\s*([A-Za-z][\w-]*)", plain)
        blocks.append({
            "code": code,
            "cap": cap.group(1).strip() if cap else "",
            "type": mtype.group(1) if mtype else "?",
        })
    return blocks


def convert_one(block, workdir):
    """单块走真实转换链，返回 (png_path, extent, stderr_tail)。"""
    doc_md = workdir / "doc.md"
    doc_md.write_text(
        "---\ntitle: fit-audit\n---\n\n```mermaid\n" + block["code"] + "```\n",
        encoding="utf-8")
    out = workdir / "doc.docx"
    env = {**os.environ, "PANDOC_DIAGRAM_TMP": str(workdir / "mmtmp")}
    proc = subprocess.run(
        ["pandoc", str(doc_md), "-o", str(out),
         "--reference-doc", str(ROOT / "reference" / "reference.docx"),
         "--lua-filter", str(ROOT / "reference" / "diagram.lua"),
         "--lua-filter", str(ROOT / "reference" / "layout.lua"),
         "--syntax-highlighting=tango"],
        cwd=ROOT, env=env,
        capture_output=True, text=True, timeout=300)
    if not out.exists():
        return None, None, proc.stderr[-400:]
    return parse_docx(out)


def parse_docx(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        media = [i for i in z.infolist()
                 if i.filename.startswith("word/media/")
                 and i.filename.lower().endswith(".png")
                 and i.file_size > 5000]
        if not media:
            return None, None, "no rendered PNG in docx"
        png = max(media, key=lambda i: i.file_size)
        data = z.read(png.filename)
        w, h = struct.unpack(">II", data[16:24])
        rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
        rid_target = {rel.get("Id"): rel.get("Target")
                      for rel in rels}
        doc = etree.fromstring(z.read("word/document.xml"))
        target = png.filename[len("word/"):]
        extent = None
        for blip in doc.iter(f"{{{A}}}blip"):
            if rid_target.get(blip.get(f"{{{R}}}embed")) == target:
                drawing = blip
                for anc in blip.iterancestors():
                    if anc.tag in (f"{{{WP}}}inline", f"{{{WP}}}anchor"):
                        ext = anc.find(f"{{{WP}}}extent")
                        if ext is not None:
                            extent = (int(ext.get("cx")), int(ext.get("cy")))
                        break
                break
        return (w, h), extent, ""


def audit(files):
    rows = []
    workroot = Path(tempfile.mkdtemp(prefix="mermaid-audit-"))
    try:
        for md in files:
            for idx, block in enumerate(extract_blocks(md), 1):
                wd = workroot / f"{md.stem}-{idx}"
                wd.mkdir(parents=True, exist_ok=True)
                (w, h), extent, err = (None, None), None, ""
                try:
                    (w, h), extent, err = convert_one(block, wd)
                except (subprocess.TimeoutExpired, Exception) as e:
                    err = str(e)[:400]
                rows.append({
                    "md": str(md.relative_to(ROOT)),
                    "idx": idx, "type": block["type"], "cap": block["cap"],
                    "px": (w, h), "extent": extent, "err": err,
                })
                rel = md.relative_to(ROOT)
                status = "OK" if extent else "FAIL"
                print(f"[{len(rows):2d}] {status} {rel}#{idx} {block['type']}",
                      flush=True)
    finally:
        shutil.rmtree(workroot, ignore_errors=True)
    return rows


def analyze(rows):
    for r in rows:
        (w, h), extent, err = r["px"], r["extent"], r["err"]
        if err or not extent or not w:
            r["status"] = "FAIL"
            continue
        css_w, css_h = w / RENDER_SCALE, h / RENDER_SCALE
        cx, cy = extent
        ratio = min(cx / (css_w * CSS_PX_EMU), cy / (css_h * CSS_PX_EMU), 1.0)
        eff = FONT_PT_AT_100 * ratio
        a4_ok = cx <= A4_W_EMU and cy <= A4_H_EMU
        # 显示纵横比须与像素纵横比一致（否则说明嵌入图与定尺寸阶段脱节）
        src_asp, disp_asp = css_w / css_h, cx / cy
        distort = abs(disp_asp - src_asp) / src_asp > 0.02
        r.update(css=(css_w, css_h), disp_cm=(cx / 360000, cy / 360000),
                 ratio=ratio, eff_pt=eff, a4_ok=a4_ok,
                 status="DISTORT" if distort
                 else ("A4_OVER" if not a4_ok
                       else ("OK" if eff >= 10.4 else "SMALL")))
    return rows


def report(rows):
    ok = [r for r in rows if r["status"] == "OK"]
    small = [r for r in rows if r["status"] == "SMALL"]
    over = [r for r in rows if r["status"] == "A4_OVER"]
    dist = [r for r in rows if r["status"] == "DISTORT"]
    fail = [r for r in rows if r["status"] == "FAIL"]
    lines = [
        "# mermaid 转图尺寸/字号审计报告",
        "",
        f"- 总数 {len(rows)}；A4 合规且字号≥五号(10.4pt) {len(ok)}；"
        f"A4 合规但字号低于五号 {len(small)}；超 A4 {len(over)}；"
        f"纵横比失真 {len(dist)}；渲染失败 {len(fail)}",
        "",
        "| 图 | 类型 | 像素(css) | 显示(cm) | 适配率 | 有效字号 | 状态 |",
        "|---|---|---|---|---|---|---|",
    ]
    for r in rows:
        if r["status"] == "FAIL":
            lines.append(f"| {r['md']}#{r['idx']} {r['cap']} | {r['type']} |"
                         f" - | - | - | - | FAIL `{r['err'][:60]}` |")
            continue
        w, h = r["css"]
        dw, dh = r["disp_cm"]
        flag = "" if r["status"] == "OK" else " ⚠"
        lines.append(
            f"| {r['md']}#{r['idx']} {r['cap']} | {r['type']} "
            f"| {w:.0f}x{h:.0f} | {dw:.1f}x{dh:.1f} "
            f"| {r['ratio']:.2f} | {r['eff_pt']:.1f}pt | {r['status']}{flag} |")
    text = "\n".join(lines) + "\n"
    REPORT.write_text(text, encoding="utf-8")
    return text


def main():
    cache = Path("/tmp/mer_audit_rows.json")
    if "--report-only" in sys.argv:
        rows = json.loads(cache.read_text(encoding="utf-8"))
    else:
        files = [Path(p) for p in sys.argv[1:] if not p.startswith("-")] \
            or find_md_files()
        files = [f for f in files if f.exists()]
        rows = audit(files)
        cache.write_text(json.dumps(rows, ensure_ascii=False), encoding="utf-8")
    rows = analyze(rows)
    text = report(rows)
    n_fail = sum(1 for r in rows if r["status"] in ("FAIL", "A4_OVER", "DISTORT"))
    print(text)
    print(f"报告已写入 {REPORT}")
    return 1 if n_fail else 0


if __name__ == "__main__":
    sys.exit(main())
