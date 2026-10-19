#!/usr/bin/env python3
"""把 md 里的 mermaid 图导出为可在 Inkscape/LibreOffice 中直接编辑的 SVG。

与管线渲染同源（同 ELK/原生布局、同字体、同黑白主题），并额外打开根级
htmlLabels:false —— 默认的 HTML 标签（foreignObject）在 Inkscape/
LibreOffice 里不可见也无法编辑，纯文本标签才能点选改字。

方向自适应与管线一致：适配 A4 不佳时按换向重排再测一次，取适配率更高的一版。

用法：
  python3 tools/mermaid_svg_export.py build/xx/yy.md 2            # 导第 2 个图
  python3 tools/mermaid_svg_export.py build/xx/yy.md --cap 状态机  # 按题注关键字
输出：md 同目录 svg/ 下的 .svg，随后用 Inkscape 打开编辑、导出 PNG 即可。
"""
import argparse
import json
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MMDC = ROOT / "tools" / "mermaid-cli" / "node_modules" / ".bin" / "mmdc"
A4_W_EMU, A4_H_EMU = 5_270_500, 7_968_480 - 360_000   # 与 diagram.lua 一致（高度已预留题注）
EMU_PER_CSS_PX = 9525
FLOW_SWAP = {"LR": "TB", "RL": "BT", "TB": "LR", "TD": "LR", "BT": "RL"}

THEME = {
    "theme": "base",
    # 根级 htmlLabels:false：ELK 的节点标签也走纯 <text>，编辑器才能改
    "htmlLabels": False,
    "themeVariables": {
        "fontSize": "17px",
        "fontFamily": "Noto Sans CJK SC, PingFang SC, Microsoft YaHei, sans-serif",
        "background": "#ffffff",
        "primaryColor": "#ffffff",
        "primaryTextColor": "#000000",
        "primaryBorderColor": "#000000",
        "lineColor": "#000000",
        "secondaryColor": "#ffffff",
        "tertiaryColor": "#ffffff",
        "mainBkg": "#ffffff",
        "nodeBorder": "#000000",
        "clusterBkg": "#ffffff",
        "clusterBorder": "#000000",
        "edgeLabelBackground": "#ffffff",
        "stateBorder": "#000000",
        "labelBoxBkgColor": "#ffffff",
        "labelBoxBorderColor": "#000000",
        "labelTextColor": "#000000",
        "altBackground": "#ffffff",
        "noteBkgColor": "#ffffff",
        "noteBorderColor": "#000000",
    },
}

ELK_EXTRA = {
    "layout": "elk",
    "flowchart": {"nodeSpacing": 30, "rankSpacing": 40, "diagramPadding": 8,
                  "useMaxWidth": False, "htmlLabels": False},
    "state": {"useMaxWidth": False, "htmlLabels": False},
}
PLAIN_EXTRA = {
    "sequence": {"actorMargin": 36, "messageMargin": 22, "noteMargin": 8,
                 "boxMargin": 6, "diagramMarginX": 18, "diagramMarginY": 10,
                 "mirrorActors": False, "useMaxWidth": False},
    "er": {"useMaxWidth": False},
    "class": {"useMaxWidth": False, "htmlLabels": False},
    "gantt": {"useMaxWidth": False},
    "pie": {"useMaxWidth": False},
    "journey": {"useMaxWidth": False},
    "requirement": {"useMaxWidth": False, "htmlLabels": False},
    "c4": {"useMaxWidth": False},
}


def first_statement_line(source):
    for line in source.splitlines():
        if not line.lstrip().startswith("%%") and line.strip():
            return line
    return ""


def diagram_type(source):
    m = re.match(r"\s*([A-Za-z][\w-]*)", first_statement_line(source) or "")
    return m.group(1) if m else ""


def fit_score(w, h):
    if not w or not h:
        return 0.0
    return min(A4_W_EMU / (w * EMU_PER_CSS_PX),
               A4_H_EMU / (h * EMU_PER_CSS_PX), 1.0)


def svg_size(path):
    head = path.open(encoding="utf-8").read(2000)
    mw = re.search(r'\bwidth="([\d.]+)', head)
    mh = re.search(r'\bheight="([\d.]+)', head)
    return (float(mw.group(1)), float(mh.group(1))) if mw and mh else (None, None)


def swap_flow_header(line):
    head, d = re.match(r"^(\s*flowchart\s+)(\S+)", line) or (None, None)
    if not head:
        head, d = re.match(r"^(\s*graph\s+)(\S+)", line) or (None, None)
    swap = FLOW_SWAP.get(d) if d else None
    return head + swap if head and swap else None


def flow_variant(source):
    lines = source.split("\n")
    for i, line in enumerate(lines):
        if line.strip() and not line.lstrip().startswith("%%"):
            repl = swap_flow_header(line)
            if repl:
                lines[i] = repl
                return "\n".join(lines)
            return None
    return None


def state_variant(source):
    lines = source.split("\n")
    header = dir_idx = None
    for i, line in enumerate(lines):
        if line.strip() and not line.lstrip().startswith("%%"):
            if header is None:
                header = i
            elif re.match(r"^direction\s+\w+\s*$", line):
                dir_idx = i
                break
    if header is None:
        return None
    if dir_idx is not None:
        d = re.match(r"^direction\s+(\w+)", lines[dir_idx]).group(1)
        if d not in FLOW_SWAP:
            return None
        lines[dir_idx] = "direction " + FLOW_SWAP[d]
    else:
        lines.insert(header + 1, "direction LR")
    return "\n".join(lines)


def render(mmdc, cfg_path, source, infile, outfile):
    Path(infile).write_text(source, encoding="utf-8")
    r = subprocess.run([str(mmdc), "--configFile", str(cfg_path),
                        "-i", str(infile), "-o", str(outfile)],
                       cwd=ROOT, capture_output=True, text=True, timeout=180)
    if not outfile.exists():
        return None, None, r.stderr[-300:]
    return (*svg_size(outfile), "")


def flatten_css(svg_path):
    """把 SVG 内 <style> 块的 CSS 规则展开为各元素的 style 属性。

    mermaid 把 fill/stroke 等写在 <style> 里，Inkscape 1.1 / LibreOffice
    对其选择器支持不全，打开会显示成黑块；展开后任何编辑器所见即所得。"""
    import re
    from lxml import etree
    text = Path(svg_path).read_text(encoding="utf-8")
    m = re.search(r"<style[^>]*>(.*?)</style>", text, re.S)
    if not m:
        return
    css = re.sub(r"@[^{}]+\{(?:[^{}]*\{[^{}]*\})*[^{}]*\}", "", m.group(1))
    rules = re.findall(r"([^{}]+)\{([^{}]*)\}", css)

    def classes(el):
        return (el.get("class") or "").split()

    def compound(el, comp):
        m = re.fullmatch(r"(?:#([\w-]+))?((?:\.[\w-]+)*)([a-zA-Z][\w-]*)?", comp)
        if not m:
            return False
        eid, clss, tag = m.group(1), m.group(2), m.group(3)
        if eid and el.get("id") != eid:
            return False
        if tag and etree.QName(el).localname != tag:
            return False
        if clss and not set(clss.strip(".").split(".")) <= set(classes(el)):
            return False
        return True

    def matches(el, comps):
        return compound(el, comps[-1]) and match_ancestors(
            el.getparent(), comps[:-1])

    def match_ancestors(el, comps):
        """comps[-1] 从 el 自身（含）向上逐层找，其余依次匹配更上层。"""
        if not comps:
            return True
        anc = el
        while anc is not None:
            if compound(anc, comps[-1]) and match_ancestors(anc.getparent(), comps[:-1]):
                return True
            anc = anc.getparent()
        return False

    tree = etree.fromstring(text.encode("utf-8"))
    for selectors, body in rules:
        body = body.strip().rstrip(";")
        if not body:
            continue
        for sel in selectors.split(","):
            comps = sel.strip().split()
            if not comps or ":" in comps[-1]:
                continue  # 跳过 :root 等伪类规则
            for el in tree.iter():
                try:
                    hit = matches(el, comps)
                except Exception:
                    hit = False
                if hit:
                    prev = el.get("style", "").rstrip(";")
                    el.set("style", f"{prev};{body}" if prev else body)
    style_el = tree.find("{http://www.w3.org/2000/svg}style")
    if style_el is not None:
        tree.remove(style_el)
    # 白色画布：编辑器里所见即文档效果（背景透明时部分工具显示为黑/棋盘格）
    NS = "http://www.w3.org/2000/svg"
    bg = etree.Element("{%s}rect" % NS, {"width": "100%", "height": "100%",
                                         "fill": "#ffffff"})
    tree.insert(0, bg)
    Path(svg_path).write_bytes(etree.tostring(tree, xml_declaration=True,
                                              encoding="UTF-8", standalone=True))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("md")
    ap.add_argument("index", nargs="?", type=int, default=None)
    ap.add_argument("--cap", default=None, help="按题注关键字选图")
    args = ap.parse_args()

    src = Path(args.md)
    single = src.suffix == ".mmd"          # 直接给 .mmd 文件：只导这一张
    md = src
    if single:
        blocks = [src.read_text(encoding="utf-8")]
        args.index = 1
    else:
        blocks = re.findall(r"```mermaid\n(.*?)```",
                            src.read_text(encoding="utf-8"), re.S)
    if args.cap and not single:
        targets = [(i, b) for i, b in enumerate(blocks, 1)
                   if args.cap in b]
        if not targets:
            sys.exit(f"没有题注包含 {args.cap!r} 的 mermaid 块")
        picks = [targets[0]]
    elif args.index is not None:
        picks = [(args.index, blocks[args.index - 1])]
    else:
        sys.exit("请给出图序号（或 --cap 题注关键字；.mmd 文件可省略）")

    work = Path(tempfile.mkdtemp(prefix="mermaid-svg-"))
    elk_cfg, plain_cfg = work / "elk.json", work / "plain.json"
    elk_cfg.write_text(json.dumps({**THEME, **ELK_EXTRA}, ensure_ascii=False, indent=1))
    plain_cfg.write_text(json.dumps({**THEME, **PLAIN_EXTRA}, ensure_ascii=False, indent=1))

    outdir = md.parent / "svg"
    outdir.mkdir(exist_ok=True)
    for idx, code in picks:
        plain = re.sub(r"^%%\|.*$", "", code, flags=re.M)
        dtype = diagram_type(code)
        use_elk = dtype in ("flowchart", "graph") or dtype.startswith("stateDiagram")
        cfg = elk_cfg if use_elk else plain_cfg
        base_out = work / f"fig{idx}-base.svg"
        var_out = work / f"fig{idx}-var.svg"
        w, h, err = render(MMDC, cfg, code, work / "in.mmd", base_out)
        if w is None:
            print(f"[失败] {md.name}#{idx}: {err}")
            continue
        best = (fit_score(w, h), base_out, (w, h))
        if best[0] < 0.85 and dtype in ("flowchart", "graph"):
            variant = flow_variant(code)
        elif best[0] < 0.85 and dtype.startswith("stateDiagram"):
            variant = state_variant(code)
        else:
            variant = None
        if variant:
            vw, vh, _ = render(MMDC, cfg, variant, work / "in_var.mmd", var_out)
            if vw and fit_score(vw, vh) > best[0]:
                best = (fit_score(vw, vh), var_out, (vw, vh))
        cap = re.match(r"%%\|\s*fig-cap:\s*(.+)", code)
        if single:
            dest = outdir / f"{md.stem}.svg"
        else:
            name = (cap.group(1).strip() if cap else f"fig{idx}")
            name = re.sub(r"[\\/:*?\"<>|\s]+", "-", name)
            dest = outdir / f"{md.stem}-图{idx}-{name}.svg"
        shutil.copy(best[1], dest)
        flatten_css(dest)
        print(f"[完成] {dest}")
        if single:
            print("       Inkscape 打开编辑 → 导出 PNG（建议 300dpi）")
        else:
            print("       Inkscape 打开编辑 → 导出 PNG（建议 300dpi）→ md 里把该 mermaid 块"
                  f"替换为 ![](svg/{dest.stem}.png) 后重转")


if __name__ == "__main__":
    main()
