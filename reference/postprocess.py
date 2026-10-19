#!/usr/bin/env python3
"""模块文档 docx 后处理：
1. 嵌入式图片（wp:inline）→ 浮动图片（wp:anchor，上下型环绕、水平居中）
2. 紧凑章节：Heading1 段落显式 pageBreakBefore=0，覆盖模板样式的段前分页
   （封面/目录分页由 layout.lua 插入的分页符保留，不受影响）
3. 公式样式：显示公式段落（含 oMathPara）pStyle 改指「公式」样式。
   正文 BodyText 为固定行距 20 磅，高于行距的公式会被上下行文字覆盖；
   「公式」样式复用图格式（自动行距、居中、段前后 18 磅），段落完整撑开。

用法： python3 postprocess.py 文件.docx
仅调整 word/document.xml，其余条目原样保留。
"""
import sys
import zipfile
from lxml import etree

WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# A4 正文区域内的图片尺寸上限（EMU）。只缩小超限图片，不放大。
MAX_IMAGE_WIDTH_EMU = 5_270_500
MAX_IMAGE_HEIGHT_EMU = 7_968_480


def q(ns, t):
    return f"{{{ns}}}{t}"


def float_images(root):
    n = 0
    for inline in list(root.iter(q(WP, "inline"))):
        parent = inline.getparent()
        idx = parent.index(inline)

        anchor = etree.Element(q(WP, "anchor"))
        for k, v in {
            "distT": "0", "distB": "0", "distL": "114300", "distR": "114300",
            "simplePos": "0", "relativeHeight": str(1000 + n),
            "behindDoc": "0", "locked": "0", "layoutInCell": "1",
            "allowOverlap": "1",
        }.items():
            anchor.set(k, v)

        sp = etree.SubElement(anchor, q(WP, "simplePos"))
        sp.set("x", "0")
        sp.set("y", "0")
        ph = etree.SubElement(anchor, q(WP, "positionH"))
        ph.set("relativeFrom", "margin")
        etree.SubElement(ph, q(WP, "align")).text = "center"
        pv = etree.SubElement(anchor, q(WP, "positionV"))
        pv.set("relativeFrom", "paragraph")
        etree.SubElement(pv, q(WP, "posOffset")).text = "0"

        # 子元素必须严格按 CT_Anchor schema 顺序：
        # simplePos, positionH, positionV, extent, effectExtent, wrap*, docPr, cNvGraphicFramePr, graphic
        # （wrap 必须在 docPr 之前，顺序错误会导致 Word 拒绝打开/报无法读取的内容）
        for tag, ns in (("extent", WP), ("effectExtent", WP)):
            el = inline.find(q(ns, tag))
            if el is not None:
                if tag == "extent":
                    cx = int(el.get("cx", "0"))
                    cy = int(el.get("cy", "0"))
                    if cx > 0 and cy > 0:
                        scale = min(
                            1.0,
                            MAX_IMAGE_WIDTH_EMU / cx,
                            MAX_IMAGE_HEIGHT_EMU / cy,
                        )
                        if scale < 1.0:
                            el.set("cx", str(round(cx * scale)))
                            el.set("cy", str(round(cy * scale)))
                anchor.append(el)
        etree.SubElement(anchor, q(WP, "wrapTopAndBottom"))
        for tag, ns in (("docPr", WP), ("cNvGraphicFramePr", WP)):
            el = inline.find(q(ns, tag))
            if el is not None:
                anchor.append(el)
        g = inline.find(q(A, "graphic"))
        if g is not None:
            anchor.append(g)

        parent[idx] = anchor
        n += 1
    return n


def compact_h1(root):
    """Heading1 段落显式关闭段前分页（直接格式优先于样式定义）。"""
    n = 0
    for p in root.iter(q(W, "p")):
        ppr = p.find(q(W, "pPr"))
        if ppr is None:
            continue
        st = ppr.find(q(W, "pStyle"))
        if st is None or st.get(q(W, "val")) != "Heading1":
            continue
        if ppr.find(q(W, "pageBreakBefore")) is not None:
            continue
        el = etree.Element(q(W, "pageBreakBefore"))
        el.set(q(W, "val"), "0")
        ppr.insert(list(ppr).index(st) + 1, el)
        n += 1
    return n


def style_formulas(root):
    """显示公式段落（含 m:oMathPara）pStyle 改指「Formula」（公式）样式。

    段落须整段为显示公式（pandoc 生成的独立公式段），行内公式不动。
    pPr 须为 w:p 首子元素、pStyle 须为 pPr 首子元素（OOXML schema 顺序）。"""
    n = 0
    for p in root.iter(q(W, "p")):
        if p.find(f".//{q(M, 'oMathPara')}") is None:
            continue
        ppr = p.find(q(W, "pPr"))
        if ppr is None:
            ppr = etree.Element(q(W, "pPr"))
            p.insert(0, ppr)
        st = ppr.find(q(W, "pStyle"))
        if st is None:
            st = etree.Element(q(W, "pStyle"))
            ppr.insert(0, st)
        if st.get(q(W, "val")) != "Formula":
            st.set(q(W, "val"), "Formula")
            n += 1
    return n


def main():
    path = sys.argv[1]
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        items = {i.filename: z.read(i.filename) for i in infos}
    doc = etree.fromstring(items["word/document.xml"])
    n_img = float_images(doc)
    n_h1 = compact_h1(doc)
    n_math = style_formulas(doc)
    items["word/document.xml"] = etree.tostring(
        doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for i in infos:
            z.writestr(i, items[i.filename])
    print(f"后处理: 图片浮动化 {n_img} 张；紧凑章节（Heading1 去分页）{n_h1} 处；公式样式（自动行距）{n_math} 处")


if __name__ == "__main__":
    main()
