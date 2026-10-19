#!/usr/bin/env python3
"""reference.docx 去敏感化：把正文实质性文本替换为格式描述（如“宋体一号 加粗 居中”）。

只修改不影响 pandoc 产物的条目：
  - word/document.xml   （pandoc 完全重写，仅 sectPr 被读取；只动 w:t 文本，不碰 sectPr/域代码）
  - docProps/core.xml   （pandoc 重新生成）
不动这些原样拷入产物的条目：headers/footers、word/media/*、theme、fontTable、
webSettings、docProps/app.xml、styles.xml、numbering.xml、settings.xml。

用法： python3 desensitize-reference.py [reference.docx 路径]
已验证：处理后 build.sh 产物与处理前逐条目一致（除 core.xml 构建时间戳）。
"""
import re
import sys
import zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def q(tag):
    return f"{{{W}}}{tag}"


# ---------------------------------------------------------------- 样式解析
def rpr_to_dict(rpr):
    """把 w:rPr 元素转成 {eastAsia, ascii, sz, bold}（只记显式设置的项）。"""
    if rpr is None:
        return {}
    d = {}
    fonts = rpr.find(q("rFonts"))
    if fonts is not None:
        ea = fonts.get(q("eastAsia"))
        asc = fonts.get(q("ascii"))
        if ea:
            d["eastAsia"] = ea
        if asc:
            d["ascii"] = asc
    sz = rpr.find(q("sz"))
    if sz is not None and sz.get(q("val")):
        d["sz"] = sz.get(q("val"))
    b = rpr.find(q("b"))
    if b is not None:
        val = b.get(q("val"))
        d["bold"] = val not in ("0", "false", "off")
    return d


def load_styles(styles_bytes):
    root = etree.fromstring(styles_bytes)
    base = {}
    dd = root.find(q("docDefaults"))
    if dd is not None:
        rprd = dd.find(q("rPrDefault"))
        if rprd is not None:
            base = rpr_to_dict(rprd.find(q("rPr")))

    styles = {}
    for st in root.findall(q("style")):
        sid = st.get(q("styleId"))
        based = st.find(q("basedOn"))
        styles[sid] = {
            "name": (st.find(q("name")).get(q("val")) if st.find(q("name")) is not None else sid),
            "basedOn": based.get(q("val")) if based is not None else None,
            "rpr": rpr_to_dict(st.find(q("rPr"))),
        }
    return base, styles


def resolve_style(sid, base, styles):
    """沿 basedOn 链自底向上合并 rPr，返回有效 {eastAsia, ascii, sz, bold}。"""
    chain = []
    while sid and sid in styles and sid not in chain:  # 防循环
        chain.append(sid)
        sid = styles[sid]["basedOn"]
    eff = dict(base)
    for s in reversed(chain):
        eff.update(styles[s]["rpr"])
    return eff


# ---------------------------------------------------------------- 描述生成
SIZE_NAMES = {
    42: "初号", 36: "小初", 26: "一号", 24: "小一", 22: "二号", 18: "小二",
    16: "三号", 15: "小三", 14: "四号", 12: "小四", 10.5: "五号", 9: "小五",
    7.5: "六号", 6.5: "小六", 5.5: "七号", 5: "八号",
}
JC_NAMES = {"center": "居中", "right": "右对齐", "both": "两端对齐",
            "distribute": "分散对齐"}


def describe(eff, jc):
    font = eff.get("eastAsia") or eff.get("ascii") or "宋体"
    if eff.get("sz"):
        pt = int(eff["sz"]) / 2.0
        size = SIZE_NAMES.get(pt, f"{pt:g}磅")
    else:
        size = "五号"
    parts = [f"{font}{size}"]
    if eff.get("bold"):
        parts.append("加粗")
    if jc in JC_NAMES:
        parts.append(JC_NAMES[jc])
    return " ".join(parts)


# ---------------------------------------------------------------- 主流程
def desensitize_document(doc_bytes, base, styles):
    root = etree.fromstring(doc_bytes)

    # 每个 w:t 归属到最近的 w:p 祖先（天然处理文本框嵌套段落）
    groups = {}  # id(w:p) -> {"p": elem, "ts": [w:t...]}
    for t in root.iter(q("t")):
        p, anc = None, t.getparent()
        while anc is not None:
            if anc.tag == q("p"):
                p = anc
                break
            anc = anc.getparent()
        if p is None:
            continue
        g = groups.setdefault(id(p), {"p": p, "ts": []})
        g["ts"].append(t)

    n = 0
    samples = []
    for g in groups.values():
        p, ts = g["p"], g["ts"]
        ppr = p.find(q("pPr"))
        sid = jc = None
        if ppr is not None:
            ps = ppr.find(q("pStyle"))
            if ps is not None:
                sid = ps.get(q("val"))
            jce = ppr.find(q("jc"))
            if jce is not None:
                jc = jce.get(q("val"))
        eff = resolve_style(sid, base, styles)
        # 首个 run 的 rPr 覆盖样式结果
        r = ts[0].getparent()
        while r is not None and r.tag != q("r"):
            r = r.getparent()
        if r is not None:
            eff.update(rpr_to_dict(r.find(q("rPr"))))

        desc = describe(eff, jc)
        old = "".join(t.text or "" for t in ts)
        ts[0].text = desc
        for t in ts[1:]:
            t.text = None
        n += 1
        if old.strip() and len(samples) < 15:
            samples.append((old[:40], desc))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                          standalone=True), n, samples


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else "reference.docx"
    with zipfile.ZipFile(path) as z:
        contents = {n: z.read(n) for n in z.namelist()}
        infos = z.infolist()

    base, styles = load_styles(contents["word/styles.xml"])
    new_doc, n, samples = desensitize_document(
        contents["word/document.xml"], base, styles)
    contents["word/document.xml"] = new_doc

    # core.xml 作者字段清理（该条目由 pandoc 重新生成，不影响产物）
    core = contents["docProps/core.xml"].decode("utf-8")
    core = re.sub(r"(<dc:creator>).*?(</dc:creator>)", r"\1user\2", core)
    core = re.sub(r"(<cp:lastModifiedBy>).*?(</cp:lastModifiedBy>)", r"\1user\2", core)
    contents["docProps/core.xml"] = core.encode("utf-8")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info in infos:
            z.writestr(info, contents[info.filename])

    print(f"已处理段落数: {n}")
    print("替换样例（原文 → 格式描述）:")
    for old, desc in samples:
        print(f"  {old!r} → {desc!r}")


if __name__ == "__main__":
    main()
