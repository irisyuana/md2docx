#!/usr/bin/env python3
"""docx 域烘焙（外科移植式）：LibreOffice 只负责计算目录，算好的目录条目与
标题书签被移植回 pandoc 原版 XML，其余部分不经 LO 改写。

为什么不做 LO 整文档往返：LO 重写的 document.xml/styles.xml 在 LibreOffice
里渲染正常，但 WPS/Word 下正文会出现整行裁切（行高语义被 LO 改写所致，
2026-09 实测）。本脚本保证 WPS 看到的正文与 pandoc 产物逐字节一致，仅三处例外：
  1. 目录 sdt 整体替换为 LO 计算结果（样式名映射回原版：
     Contents1/2/3→TOC1/2/3、ContentsHeading→TOCHeading，剥 IndexLink/webHidden）；
  2. 标题段落插入 LO 的 __RefHeading__* 书签（目录超链接的跳转目标）；
  3. settings.xml 剥离 updateFields、document.xml 剥离 w:dirty。

效果：Word/WPS 打开即显示目录（缓存内容），不再弹「是否更新域」。
注意：目录页码按 LibreOffice 排版计算，与 Word 实际排版可能有 ±1 页出入；
     Word 里手改文档后目录不自动变，需手动更新（选中目录按 F9，会整表重建）。
     重转 markdown 由管线自动重新烘焙，无需手动。

用法： python3 bake_fields.py 文件1.docx [文件2.docx ...]
全程先写 .tmp 再原子替换，失败文件保持原样；任一失败退出码 1。
"""
import bisect
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import time
import zipfile

import uno
from com.sun.star.beans import PropertyValue
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}

UPDATE_FIELDS_RE = re.compile(
    r"<w:updateFields\b[^>]*/>|<w:updateFields\b[^>]*>.*?</w:updateFields>",
    re.S,
)
DIRTY_RE = re.compile(r'\s+w:dirty="true"')
DONOR_ID_BASE = 10000  # 移植书签的 w:id 起点，避开 pandoc 原有 id


def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def q(t):
    return f"{{{W}}}{t}"


def para_text(p):
    return re.sub(r"\s+", "", "".join(p.itertext()))


def strip_update_fields(path):
    """settings.xml 剥离 updateFields；document.xml 剥离 dirty。"""
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        items = {i.filename: z.read(i.filename) for i in infos}
    changed = False
    if "word/settings.xml" in items:
        new = UPDATE_FIELDS_RE.sub("", items["word/settings.xml"].decode("utf-8"))
        if new != items["word/settings.xml"].decode("utf-8"):
            items["word/settings.xml"] = new.encode("utf-8")
            changed = True
    if "word/document.xml" in items:
        new = DIRTY_RE.sub("", items["word/document.xml"].decode("utf-8"))
        if new != items["word/document.xml"].decode("utf-8"):
            items["word/document.xml"] = new.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for i in infos:
            z.writestr(i, items[i.filename])
    return changed


def find_toc_sdt(root):
    """定位包含 TOC 域指令的 w:sdt 元素。"""
    for sdt in root.iter(q("sdt")):
        for instr in sdt.iter(q("instrText")):
            if instr.text and "TOC" in instr.text:
                return sdt
    return None


def style_name_map(styles_root):
    """样式 ID → 样式名。部分 LibreOffice 版本（实测 7.3.7-0ubuntu0.22.04.13）
    导出 docx 时把内置标题/目录样式写成数字 ID（如 Heading 1 写作 "1"），
    仅凭样式 ID 无法识别标题，需回查 styles.xml 的样式名。"""
    names = {}
    if styles_root is None:
        return names
    for st in styles_root.iter(q("style")):
        sid = st.get(q("styleId"))
        nm = st.find(q("name"))
        if sid is not None:
            names[sid] = nm.get(q("val"), "") if nm is not None else ""
    return names


def style_rewrite(sdt, names=None):
    """供体目录样式名映射回原版，并剥掉链接字符样式。"""
    names = names or {}
    for p in sdt.iter(q("pStyle")):
        v = p.get(q("val"), "")
        nm = names.get(v, v)
        m = re.fullmatch(r"(?:Contents|TOC)\s*([123])", v) or re.fullmatch(
            r"(?:Contents|TOC)\s*([123])", nm)
        if m:
            p.set(q("val"), f"TOC{m.group(1)}")
        elif v == "ContentsHeading" or re.fullmatch(
                r"(?:Contents|TOC)\s*Heading", nm):
            p.set(q("val"), "TOCHeading")
    for tag in ("rStyle",):
        for el in list(sdt.iter(q(tag))):
            if el.get(q("val")) in ("IndexLink",):
                el.getparent().remove(el)
    for el in list(sdt.iter(q("webHidden"))):
        el.getparent().remove(el)


def _heading_style_name(p, names=None):
    """段落标题样式名；非标题段落返回 None。"""
    ppr = p.find(q("pPr"))
    if ppr is None:
        return None
    st = ppr.find(q("pStyle"))
    if st is None:
        return None
    val = st.get(q("val"), "")
    nm = (names or {}).get(val, val)
    return nm if re.match(r"(Appendix\s*)?Heading\b", nm, re.I) else None


def _is_heading(p, names=None):
    return _heading_style_name(p, names) is not None


def donor_bookmarks(donor_root, names=None):
    """按文档顺序取供体标题书签 [(anchor, heading_text)]。

    LO 通常把书签放在标题段落内部，但个别标题（实测如文末「附录A」）的
    书签是其段落的兄弟节点，两种位置都要认领。"""
    order = list(donor_root.iter())
    pos = {id(el): i for i, el in enumerate(order)}
    heads = [p for p in order if p.tag == q("p") and _is_heading(p, names)]
    head_pos = sorted(pos[id(p)] for p in heads)
    out = []
    for b in donor_root.iter(q("bookmarkStart")):
        name = b.get(q("name"), "")
        if not name.startswith("__RefHeading__"):
            continue
        hp = next((p for p in b.iterancestors(q("p"))
                   if _is_heading(p, names)), None)
        if hp is not None:  # 书签在标题段落内
            out.append((name, para_text(hp)))
        else:               # 书签是段落兄弟节点：归属其后第一个标题
            i = bisect.bisect_right(head_pos, pos[id(b)])
            if i >= len(heads):
                raise RuntimeError(f"书签 {name} 后找不到标题段落")
            out.append((name, para_text(heads[i])))
    return out


def transplant(recipient_root, donor_root, donor_names=None,
               recipient_names=None):
    """把供体目录 sdt 与标题书签移植进接收方，返回 (条目数, 书签数)。"""
    rec_sdt = find_toc_sdt(recipient_root)
    don_sdt = find_toc_sdt(donor_root)
    if rec_sdt is None:  # 本来就没有目录的文档：只剥标志
        return 0, 0
    if don_sdt is None:
        raise RuntimeError("LO 产物中没有目录域，无法移植")

    style_rewrite(don_sdt, donor_names)
    marks = donor_bookmarks(donor_root, donor_names)

    # 目录条目里的锚 → 校验都有对应标题书签
    entry_anchors = [h.get(q("anchor"), "") for h in don_sdt.iter(q("hyperlink"))]
    anchor_text = dict(marks)
    missing = [a for a in entry_anchors if a not in anchor_text]
    if missing:
        raise RuntimeError(f"目录锚缺书签: {missing[:3]}")

    # 接收方标题段落（按文档顺序）。样式必须精确匹配 Heading*/AppendixHeading*：
    # 用 contains 会把 TOCHeading（目录页标题，文本恰为「目录」）也卷进来，
    # 导致正文里同名的「目录」章节拿不到书签
    def is_rec_heading(p):
        return _is_heading(p, recipient_names)

    rec_heads = [p for p in recipient_root.iter(q("p")) if is_rec_heading(p)]

    cursor = 0
    placed = 0
    for i, (anchor, text) in enumerate(marks):
        hit = None
        for j in range(cursor, len(rec_heads)):
            if para_text(rec_heads[j]) == text:
                hit = j
                break
        if hit is None:
            raise RuntimeError(f"标题未匹配: {text[:30]!r}（锚 {anchor}）")
        cursor = hit + 1
        p = rec_heads[hit]
        ppr = p.find(q("pPr"))
        idx = list(p).index(ppr) + 1
        for tag in ("bookmarkStart", "bookmarkEnd"):
            b = etree.Element(q(tag))
            b.set(q("id"), str(DONOR_ID_BASE + i))
            if tag == "bookmarkStart":
                b.set(q("name"), anchor)
            p.insert(idx, b)
        placed += 1

    # 整体替换目录 sdt（保留原节点的尾部空白，做到字节级最小改动）
    import copy
    new_sdt = copy.deepcopy(don_sdt)
    new_sdt.tail = rec_sdt.tail
    parent = rec_sdt.getparent()
    parent[list(parent).index(rec_sdt)] = new_sdt
    return len(entry_anchors), placed


def bake_one(desktop, path):
    """LO 计算目录 → 从供体提取目录 sdt+书签移植回原文档 → 剥标志 → 原子替换。"""
    donor = path + ".bake.donor.tmp"
    shutil.copyfile(path, donor)

    def strip_dirty(data):
        return DIRTY_RE.sub("", data.decode("utf-8")).encode("utf-8")

    # LO 对带 dirty 的索引起 update() 前先剥掉（干净，且 LO 本来也不回写）
    with zipfile.ZipFile(donor) as z:
        infos = z.infolist()
        items = {i.filename: z.read(i.filename) for i in infos}
    items["word/document.xml"] = strip_dirty(items["word/document.xml"])
    with zipfile.ZipFile(donor, "w", zipfile.ZIP_DEFLATED) as z:
        for i in infos:
            z.writestr(i, items[i.filename])

    url = uno.systemPathToFileUrl(os.path.abspath(donor))
    doc = desktop.loadComponentFromURL(url, "_blank", 0, (
        prop("Hidden", True),
        prop("ReadOnly", False),
        prop("UpdateDocMode", 3),  # QUIET_UPDATE：加载时不更新外部链接
        prop("FilterName", "MS Word 2007 XML"),
    ))
    n = 0
    try:
        indexes = doc.getDocumentIndexes()
        n = indexes.getCount()
        try:
            doc.getTextFields().refresh()  # 部分 pyuno 版本缺该接口，失败无妨
        except Exception:
            pass
        for _ in range(2):
            for i in range(indexes.getCount()):
                indexes.getByIndex(i).update()
            doc.refresh()
        out = path + ".bake.out.tmp"
        doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(out)), (
            prop("FilterName", "MS Word 2007 XML"),
            prop("Overwrite", True),
        ))
    finally:
        doc.close(False)

    # —— 外科移植：只动 document.xml 的目录区与 settings.xml 标志 ——
    with zipfile.ZipFile(out) as z:
        donor_root = etree.fromstring(z.read("word/document.xml"))
        donor_names = style_name_map(etree.fromstring(z.read("word/styles.xml")))
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        items = {i.filename: z.read(i.filename) for i in infos}
    recipient_root = etree.fromstring(items["word/document.xml"])
    recipient_names = style_name_map(
        etree.fromstring(items["word/styles.xml"]))

    n_entries, n_marks = transplant(
        recipient_root, donor_root, donor_names, recipient_names)

    items["word/document.xml"] = etree.tostring(
        recipient_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    settings = items.get("word/settings.xml", b"").decode("utf-8")
    items["word/settings.xml"] = UPDATE_FIELDS_RE.sub(
        "", settings).encode("utf-8")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for i in infos:
            z.writestr(i, items[i.filename])
    os.replace(out, path)
    os.remove(donor)
    return n, n_entries, n_marks


def main():
    paths = sys.argv[1:]
    if not paths:
        print(__doc__)
        sys.exit(2)

    profile = tempfile.mkdtemp(prefix="lo_bake_")
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        port = s.getsockname()[1]
    proc = subprocess.Popen(
        [
            "soffice", "--headless", "--invisible", "--norestore",
            "--nologo", "--nodefault", "--nofirststartwizard",
            f"-env:UserInstallation=file://{profile}",
            f"--accept=socket,host=127.0.0.1,port={port};urp;"
            "StarOffice.ServiceManager",
        ],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )

    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    ctx = None
    for _ in range(120):  # 冷启动首次可能要几十秒
        try:
            ctx = resolver.resolve(
                f"uno:socket,host=127.0.0.1,port={port};urp;"
                "StarOffice.ComponentContext")
            break
        except Exception:
            if proc.poll() is not None:
                break
            time.sleep(0.5)
    if ctx is None:
        print("bake_fields: 无法连接 LibreOffice（soffice 启动失败？）",
              file=sys.stderr)
        shutil.rmtree(profile, ignore_errors=True)
        sys.exit(1)

    desktop = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx)
    failures = []
    try:
        for path in paths:
            try:
                if not os.path.isfile(path):
                    raise FileNotFoundError(path)
                n, entries, marks = bake_one(desktop, path)
                print(f"烘焙 {path}：更新目录 {n} 个；"
                      f"移植条目 {entries} 条/书签 {marks} 个")
            except Exception as e:
                import traceback
                traceback.print_exc()
                failures.append((path, e))
            finally:
                for suffix in (".bake.donor.tmp", ".bake.out.tmp"):
                    if os.path.exists(path + suffix):
                        os.remove(path + suffix)
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()
        shutil.rmtree(profile, ignore_errors=True)

    if failures:
        for path, e in failures:
            print(f"烘焙失败 {path}：{e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
