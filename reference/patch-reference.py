#!/usr/bin/env python3
"""对 reference.docx 做一次性样式补丁（自动备份原文件）。

适配 2025-09 新模板（WPS 导出版，样式含 rsid/theme 属性）：
1. 保留 Heading1-9 的多级自动编号（md 手写编号由 layout.lua 构建时剥离）
2. Title：去 outlineLvl（不进目录）、去首行缩进（居中标题需要）、字号保持 26pt
3. 段前分页：Heading1、TOCHeading 开启；Heading2-9 显式关闭（阻断继承）
4. TOCHeading：basedOn 切断为 Normal（与正文标题完全解耦）
5. Table：字体设为 仿宋_GB2312 + Times New Roman，五号（10.5pt）
6. 克隆 AppendixHeading1-4 无编号标题样式（附录章节使用）
7. settings.xml 增加 updateFields（Word/WPS 打开时自动刷新目录域）
"""
import os
import re
import shutil
import zipfile

PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'reference.docx')

with zipfile.ZipFile(PATH) as zin:
    items = {n: zin.read(n) for n in zin.namelist()}

styles = items['word/styles.xml'].decode('utf-8')


def style_block(styles_xml, sid):
    m = re.search(
        r'<w:style\b[^>]*w:styleId="%s"[^>]*>.*?</w:style>' % re.escape(sid),
        styles_xml, re.S)
    return m.group(0) if m else None


def add_page_break_before(block, on=True):
    """在 pPr 中按 OOXML schema 顺序插入 pageBreakBefore
    （须位于 keepNext/keepLines 之后、widowControl/numPr/tabs/spacing 之前）。"""
    val = '' if on else ' w:val="0"'
    tag = '<w:pageBreakBefore%s />' % val
    for anchor in ('<w:widowControl', '<w:numPr', '<w:suppressLineNumbers',
                   '<w:pBdr', '<w:shd', '<w:tabs', '<w:spacing', '<w:ind',
                   '<w:jc', '<w:outlineLvl', '</w:pPr>'):
        i = block.find(anchor)
        if i >= 0:
            return block[:i] + tag + block[i:]
    return block


# 0. 模板去重：WPS 导出的模板可能含重复 styleId（实测 Figure 出现两次），
#    重复 ID 会触发 Word“文件损坏需要修复”。保留首次出现，丢弃后续重复。
seen_ids = set()
for m in re.finditer(r'<w:style\b[^>]*>.*?</w:style>', styles, re.S):
    sid = re.search(r'w:styleId="([^"]+)"', m.group(0)).group(1)
    if sid in seen_ids:
        styles = styles.replace(m.group(0), '')
        print('模板去重: 移除重复样式', sid)
    else:
        seen_ids.add(sid)


# 1.（自动编号保留，不做修改）

# 2. Title：去 outlineLvl、去首行缩进；字体改宋体、字号小初（36pt = sz72）
block = style_block(styles, 'Title')
if block:
    nb = block
    if '<w:outlineLvl' in nb:
        nb = re.sub(r'<w:outlineLvl[^/]*/>', '', nb)
        print('Title: 已移除 outlineLvl')
    if '<w:ind ' in nb:
        nb = re.sub(r'<w:ind[^/]*/>', '', nb)
        print('Title: 已移除首行缩进')
    nb2 = re.sub(r'<w:rFonts[^/]*/>',
                 '<w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体" w:cs="宋体" />',
                 nb)
    if nb2 != nb:
        nb = nb2
        print('Title: 字体已改为宋体')
    nb2 = re.sub(r'<w:sz w:val="\d+"', '<w:sz w:val="72"', nb)
    nb2 = re.sub(r'<w:szCs w:val="\d+"', '<w:szCs w:val="72"', nb2)
    if nb2 != nb:
        nb = nb2
        print('Title: 字号已改为小初（36pt）')
    if nb != block:
        styles = styles.replace(block, nb)

# 3. 段前分页：Heading1、TOCHeading 开启；Heading2-9 显式关闭
for sid in ('Heading1', 'TOCHeading'):
    block = style_block(styles, sid)
    if block and 'pageBreakBefore' not in block:
        styles = styles.replace(block, add_page_break_before(block, True))
        print('%s: 已添加段前分页' % sid)
for i in range(2, 10):
    sid = 'Heading%d' % i
    block = style_block(styles, sid)
    if block and 'pageBreakBefore' not in block:
        styles = styles.replace(block, add_page_break_before(block, False))
        print('%s: 已显式关闭段前分页' % sid)

# 4. TOCHeading：basedOn 切断为 Normal（不继承 Heading1 的任何属性）
block = style_block(styles, 'TOCHeading')
if block:
    nb = re.sub(r'<w:basedOn[^/]*/>', '<w:basedOn w:val="Normal" />', block)
    if nb != block:
        styles = styles.replace(block, nb)
        print('TOCHeading: basedOn 已切换为 Normal')

# 5. Table 字体：仿宋_GB2312 + Times New Roman，五号（10.5pt）
block = style_block(styles, 'Table')
if block and '仿宋' not in block:
    fonts = ('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"'
             ' w:eastAsia="仿宋_GB2312" w:cs="Times New Roman" />')
    if '<w:rPr>' in block:
        # 合并进已有 rPr（新模板 Table 自带 kern/sz）
        nb = block.replace('<w:rPr>', '<w:rPr>' + fonts, 1)
    else:
        # 无 rPr 时新建（CT_Style 中 rPr 须位于 tblPr 之前）
        nb = block.replace('<w:tblPr>', '<w:rPr>' + fonts + '</w:rPr><w:tblPr>', 1)
    nb = re.sub(r'<w:sz w:val="20" />', '<w:sz w:val="21" />', nb)
    nb = re.sub(r'<w:szCs w:val="20" />', '<w:szCs w:val="21" />', nb)
    styles = styles.replace(block, nb)
    print('Table: 已设置表格字体（仿宋_GB2312/Times New Roman，五号）')

# 6. 克隆无编号附录标题样式：保留 basedOn=HeadingN 继承链（附录标题字体字号
#    与对应级别章节标题一致，如 Heading3 的黑体小四靠继承 Heading2 获得），
#    仅用 numId=0 显式关闭自动编号 —— 与模板 TOCHeading 同一套做法。
appendix_styles = ''
numoff = '<w:numPr><w:numId w:val="0" /></w:numPr>'
for i in range(1, 5):
    src = style_block(styles, 'Heading%d' % i)
    if not src:
        continue
    clone = src
    clone = clone.replace('w:styleId="Heading%d"' % i,
                          'w:styleId="AppendixHeading%d"' % i)
    clone = clone.replace('<w:name w:val="heading %d"/>' % i,
                          '<w:name w:val="Appendix Heading %d"/>' % i)
    clone = re.sub(r'<w:numPr>.*?</w:numPr>', '', clone, flags=re.S)
    clone = re.sub(r'<w:link[^/]*/>', '', clone)
    clone = re.sub(r'<w:next[^/]*/>', '', clone)
    # numPr 按 OOXML schema 顺序插入：keepNext/keepLines/pbb 之后、spacing 之前
    for anchor in ('<w:spacing', '<w:ind', '<w:jc', '<w:outlineLvl', '</w:pPr>'):
        pos = clone.find(anchor)
        if pos >= 0:
            clone = clone[:pos] + numoff + clone[pos:]
            break
    if i == 1 and 'pageBreakBefore' not in clone:
        clone = add_page_break_before(clone, True)
    appendix_styles += clone
    print('AppendixHeading%d: 已创建（继承标题字体，仅关闭编号）' % i)
if appendix_styles:
    pos = styles.rindex('</w:styles>')
    styles = styles[:pos] + appendix_styles + styles[pos:]

# 6.5 Block Text（引用块）：模板未定义西文字体，英文回落 Calibri；
#     显式指定 Times New Roman / 仿宋，与正文字体一致
m = style_block(styles, 'BlockText')
if m and '<w:rFonts' not in m.group(0):
    block = m.group(0)
    fonts = ('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"'
             ' w:cs="Times New Roman" w:eastAsia="仿宋"/>')
    if '<w:rPr>' in block:
        new = block.replace('<w:rPr>', '<w:rPr>' + fonts, 1)
    else:
        new = block.replace('</w:style>', '<w:rPr>' + fonts + '</w:rPr></w:style>')
    styles = styles.replace(block, new)
    print('Block Text: 已补充西文字体（Times New Roman/仿宋）')

# 6.6 Source Code 段落样式：添加 f8f8f8 段落底纹，使全部代码块（含未标注语言的）
#     具有统一的代码块灰底外观（run 级 VerbatimChar 底纹会触发 Word"无法读取的内容"，勿用）
m = style_block(styles, 'SourceCode')
if m and '<w:shd' not in m.group(0):
    block = m.group(0)
    shd = '<w:shd w:val="clear" w:fill="f8f8f8" />'
    if '<w:ind' in block:
        new = block.replace('<w:ind', shd + '<w:ind', 1)
    else:
        new = block.replace('</w:pPr>', shd + '</w:pPr>', 1)
    styles = styles.replace(block, new)
    print('Source Code: 已添加段落底纹（f8f8f8）')

# 6.7 Source Code 字号：五号(21) → 小四(24)，与正文一致
m = style_block(styles, 'SourceCode')
if m:
    block = m.group(0)
    new = re.sub(r'(<w:szCs? w:val=")"?21"?', r'\g<1>24', block)
    if new != block:
        styles = styles.replace(block, new)
        print('Source Code: 字号已改为小四（24）')

# 6.8 docDefaults 默认字体：五号+Calibri 主题 → Times New Roman/仿宋/小四，
#     使未挂样式的文本（列表项等）与正文字体一致
defaults = re.search(r'<w:rPrDefault><w:rPr>.*?</w:rPr></w:rPrDefault>', styles, re.S)
if defaults and 'w:asciiTheme' in defaults.group(0):
    block = defaults.group(0)
    new = block.replace('<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"'
                        ' w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>',
                        '<w:rFonts w:ascii="Times New Roman" w:eastAsia="仿宋"'
                        ' w:hAnsi="Times New Roman" w:cs="Times New Roman"/>')
    new = re.sub(r'<w:sz w:val="21"', '<w:sz w:val="24"', new)
    new = re.sub(r'<w:szCs w:val="22"', '<w:szCs w:val="24"', new)
    styles = styles.replace(block, new)
    print('docDefaults: 已改为 Times New Roman/仿宋/小四')

items['word/styles.xml'] = styles.encode('utf-8')

# 7. settings.xml: updateFields
settings = items['word/settings.xml'].decode('utf-8')
if 'updateFields' not in settings:
    settings = re.sub(r'(<w:settings[^>]*>)',
                      r'\1<w:updateFields w:val="true" />',
                      settings, count=1)
    items['word/settings.xml'] = settings.encode('utf-8')
    print('settings.xml: 已添加 updateFields')

shutil.copy(PATH, PATH + '.orig')
with zipfile.ZipFile(PATH, 'w', zipfile.ZIP_DEFLATED) as zout:
    for name, data in items.items():
        zout.writestr(name, data)
print('补丁完成（原始文件备份为 reference.docx.orig）')
