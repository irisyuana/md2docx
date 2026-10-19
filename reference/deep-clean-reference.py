#!/usr/bin/env python3
"""reference.docx 深度清理（在 desensitize-reference.py 之后执行）：
  - word/media/*   → 同扩展名 1x1 占位图（png/jpeg 用 PIL；emf/wmf 用最小空图元文件）
  - word/embeddings/* → 置空（剥离嵌入的 Visio 源文件与 OLE 对象）
  - docProps/app.xml  → 公司名改为指定值，原模板文档的统计信息归零

注意：media 与 app.xml 会被 pandoc 原样拷入产物，但产物正文并不引用这些图片
（只引用 pandoc 自行渲染的 mermaid PNG），因此产物打开后视觉完全一致，
仅 zip 字节与体积变化（瘦身约 3MB）。
"""
import io
import struct
import sys
import zipfile

from PIL import Image

COMPANY = "irisyuana"


def make_placeholders():
    ph = {}
    buf = io.BytesIO()
    Image.new("RGB", (1, 1), (255, 255, 255)).save(buf, "PNG")
    ph[".png"] = buf.getvalue()
    buf = io.BytesIO()
    Image.new("RGB", (1, 1), (255, 255, 255)).save(buf, "JPEG", quality=50)
    ph[".jpeg"] = ph[".jpg"] = buf.getvalue()
    # 最小空 EMF：仅一个 EMR_HEADER 记录
    ph[".emf"] = struct.pack(
        "<II4i4iIIIIHHIII4i",
        1, 88,                      # iType=EMR_HEADER, nSize
        0, 0, 0, 0,                 # rclBounds
        0, 0, 0, 0,                 # rclFrame
        0x464D4520, 0x00010000,     # " EMF", version
        88, 1,                      # nBytes, nRecords
        1, 0,                       # nHandles, reserved
        0, 0, 0,                    # nDescription, offDescription, nPalEntries
        1920, 1080, 508, 285,       # szlDevice, szlMillimeters
    )
    # 最小空 WMF：标准头 + EOF 记录
    ph[".wmf"] = (struct.pack("<HHHIHIH", 2, 9, 0x0300, 12, 0, 3, 0)
                  + struct.pack("<IH", 3, 0))
    return ph


APP_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
    'extended-properties" xmlns:vt="http://schemas.openxmlformats.org/'
    'officeDocument/2006/docPropsVTypes">'
    "<Template>Normal.dotm</Template>"
    "<TotalTime>0</TotalTime><Pages>1</Pages><Words>0</Words>"
    "<Characters>0</Characters>"
    "<Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity>"
    "<Lines>0</Lines><Paragraphs>0</Paragraphs><ScaleCrop>false</ScaleCrop>"
    f"<Company>{COMPANY}</Company><LinksUpToDate>false</LinksUpToDate>"
    "<CharactersWithSpaces>0</CharactersWithSpaces><SharedDoc>false</SharedDoc>"
    "<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion>"
    "</Properties>"
)


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else "reference.docx"
    ph = make_placeholders()
    n_media = n_emb = 0
    with zipfile.ZipFile(path) as z:
        contents = {n: z.read(n) for n in z.namelist()}
        infos = z.infolist()

    for info in infos:
        n = info.filename
        if n.startswith("word/media/"):
            ext = "." + n.rsplit(".", 1)[-1].lower()
            if ext not in ph:
                raise SystemExit(f"未知图片格式，终止以免误清理: {n}")
            contents[n] = ph[ext]
            n_media += 1
        elif n.startswith("word/embeddings/"):
            contents[n] = b""
            n_emb += 1
        elif n == "docProps/app.xml":
            contents[n] = APP_XML.encode("utf-8")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info in infos:
            z.writestr(info, contents[info.filename])

    print(f"占位图替换: {n_media} 张；嵌入对象置空: {n_emb} 个；app.xml 已重写（Company={COMPANY}）")


if __name__ == "__main__":
    main()
