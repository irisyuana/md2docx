#!/usr/bin/env bash
# 校验 docx-build/ 与 build/ 的镜像关系，并打包交付压缩包 docx.tar.gz
#
# 背景：build/ 只存 Markdown 源，docx-build/ 由 md2docx.sh / build.sh 直接生成
# Word 文档（目录结构与 build/ 下的 md 镜像）。本脚本做两件事：
#   1. 一致性校验：docx-build/ 中每个 docx 在 build/ 都有同构路径的 md 源，
#      build/ 中每个 md 在 docx-build/ 都有同名 docx 产物（缺失仅告警，不阻断）；
#   2. 打包：docx-build/ → docx.tar.gz（先删旧包再重建，保证纯净）。
#
# 用法：
#   ./collect-docx.sh                # 校验 + 打包
set -euo pipefail
cd "$(dirname "$0")"

SRC="build"
DST="docx-build"

if [ ! -d "$DST" ]; then
  echo "错误：$DST/ 不存在，请先运行 ./build.sh 生成 Word 文档" >&2
  exit 1
fi

missing_out=0
while IFS= read -r -d '' md; do
  rel="${md#"$SRC"/}"
  out="$DST/${rel%.md}.docx"
  if [ ! -f "$out" ]; then
    echo "告警：缺产物 $out"
    missing_out=$((missing_out + 1))
  fi
done < <(find "$SRC" -name '*.md' -print0 | sort -z)

orphan=0
while IFS= read -r -d '' out; do
  rel="${out#"$DST"/}"
  md_path="$SRC/${rel%.docx}.md"
  if [ ! -f "$md_path" ]; then
    echo "告警：孤立产物（无 md 源）$out"
    orphan=$((orphan + 1))
  fi
done < <(find "$DST" -name '*.docx' -print0 | sort -z)

total=$(find "$DST" -name '*.docx' | wc -l)
echo "校验完成：$DST/ 共 $total 个 docx；缺产物 $missing_out 个，孤立产物 $orphan 个"

# build/ 内不允许再出现 docx（防止旧产物回流）
stray=$(find "$SRC" -name '*.docx' | head -5)
if [ -n "$stray" ]; then
  echo "告警：$SRC/ 内发现 docx（应只存 md）："
  echo "$stray"
fi

rm -f docx.tar.gz
tar -czf docx.tar.gz "$DST/"
echo "已打包：docx.tar.gz（$(du -h docx.tar.gz | cut -f1)）"
