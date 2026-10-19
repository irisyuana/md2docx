#!/usr/bin/env bash
# 构建入口：把 build/ 下的 Markdown 批量转换为 docx-build/ 下的 Word 文档。
#
# 用法：
#   ./build.sh                      全量增量转换
#   ./build.sh --force              忽略增量，全部重转
#   ./build.sh --no-bake            跳过域烘焙
#   ./build.sh build/xx/yy.md ...   只转换指定文件
#
# 实际转换由 md2docx.sh 完成，本脚本仅做统一入口。
set -euo pipefail
cd "$(dirname "$0")"

exec ./md2docx.sh "$@"
