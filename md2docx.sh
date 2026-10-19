#!/usr/bin/env bash
# 通用 MD → DOCX 批量转换器（复用 reference/reference.docx 模板）
#
# 目录约定（重要）：
#   build/      仅存放 Markdown 源（不再有任何 docx）；
#   docx-build/ 仅存放生成的 Word 文档，目录结构与 build/ 下的 md 一一镜像。
#
# 用法：
#   ./md2docx.sh                     转换全部 md（增量：仅转比 docx 新的）
#   ./md2docx.sh --force             忽略增量，全部重转
#   ./md2docx.sh --overwrite         连已存在的同名非本工具 docx 也覆盖（谨慎）
#   ./md2docx.sh --no-bake           跳过域烘焙（默认烘焙，见下）
#   ./md2docx.sh build/xx/yy.md ...   只转换指定 md
#
# 输出位置规则：build/<相对路径>/x.md → docx-build/<相同相对路径>/x.docx
#   （build/<目录>/x.md → docx-build/<目录>/x.docx）
# 全量模式下顺带清理 docx-build/ 中已无对应 md 的孤立 docx。
# 已有同名 docx 且非本工具生成时自动跳过，避免覆盖原件。
#
# 域烘焙（默认开启）：转换后由 reference/bake_fields.py 用 LibreOffice 无头
# 刷新目录与域并固化结果、剥离 updateFields 标志——Word/WPS 打开即显示目录，
# 不再弹「是否更新域」。页码按 LO 排版计算，与 Word 可有 ±1 页出入；
# 在 Word 里手改文档后目录不会自动变，需手动更新（选中目录按 F9）。
set -euo pipefail
cd "$(dirname "$0")"

REF="reference/reference.docx"
BUILD_DIR="build"
OUT_DIR="docx-build"
MANIFEST=".md2docx.manifest"    # 记录本工具生成的 docx，用于增量与保护判断
FORCE=0 OVERWRITE=0 BAKE=1 EXPLICIT=0 FILES=()
BAKED=()                        # 本轮成功转换、待烘焙的 docx

for a in "$@"; do
  case "$a" in
    --force) FORCE=1 ;;
    --overwrite) OVERWRITE=1 ;;
    --no-bake) BAKE=0 ;;
    *) FILES+=("$a"); EXPLICIT=1 ;;
  esac
done
touch "$MANIFEST"

convert_one() {
  local rel="$1"
  rel="${rel#./}"
  local out="$OUT_DIR/${rel#"$BUILD_DIR"/}"
  out="${out%.md}.docx"
  mkdir -p "$(dirname "$out")"

  if grep -qxF "$out" "$MANIFEST"; then
    if [[ $EXPLICIT -eq 0 && $FORCE -eq 0 && "$out" -nt "$rel" ]]; then
      echo "跳过（已是最新） $out"
      return 0
    fi
  elif [[ -e "$out" && $OVERWRITE -eq 0 && $EXPLICIT -eq 1 ]]; then
    # 仅显式模式保护非本工具生成的同名 docx；
    # 全量模式下 docx-build/ 完全由本工具维护，过期即重转
    echo "跳过（已有同名原始 docx） $out"
    return 0
  fi

  echo "转换 $rel → $out"
  pandoc "$rel" -o "$out" \
    --reference-doc="$REF" \
    --resource-path="$(dirname "$rel"):$PWD" \
    --lua-filter=reference/diagram.lua \
    --lua-filter=reference/layout.lua \
    --syntax-highlighting=tango
  python3 reference/postprocess.py "$out"   # 图片浮动化 + 紧凑章节（H1 去分页）
  if [ "$BAKE" -eq 1 ]; then BAKED+=("$out"); fi
  { grep -vxF "$out" "$MANIFEST" || true; } > "$MANIFEST.tmp"
  mv "$MANIFEST.tmp" "$MANIFEST"
  echo "$out" >> "$MANIFEST"
}

bake_all() {
  if [ "$BAKE" -eq 0 ] || [ "${#BAKED[@]}" -eq 0 ]; then return 0; fi
  echo "== 域烘焙（LibreOffice 无头刷新目录，Word 打开不再弹更新域）=="
  python3 reference/bake_fields.py "${BAKED[@]}"
}

if [ "$EXPLICIT" -eq 1 ]; then
  for f in "${FILES[@]}"; do convert_one "$f"; done
  bake_all
else
  # 全量模式：扫描 build/ 全部 md；总体设计方案走 build.sh 专用排版，不在此列
  while IFS= read -r f; do convert_one "$f"; done < <(
    find "$BUILD_DIR" -name '*.md' | sort
  )

  # 清理 docx-build/ 中已无对应 md 的孤立 docx，保持两目录镜像
  while IFS= read -r -d '' out; do
    rel="${out#"$OUT_DIR"/}"
    md_path="$BUILD_DIR/${rel%.docx}.md"
    if [ ! -f "$md_path" ]; then
      echo "清理无源 docx: $out"
      rm "$out"
    fi
  done < <(find "$OUT_DIR" -name '*.docx' -print0 2>/dev/null || true)
  find "$OUT_DIR" -type d -empty -delete 2>/dev/null || true

  # 同步清理产物清单中已不存在的条目
  if [ -s "$MANIFEST" ]; then
    while IFS= read -r entry; do
      [ -f "$entry" ] && echo "$entry"
    done < "$MANIFEST" > "$MANIFEST.tmp"
    mv "$MANIFEST.tmp" "$MANIFEST"
  fi

  bake_all
fi
