# md2docx：Markdown 转 DOCX 工具链

基于 pandoc 和 Word 参考模板，把结构化 Markdown 批量转换为排版规范的 DOCX。支持 Mermaid 图渲染、封面与目录生成、标题自动编号、表格列宽分配、图表题注、LibreOffice 域烘焙等。

## 特性

- 全局样式由 `reference/reference.docx` 控制，正文格式与模板一致；
- Mermaid 代码块自动渲染为 PNG，并生成「图 N-M」题注；
- 首个一级标题作为封面，紧随其后的表格作为文档信息表；
- 标题编号剥离后由 Word 多级编号接管，附录自动切换为无编号样式；
- 表格按内容水位分配列宽，短表头保证单行显示；
- 构建后用 LibreOffice 无头刷新目录与域，Word/WPS 打开即显示目录；
- 增量转换，只处理比产物新的 Markdown；
- 附带模板去敏感化、深度清理和格式审计脚本。

## 目录结构

```text
build/          Markdown 源文件
docx-build/     生成的 DOCX，目录结构与 build/ 镜像
reference/      Word 模板、Lua 过滤器与 Python 后处理
tools/          格式审计、Mermaid 导出等辅助脚本
build.sh        构建入口
md2docx.sh      通用批量转换器
collect-docx.sh 校验镜像关系并打包
```

`build/` 与 `docx-build/` 均不纳入版本控制。转换规则是 `build/<相对路径>/x.md` → `docx-build/<相对路径>/x.docx`。

## 依赖

- pandoc ≥ 3.0
- mermaid-cli（提供 `mmdc` 命令）
- LibreOffice（提供 `soffice`，用于域烘焙）
- Python 3.8+，以及 `lxml`、`Pillow`
- 系统中可用的中文字体

## 快速开始

```bash
# 把 Markdown 放进 build/ 的任意子目录
mkdir -p build/example
cp your-document.md build/example/

# 全量增量转换
./build.sh

# 只转换指定文件
./md2docx.sh build/example/your-document.md

# 忽略增量、全部重转
./md2docx.sh --force

# 跳过 LibreOffice 域烘焙
./md2docx.sh --no-bake

# 校验产物完整性并打包
./collect-docx.sh
```

`build.sh` 是 `md2docx.sh` 的薄封装；需要传给转换器的参数可直接透传。

## Markdown 约定

| Markdown 写法 | DOCX 效果 |
|---|---|
| 文档首个 `# 标题` | 封面大标题，独占一页 |
| 紧随封面的表格 | 文档信息表（无表头、首列加粗） |
| `# 1 引言` 等带编号标题 | 剥离手写编号，改用 Word 自动多级编号 |
| 每个一级标题 | 自动另起一页 |
| `# 附录` 之后的标题 | 切换为无编号的附录样式 |
| `---` | 丢弃，由章前分页取代 |
| ` ```mermaid ` 代码块 | 渲染为 PNG，套用「图」样式并加题注 |
| 普通表格 | 自动分配列宽 |
| 行内代码 | 还原为正文文本 |
| 围栏代码块 | 保持等宽字体，按语言着色 |

Mermaid 题注可在代码块首行写 `%%| fig-cap: 说明文字`。完整输入规范见 [MD格式约束.md](MD格式约束.md)。

## 排版实现

- `reference/layout.lua`：封面、目录、标题编号、表格列宽；
- `reference/diagram.lua`：Mermaid 调用与图片插入；
- `reference/postprocess.py`：图片浮动化与章节收尾；
- `reference/bake_fields.py`：调用 LibreOffice 刷新目录和域，并剥离 `updateFields` 标志；
- `reference/patch-reference.py`：为模板补充标题分页、附录样式等定义；
- `reference/desensitize-reference.py`：把模板正文替换为格式描述；
- `reference/deep-clean-reference.py`：替换模板图片为占位图、剥离嵌入对象、重写文档属性。

## 准备或更换模板

仓库自带的 `reference/reference.docx` 已完成补丁与清理。更换模板时依次执行：

```bash
python3 reference/patch-reference.py
python3 reference/desensitize-reference.py reference/reference.docx
python3 reference/deep-clean-reference.py reference/reference.docx
```

`desensitize-reference.py` 与 `deep-clean-reference.py` 会清除模板原有的正文、图片和嵌入对象，避免把模板来源内容带进产物。若只调整样式，不要直接改 `reference.docx`，应修改 `patch-reference.py` 后重新执行上述流程。

## 常见问题

| 现象 | 原因与处理 |
|---|---|
| 目录空白 | 该产物未烘焙或为旧版本，重跑 `./build.sh` |
| 打开时提示更新域 | 产物带有 `updateFields` 标志，重跑烘焙流程 |
| 标题编号重复 | 模板补丁丢失，重新执行模板准备流程 |
| Mermaid 渲染失败 | 检查 `mmdc` 是否可用、临时目录权限是否正常 |
| 表格列宽均分 | 确认两个 Lua 过滤器都已挂载 |
| 图片不可见或无法编辑 | 使用 `tools/mermaid_svg_export.py` 导出可编辑 SVG |

## 辅助脚本

- `tools/md_audit.py`：按格式约束扫描全部 Markdown；
- `tools/md_fix_mech.py`：修复标题层级、Emoji、代码语言标识等机械问题；
- `tools/md_insert_caps.py`：批量插入表格题注；
- `tools/md_style_audit.py`：排查口语化表述；
- `tools/mermaid_fit_audit.py`：检查 Mermaid 图的尺寸与有效字号；
- `tools/mermaid_svg_export.py`：导出可在 Inkscape/LibreOffice 中编辑的 SVG。
