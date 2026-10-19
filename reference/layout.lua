--[[
layout – 封面、目录、标题编号、表格列宽与分隔线的排版处理。

配合 reference/reference.docx 中 patch-reference.py 打过的样式补丁使用：
- 文档第一个标题 → Title 样式（封面大标题，前置留白）
- 标题手写编号（“1.1 编写目的”）→ 剥离，由 Heading 样式的自动编号接管
- 附录节标题（“附录”及其子节）→ AppendixHeading1-4 无编号样式
- 封面信息表 → 去掉表头、首列加粗；其后插入目录域
- 水平分隔线（---）→ 丢弃（章与章之间由 Heading1 的段前分页取代）
- 所有表格 → 列宽按内容显示宽度比例分配（均分改为按内容）
]]
local pandoc = require 'pandoc'
local utils = require 'pandoc.utils'
local stringify = utils.stringify

-- 目录域 XML：与 pandoc --toc 生成的结构一致（无 \h，目录条目不做超链接）。
-- settings.xml 中的 updateFields 会让 Word/WPS 打开时自动填充目录。
local TOC_FIELD_XML = [[<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents" /><w:docPartUnique /></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:pPr><w:pStyle w:val="TOCHeading" /></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia" /></w:rPr><w:t xml:space="preserve">目录</w:t></w:r></w:p><w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true" /><w:instrText xml:space="preserve">TOC \o &quot;1-3&quot; \z \u</w:instrText><w:fldChar w:fldCharType="separate" /><w:fldChar w:fldCharType="end" /></w:r></w:p></w:sdtContent></w:sdt>]]

-- 手动分页
local PAGE_BREAK_XML = [[<w:p><w:r><w:br w:type="page" /></w:r></w:p>]]

-- 图片段落样式（模板中的“图”样式：居中、自动行距，图片不会被裁切/遮挡）
local IMAGE_STYLE = '图'

-- Word 原生题注域：label+章节号(STYLEREF)+序号(SEQ)，写入缓存结果以便
-- 未更新域时直接显示；Word/WPS 打开时经 updateFields 自动重算。
-- style_id 必须是模板中的样式 ID（RawBlock 不做样式名翻译）。
local function field_xml (instr, cached)
  return '<w:r><w:fldChar w:fldCharType="begin" /></w:r>'
    .. '<w:r><w:instrText xml:space="preserve"> ' .. instr .. ' </w:instrText></w:r>'
    .. '<w:r><w:fldChar w:fldCharType="separate" /></w:r>'
    .. '<w:r><w:t>' .. cached .. '</w:t></w:r>'
    .. '<w:r><w:fldChar w:fldCharType="end" /></w:r>'
end

local function xml_escape (s)
  return (s:gsub('&', '&amp;'):gsub('<', '&lt;'):gsub('>', '&gt;'))
end

local function caption_block (label, chapter, seq, style_id, text)
  local body
  if chapter then
    -- 章节编号格式：图 2-1 / 表 3-2
    body = '<w:r><w:rPr><w:rFonts w:hint="eastAsia" /></w:rPr>'
      .. '<w:t xml:space="preserve">' .. label .. ' </w:t></w:r>'
      .. field_xml('STYLEREF 1 \\s', tostring(chapter))
      .. '<w:r><w:t>-</w:t></w:r>'
      .. field_xml('SEQ ' .. label .. ' \\* ARABIC \\s 1', tostring(seq))
  else
    -- 附录格式：附表 1（全局序号，无章节前缀）
    body = '<w:r><w:rPr><w:rFonts w:hint="eastAsia" /></w:rPr>'
      .. '<w:t xml:space="preserve">' .. label .. ' </w:t></w:r>'
      .. field_xml('SEQ ' .. label .. ' \\* ARABIC', tostring(seq))
  end
  if text and text ~= '' then
    body = body .. '<w:r><w:t xml:space="preserve"> ' .. xml_escape(text) .. '</w:t></w:r>'
  end
  return pandoc.RawBlock('openxml',
    '<w:p><w:pPr><w:pStyle w:val="' .. style_id .. '" /></w:pPr>' .. body .. '</w:p>')
end

-- 全角/CJK 字符按 2 个单位计宽
local function display_width (s)
  local w = 0
  for _, code in utf8.codes(s) do
    if (code >= 0x2e80 and code <= 0x9fff)
      or (code >= 0xff00 and code <= 0xffef) then
      w = w + 2
    else
      w = w + 1
    end
  end
  return w
end

-- 列宽分配（水位模型）：
-- 1. 每列“期望宽度”= 最长内容单行放下，上限 55%（防止超长列独占）
-- 2. 硬下限 = max(min_w, 短表头单行需求)——表头 ≤5 个全角字的列保证表头完整
-- 3. 期望总和 ≤ 版宽：多余空间按期望比例分摊（所有列保持单行）
--    期望总和 > 版宽：从各列“期望−下限”按比例收缩，任何列不破下限
local function apply_content_widths (tbl, min_w)
  local ncol = #tbl.colspecs
  if ncol == 0 then
    return
  end
  min_w = min_w or 0.06
  local PAGE_PT = 415   -- A4 版面宽度（pt）
  local U_PT = 5.25     -- 五号字半宽字符（pt）
  local PAD_PT = 16     -- 单元格左右内边距（pt，含安全余量）

  local maxw = {}
  for i = 1, ncol do
    maxw[i] = 2
  end
  local function scan_row (row)
    for i, cell in ipairs(row.cells) do
      if i <= ncol then
        local w = display_width(stringify(cell.content))
        if w > maxw[i] then
          maxw[i] = w
        end
      end
    end
  end
  for _, row in ipairs(tbl.head.rows) do scan_row(row) end
  for _, body in ipairs(tbl.bodies) do
    for _, row in ipairs(body.body) do scan_row(row) end
  end
  if tbl.foot then
    for _, row in ipairs(tbl.foot.rows) do scan_row(row) end
  end

  -- 短表头（≤5 个全角字）单行显示所需占比：宽×6pt + 16pt 内边距
  local mins, wants = {}, {}
  local sum_wants = 0
  for i = 1, ncol do
    local need = 0
    local hw = 0
    for _, row in ipairs(tbl.head.rows) do
      local cell = row.cells[i]
      if cell then
        local w = display_width(stringify(cell.content))
        if w > hw then
          hw = w
        end
      end
    end
    if hw > 0 and hw <= 10 then
      need = (hw * 6 + 20) / PAGE_PT
    end
    mins[i] = math.max(min_w, need)
    wants[i] = math.min((maxw[i] * U_PT + PAD_PT) / PAGE_PT, 0.55)
    if wants[i] < mins[i] then
      wants[i] = mins[i]
    end
    sum_wants = sum_wants + wants[i]
  end

  local fracs = {}
  if sum_wants <= 1 then
    -- 所有列都能单行放下：多余空间按期望比例分摊
    for i = 1, ncol do
      fracs[i] = wants[i] + wants[i] * (1 - sum_wants) / sum_wants
    end
  else
    -- 超出版宽：按“期望 − 下限”的比例收缩（水位法），下限刚性保留
    local excess = sum_wants - 1
    local shrinkable = 0
    for i = 1, ncol do
      shrinkable = shrinkable + (wants[i] - mins[i])
    end
    local k = 1
    if shrinkable > 0 then
      k = math.max(0, 1 - excess / shrinkable)
    end
    for i = 1, ncol do
      fracs[i] = mins[i] + (wants[i] - mins[i]) * k
    end
  end

  local sum = 0
  for i = 1, ncol do
    sum = sum + fracs[i]
  end
  local specs = pandoc.List()
  for i = 1, ncol do
    specs:insert({tbl.colspecs[i][1], fracs[i] / sum})
  end
  tbl.colspecs = specs
end

-- 表格内不使用等宽字体：行内代码（`...`）还原为普通文本，
-- 仅独立代码块保留 Source Code 样式
local function strip_table_code (tbl)
  local function plain (content)
    return content:walk{
      Code = function (c)
        return pandoc.Str(c.text)
      end,
    }
  end
  local function new_row (row)
    local cells = pandoc.List()
    for _, c in ipairs(row.cells) do
      cells:insert(pandoc.Cell(plain(c.content)))
    end
    return pandoc.Row(cells)
  end
  for i, row in ipairs(tbl.head.rows) do
    tbl.head.rows[i] = new_row(row)
  end
  for _, body in ipairs(tbl.bodies) do
    for i, row in ipairs(body.body) do
      body.body[i] = new_row(row)
    end
  end
  if tbl.foot then
    for i, row in ipairs(tbl.foot.rows) do
      tbl.foot.rows[i] = new_row(row)
    end
  end
end

-- 封面信息表：去掉表头行，首列加粗，标签列窄、内容列宽
local function make_cover_table (tbl)
  strip_table_code(tbl)
  tbl.head = pandoc.TableHead{}
  for _, body in ipairs(tbl.bodies) do
    for ri, row in ipairs(body.body) do
      if #row.cells > 0 then
        local cells = pandoc.List()
        cells:insert(pandoc.Cell{pandoc.Plain{pandoc.Strong(
          pandoc.utils.blocks_to_inlines(row.cells[1].content))}})
        for ci = 2, #row.cells do
          cells:insert(row.cells[ci])
        end
        body.body[ri] = pandoc.Row(cells)
      end
    end
  end
  apply_content_widths(tbl, 0.18)
end

-- 剥离标题中的手写编号（“1.1 编写目的” → “编写目的”），
-- 编号由 reference.docx 中 Heading 样式的多级自动编号提供。
-- 行内代码一律还原为正文字体（与表格内代码的处理保持一致）；
-- 围栏代码块（无论有无语言标记）保持代码样式：带标记的由 writer 做 tango 语法着色，
-- 无标记的为纯等宽代码段落。
function Code (el)
  return pandoc.Str(el.text)
end

function Header (h)
  local inl = h.content
  if #inl >= 2
    and inl[1].t == 'Str'
    and inl[1].text:match('^%d+[%d%.]*$')
    and inl[2].t == 'Space'
  then
    local rest = pandoc.List()
    for i = 3, #inl do
      rest:insert(inl[i])
    end
    h.content = rest
    return h
  end
end

-- 独立图片段落 → 套用“图”样式（居中、自动行距）
-- 返回 图片块列表, 题注文字（来自 cap-text 属性或 Figure caption）
local function style_image_block (block)
  local img
  if block.t == 'Plain' and #block.content == 1 and block.content[1].t == 'Image' then
    img = block.content[1]
  elseif block.t == 'Figure' then
    for _, b in ipairs(block.content) do
      if b.t == 'Plain' then
        for _, inl in ipairs(b.content) do
          if inl.t == 'Image' then img = inl end
        end
      end
    end
  end
  if not img then
    return nil, nil
  end
  local blocks = pandoc.List{
    pandoc.Div({pandoc.Para{img}}, pandoc.Attr('', {}, {{'custom-style', IMAGE_STYLE}}))
  }
  -- 题注文字：优先取 cap-text 属性（diagram.lua 写入），Figure caption 兜底
  local text = img.attributes['cap-text']
  if (not text or text == '') and block.t == 'Figure' then
    local ok, s = pcall(stringify, block.caption)
    if ok then
      text = s
    end
  end
  return blocks, text
end

-- 封面标题分行：在破折号处分两行居中显示（小初字号下单行放不下，
-- 自然折行会断在词中间；去掉破折号按“文档 / 设计方案”两行更符合封面惯例）
local function split_title (inlines)
  local out = pandoc.List()
  local skip_space = false
  for _, it in ipairs(inlines) do
    if it.t == 'Str' and it.text == '—' then
      while #out > 0 and out[#out].t == 'Space' do
        table.remove(out)
      end
      out:insert(pandoc.LineBreak())
      skip_space = true
    elseif skip_space and it.t == 'Space' then
      skip_space = false
    else
      out:insert(it)
      skip_space = false
    end
  end
  return out
end

function Pandoc (doc)
  local out = pandoc.List()
  local title_done = false
  local info_table_pending = false
  local in_appendix = false
  local chapter = 0
  local fig_seq, tbl_seq = {}, {}
  local appendix_tbl_seq = 0
  local table_caption_pending = nil

  local function semantic_table_caption(block)
    if block.t ~= 'Para' or #block.content ~= 1 then
      return nil
    end
    local item = block.content[1]
    if item.t ~= 'Strong' then
      return nil
    end
    local text = stringify(item)
    if text == '' then
      return nil
    end
    return text
  end

  for _, block in ipairs(doc.blocks) do
    local table_caption_text = nil
    local skip_block = false

    -- 先处理前一块暂存的语义表题：只有紧邻的 Table 才消费为题注，
    -- 否则恢复原段落，避免误吞正文；恢复后当前块仍可参与表题识别
    -- （连续两个加粗段落时，前一段是正文引导语，后一段才是表题）。
    if table_caption_pending then
      if block.t == 'Table' and not info_table_pending then
        table_caption_text = table_caption_pending.text
        table_caption_pending = nil
      else
        out:insert(table_caption_pending.block)
        table_caption_pending = nil
      end
    end

    -- 封面信息表之前的内容不能被识别为普通表题；其余独立加粗段落先暂存。
    if not table_caption_pending
      and not table_caption_text
      and not info_table_pending
      and block.t == 'Para'
    then
      local text = semantic_table_caption(block)
      if text then
        table_caption_pending = {block = block, text = text}
        skip_block = true
      end
    end

    if skip_block then
      -- 语义表题已暂存，当前轮完全不输出；后续非邻接块会由 pending 恢复一次。

    elseif block.t == 'HorizontalRule' then
      -- 丢弃分隔线

    elseif not title_done and block.t == 'Header' then
      -- 封面页：留白 + 大标题，随后另起目录页
      title_done = true
      info_table_pending = true
      for _ = 1, 4 do
        out:insert(pandoc.Para{pandoc.Str ''})
      end
      out:insert(pandoc.Div(
        {pandoc.Para(split_title(block.content))},
        pandoc.Attr('', {}, {{'custom-style', 'Title'}})
      ))
      out:insert(pandoc.RawBlock('openxml', TOC_FIELD_XML))
      out:insert(pandoc.RawBlock('openxml', PAGE_BREAK_XML))

    elseif info_table_pending and block.t == 'Table' then
      -- 文档信息表：正文开头（目录页之后），去表头、首列加粗，不编号
      info_table_pending = false
      make_cover_table(block)
      out:insert(block)

    elseif block.t == 'Plain' or block.t == 'Figure' then
      -- 独立图片 → “图”样式 + 图题（图下方，“图 N + 说明”同行）
      local img_blocks, cap_text = style_image_block(block)
      if img_blocks then
        fig_seq[chapter] = (fig_seq[chapter] or 0) + 1
        out:extend(img_blocks)
        out:insert(caption_block('图', chapter, fig_seq[chapter], 'ImageCaption', cap_text))
      else
        out:insert(block)
      end

    elseif block.t == 'Header' then
      -- 附录节（以“附录”开头的一级标题及其子标题）使用无编号样式；
      -- 注意 docx writer 会忽略 Header 上的 custom-style，须转为 Div+Para。
      -- custom-style 按样式名匹配（“Appendix Heading 1”，带空格），
      -- 若写成样式 ID pandoc 查不到会自动补一个默认样式的同名 ID 副本，
      -- 造成重复 styleId 触发 Word 修复。
      if in_appendix
        or (block.level == 1 and stringify(block.content):match('^附录'))
      then
        local appendix_start = not in_appendix and block.level == 1
        in_appendix = true
        -- 显式分页兜底（AppendixHeading1 样式自带段前分页，双保险且不产生空页）
        if appendix_start then
          out:insert(pandoc.RawBlock('openxml', PAGE_BREAK_XML))
        end
        out:insert(pandoc.Div(
          {pandoc.Para(block.content)},
          pandoc.Attr('', {},
            {{'custom-style', 'Appendix Heading ' .. math.min(block.level, 4)}})
        ))
      else
        -- 封面占用源 Markdown 的第一个 #；正文标题整体提升一级：
        -- ## -> Heading1，### -> Heading2，依此类推。
        -- 这样题注中的 STYLEREF 1 \\s 才能找到正文一级标题。
        local output_level = math.max(1, block.level - 1)
        block.level = output_level
        if output_level == 1 then
          chapter = chapter + 1
        end
        out:insert(block)
      end

    else
      if block.t == 'Table' then
        strip_table_code(block)
        apply_content_widths(block)
        -- 表题（表上方）；附录表格用独立“附表”序列
        if in_appendix then
          appendix_tbl_seq = appendix_tbl_seq + 1
          out:insert(caption_block('附表', nil, appendix_tbl_seq, 'TableCaption', table_caption_text))
        else
          tbl_seq[chapter] = (tbl_seq[chapter] or 0) + 1
          out:insert(caption_block('表', chapter, tbl_seq[chapter], 'TableCaption', table_caption_text))
        end
      end
      out:insert(block)
    end
  end

  -- 文档以加粗段落结尾时，没有后续表格可消费，恢复为普通正文。
  if table_caption_pending then
    out:insert(table_caption_pending.block)
  end

  doc.blocks = out
  return doc
end
