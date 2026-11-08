--[[
looseleaf.lua - pandoc Lua filter: OneNote .docx -> Obsidian-friendly Markdown.

Part of Looseleaf (GPL-3.0). Implements the content rules in docs/DESIGN.md section 6.5.
Requires pandoc 3.x. Run after ConvertTo-LLCodeMarkedDocx (rule C2).

Options are passed as metadata (-M key=value). Defaults in brackets.
  looseleaf-image-prefix        base name for images                  [image]
  looseleaf-media-dir           folder to write images into           [] (images not written)
  looseleaf-box-tables          Auto|CodeBlock|Callout|Unwrap|Table   [Auto]       C1
  looseleaf-box-code-threshold  share of monospace text for code      [0.6]        C1
  looseleaf-indented            Unindent|Quote                        [Unindent]   C4
  looseleaf-glyph-bullets       true|false                            [true]       C5
  looseleaf-underline           Html|Drop                             [Html]       C8
  looseleaf-clip-sources        Normalize|Keep                        [Normalize]  C9
  looseleaf-image-width         Obsidian|Drop                         [Obsidian]   C10
  looseleaf-max-image-width     pixels                                [1000]       C10

Images are written by this filter (not by --extract-media) and referenced with the
placeholder 'LLMEDIA:<name>'. The caller replaces it with the final link (LinkStyle).
]]

local stringify = pandoc.utils.stringify

local NBSP = '\u{00A0}'
local MEDIA_TOKEN = 'LLMEDIA:'

local opt = {
  ['image-prefix'] = 'image',
  ['media-dir'] = '',
  ['box-tables'] = 'Auto',
  ['box-code-threshold'] = '0.6',
  ['indented'] = 'Unindent',
  ['glyph-bullets'] = 'true',
  ['underline'] = 'Html',
  ['clip-sources'] = 'Normalize',
  ['image-width'] = 'Obsidian',
  ['max-image-width'] = '1000',
}

local function read_options(meta)
  for key, _ in pairs(opt) do
    local value = meta['looseleaf-' .. key]
    if type(value) == 'boolean' then
      opt[key] = tostring(value) -- -M key=true arrives as a YAML boolean
    elseif value ~= nil then
      opt[key] = stringify(value)
    end
  end
end

-- ---------------------------------------------------------------------------
-- Helpers
-- ---------------------------------------------------------------------------

local function is_blank(text)
  return (text:gsub(NBSP, ''):gsub('%s', '')) == ''
end

-- Plain text of inlines; line breaks become newlines.
local function inlines_text(inlines)
  local out = {}
  for _, il in ipairs(inlines) do
    local t = il.t
    if t == 'Str' then
      out[#out + 1] = il.text
    elseif t == 'Space' then
      out[#out + 1] = ' '
    elseif t == 'SoftBreak' or t == 'LineBreak' then
      out[#out + 1] = '\n'
    elseif t == 'Code' or t == 'Math' or t == 'RawInline' then
      out[#out + 1] = il.text
    elseif il.content then
      out[#out + 1] = inlines_text(il.content)
    end
  end
  return table.concat(out)
end

-- Plain text of blocks; one line per paragraph or code line.
local function blocks_text(blocks)
  local parts = {}
  for _, b in ipairs(blocks) do
    if b.t == 'CodeBlock' then
      parts[#parts + 1] = b.text
    elseif b.t == 'Para' or b.t == 'Plain' or b.t == 'Header' then
      parts[#parts + 1] = inlines_text(b.content)
    elseif b.t == 'BulletList' or b.t == 'OrderedList' then
      for _, item in ipairs(b.content) do
        parts[#parts + 1] = blocks_text(item)
      end
    elseif b.t == 'LineBlock' then
      for _, line in ipairs(b.content) do
        parts[#parts + 1] = inlines_text(line)
      end
    elseif b.content then
      parts[#parts + 1] = blocks_text(b.content)
    end
  end
  return table.concat(parts, '\n')
end

-- Share of non-whitespace characters that are code (inline Code or CodeBlock).
local function code_share(blocks)
  local code, total = 0, 0
  local function count(text)
    local stripped = text:gsub(NBSP, ''):gsub('%s', '')
    return #stripped
  end
  pandoc.Blocks(blocks):walk {
    CodeBlock = function(cb) local n = count(cb.text); code = code + n; total = total + n end,
    Code = function(c) local n = count(c.text); code = code + n; total = total + n end,
    Str = function(s) total = total + count(s.text) end,
  }
  if total == 0 then return 0 end
  return code / total
end

-- Does a box read like prose (a note) rather than code/commands/output?
-- Prose: contains lists or links, or most non-empty lines are sentences
-- (at least 8 words and ending in . ! or ?). Language-neutral and domain-neutral.
local function looks_like_prose(blocks)
  local rich = false
  pandoc.Blocks(blocks):walk {
    BulletList = function() rich = true end,
    OrderedList = function() rich = true end,
    -- bare URLs (autolinks: text == target) are common in commands and don't count
    Link = function(l) if stringify(l.content) ~= l.target then rich = true end end,
  }
  if rich then return true end
  local lines, sentences = 0, 0
  for line in (blocks_text(blocks) .. '\n'):gmatch('(.-)\n') do
    line = line:gsub(NBSP, ' '):gsub('^%s+', ''):gsub('%s+$', '')
    if line ~= '' then
      lines = lines + 1
      local _, words = line:gsub('%S+', '')
      if words >= 8 and line:match('[%.!%?]$') then sentences = sentences + 1 end
    end
  end
  return lines > 0 and sentences / lines >= 0.5
end

local function contains_media(blocks)
  local found = false
  pandoc.Blocks(blocks):walk {
    Image = function() found = true end,
    RawInline = function(r) if r.text:find(MEDIA_TOKEN, 1, true) then found = true end end,
  }
  return found
end

-- A box holding headings, tables or images is a layout container: unwrap it.
local function is_container(blocks)
  local structural, images = false, 0
  pandoc.Blocks(blocks):walk {
    Header = function() structural = true end,
    Table = function() structural = true end,
    Image = function() images = images + 1 end,
    RawInline = function(r)
      local _, n = r.text:gsub(MEDIA_TOKEN, '')
      images = images + n
    end,
  }
  return structural or images > 0
end

local function trim_text(text)
  text = text:gsub(NBSP, ' ')
  local lines = {}
  for line in (text .. '\n'):gmatch('(.-)\n') do
    lines[#lines + 1] = (line:gsub('%s+$', ''))
  end
  while #lines > 0 and lines[1] == '' do table.remove(lines, 1) end
  while #lines > 0 and lines[#lines] == '' do table.remove(lines) end
  return table.concat(lines, '\n')
end

-- ---------------------------------------------------------------------------
-- 1. OneNote page header (title, date, time)
-- ---------------------------------------------------------------------------

local function remove_onenote_header(blocks)
  -- OneNote starts every page with: title, date, time (format depends on locale).
  for i = 1, math.min(4, #blocks) do
    local b = blocks[i]
    if b.t ~= 'Para' and b.t ~= 'Plain' and b.t ~= 'Header' then break end
    local text = stringify(b):gsub(NBSP, ' ')
    if text:match('^%s*%d%d?:%d%d') then
      for _ = 1, i do blocks:remove(1) end
      return blocks
    end
  end
  return blocks
end

-- ---------------------------------------------------------------------------
-- 2. Images (C10): order of appearance, written to media-dir, placeholder links
-- ---------------------------------------------------------------------------

local image_count = 0

local UNIT_TO_PX = { ['in'] = 96, ['cm'] = 96 / 2.54, ['mm'] = 96 / 25.4, ['pt'] = 96 / 72, ['px'] = 1, [''] = 1 }

local function image_width_px(img)
  local w = img.attributes and img.attributes.width
  if not w then return nil end
  local value, unit = w:match('^([%d%.]+)(%a*)$')
  local factor = UNIT_TO_PX[unit or '']
  if not value or not factor then return nil end
  local px = math.floor(tonumber(value) * factor + 0.5)
  local max = tonumber(opt['max-image-width']) or 1000
  if px > max then px = max end
  if px < 1 then return nil end
  return px
end

local function handle_image(img)
  if img.src:match('^%a+://') then
    return nil -- external image: leave untouched
  end
  image_count = image_count + 1
  local ext = (img.src:match('%.(%w+)$') or 'png'):lower()
  local name = string.format('%s-%02d.%s', opt['image-prefix'], image_count, ext)

  local _, contents = pandoc.mediabag.lookup(img.src)
  if contents and opt['media-dir'] ~= '' then
    local f = io.open(opt['media-dir'] .. '/' .. name, 'wb')
    if f then
      f:write(contents)
      f:close()
    else
      io.stderr:write('looseleaf: cannot write image ' .. name .. '\n')
    end
  elseif not contents then
    io.stderr:write('looseleaf: image not found in document: ' .. img.src .. '\n')
  end

  local size = ''
  if opt['image-width'] == 'Obsidian' then
    local px = image_width_px(img)
    if px then size = '|' .. px end
  end
  return pandoc.RawInline('markdown', string.format('![%s](%s%s)', size, MEDIA_TOKEN, name))
end

-- ---------------------------------------------------------------------------
-- 3. Inline formatting: underline / highlight (C8), empty spans, nbsp
-- ---------------------------------------------------------------------------

local inline_filter = {
  Underline = function(el)
    if opt['underline'] == 'Drop' then return el.content end
    -- 'markdown' raw format is emitted verbatim; 'html' would become `<u>`{=html}
    local out = pandoc.Inlines { pandoc.RawInline('markdown', '<u>') }
    out:extend(el.content)
    out:insert(pandoc.RawInline('markdown', '</u>'))
    return out
  end,
  Span = function(el)
    if el.classes:includes('mark') then
      local out = pandoc.Inlines { pandoc.RawInline('markdown', '==') }
      out:extend(el.content)
      out:insert(pandoc.RawInline('markdown', '=='))
      return out
    end
    if el.identifier == '' and #el.attributes == 0 then
      return el.content -- drop meaningless spans, e.g. []{.underline} leftovers
    end
  end,
  Str = function(s)
    if s.text:find(NBSP, 1, true) then
      local text = s.text:gsub(NBSP, ' ')
      if text:match('^%s*$') then return pandoc.Space() end
      return pandoc.Str(text)
    end
  end,
}

-- ---------------------------------------------------------------------------
-- 4. Paragraph clean-up (C6) and clip sources (C9)
-- ---------------------------------------------------------------------------

local CLIP_PATTERNS = { '^Pasted from%s*<(.-)>%s*$', '^From%s*<(.-)>%s*$', '^Source:%s*<(.-)>%s*$' }

local function trim_inlines(inlines)
  local strip = { Space = true, SoftBreak = true, LineBreak = true }
  while #inlines > 0 and strip[inlines[#inlines].t] do inlines:remove(#inlines) end
  while #inlines > 0 and strip[inlines[1].t] do inlines:remove(1) end
  -- collapse a line break followed only by spaces and another line break
  local out = pandoc.Inlines {}
  for _, il in ipairs(inlines) do
    local prev = out[#out]
    if il.t == 'Space' and prev and (prev.t == 'LineBreak' or prev.t == 'Space') then
      -- skip
    elseif il.t == 'LineBreak' and prev and prev.t == 'LineBreak' then
      -- skip
    else
      out:insert(il)
    end
  end
  return out
end

local function clean_para(para)
  local content = trim_inlines(para.content)
  if #content == 0 then return {} end

  if opt['clip-sources'] == 'Normalize' then
    local text = stringify(content):gsub(NBSP, ' ')
    for _, pattern in ipairs(CLIP_PATTERNS) do
      local url = text:match(pattern)
      if url and url:match('^%a+://') then
        return pandoc.Para { pandoc.Str('Source:'), pandoc.Space(), pandoc.Link(url, url) }
      end
    end
  end

  local has_media = contains_media({ pandoc.Plain(content) })
  if not has_media and is_blank(inlines_text(content)) then return {} end
  para.content = content
  return para
end

-- ---------------------------------------------------------------------------
-- 5. Tables: single-cell boxes (C1) and multi-cell tables (C3)
-- ---------------------------------------------------------------------------

local function table_cells(tbl)
  local cells = {}
  local function add_rows(rows)
    for _, row in ipairs(rows) do
      for _, cell in ipairs(row.cells) do cells[#cells + 1] = cell end
    end
  end
  add_rows(tbl.head.rows)
  for _, body in ipairs(tbl.bodies) do
    add_rows(body.head)
    add_rows(body.body)
  end
  add_rows(tbl.foot.rows)
  return cells
end

local function box_to_blocks(cell)
  local blocks = cell.contents
  local mode = opt['box-tables']
  if mode == 'Table' then return nil end
  if is_blank(blocks_text(blocks)) and not contains_media(blocks) then return {} end
  if mode == 'Unwrap' or is_container(blocks) then return blocks end

  -- Auto: code if mostly monospace (C2), otherwise code unless the box reads like prose.
  -- OneNote has no code block, so a plain one-cell table is the usual way to hold commands.
  local threshold = tonumber(opt['box-code-threshold']) or 0.6
  local as_code = mode == 'CodeBlock'
      or (mode == 'Auto' and (code_share(blocks) >= threshold or not looks_like_prose(blocks)))
  if as_code and not contains_media(blocks) then
    -- 'llbox' marks a box: separate boxes stay separate code blocks (see merge_code_blocks)
    return { pandoc.CodeBlock(trim_text(blocks_text(blocks)), pandoc.Attr('', { 'llbox' })) }
  end

  local inner = pandoc.Blocks { pandoc.Plain { pandoc.RawInline('markdown', '[!note]') } }
  inner:extend(blocks)
  return { pandoc.BlockQuote(inner) }
end

local BR = pandoc.RawInline('markdown', '<br>')

-- Make a cell representable in a pipe table: paragraphs joined with <br>.
local function flatten_cell(cell)
  local blocks = cell.contents
  local inlines = pandoc.Inlines {}
  for _, b in ipairs(blocks) do
    if b.t ~= 'Para' and b.t ~= 'Plain' then return end -- lists etc.: keep (HTML table)
    if #inlines > 0 then inlines:insert(BR) end
    inlines:extend(b.content)
  end
  inlines = inlines:walk { LineBreak = function() return BR end, SoftBreak = function() return BR end }
  cell.contents = { pandoc.Plain(inlines) }
end

local function handle_table(tbl)
  local cells = table_cells(tbl)
  if #cells == 1 then
    return box_to_blocks(cells[1])
  end
  for _, cell in ipairs(cells) do flatten_cell(cell) end
  return tbl
end

-- ---------------------------------------------------------------------------
-- 6. Glyph bullets (C5) and adjacent code blocks
-- ---------------------------------------------------------------------------

local GLYPHS = {
  ['●'] = true, ['•'] = true, ['○'] = true, ['◦'] = true, ['▪'] = true, ['▫'] = true,
  ['■'] = true, ['□'] = true, ['➢'] = true, ['►'] = true, ['▶'] = true, ['–'] = true,
  ['+'] = true, ['-'] = true, ['*'] = true,
}

local function strip_glyph(block)
  if block.t ~= 'Para' and block.t ~= 'Plain' then return nil end
  local c = block.content
  if #c == 0 or c[1].t ~= 'Str' then return nil end
  local text = c[1].text
  local glyph = text:match('^[%z\1-\127\194-\244][\128-\191]*')
  if not glyph or not GLYPHS[glyph] then return nil end
  local rest = text:sub(#glyph + 1)
  local out = c:clone()
  if rest == '' then
    if not (c[2] and c[2].t == 'Space') then return nil end
    out:remove(1)
    out:remove(1)
  else
    if #glyph == 1 then return nil end -- ASCII glyph must stand alone ("+5" is not a bullet)
    out[1] = pandoc.Str(rest)
  end
  if #out == 0 then return nil end
  return out
end

local function convert_glyph_lists(blocks)
  local out = pandoc.Blocks {}
  local items = nil
  local function flush()
    if items then
      out:insert(pandoc.BulletList(items))
      items = nil
    end
  end
  for _, b in ipairs(blocks) do
    local inlines = strip_glyph(b)
    if inlines then
      items = items or {}
      items[#items + 1] = { pandoc.Plain(inlines) }
    else
      flush()
      out:insert(b)
    end
  end
  flush()
  return out
end

-- pandoc writes attribute-less code blocks indented; Obsidian users expect fences.
local function fence_code_block(cb)
  local text = trim_text(cb.text)
  if text == '' then return {} end
  local longest = 2
  for run in text:gmatch('`+') do
    if #run > longest then longest = #run end
  end
  local fence = string.rep('`', longest + 1)
  local lang = ''
  for _, class in ipairs(cb.classes) do
    if class ~= 'llbox' then lang = class; break end
  end
  return pandoc.RawBlock('markdown', fence .. lang .. '\n' .. text .. '\n' .. fence)
end

local function merge_code_blocks(blocks)
  local out = pandoc.Blocks {}
  for _, b in ipairs(blocks) do
    local prev = out[#out]
    -- merge consecutive code paragraphs (rule C2), but never two separate boxes (C1)
    if b.t == 'CodeBlock' and prev and prev.t == 'CodeBlock'
        and not b.classes:includes('llbox') and not prev.classes:includes('llbox') then
      prev.text = prev.text .. '\n' .. b.text
    else
      out:insert(b)
    end
  end
  return out
end

-- ---------------------------------------------------------------------------
-- Main
-- ---------------------------------------------------------------------------

function Pandoc(doc)
  read_options(doc.meta)

  doc.blocks = remove_onenote_header(doc.blocks)

  -- Images first, so numbering follows document order and boxes know about them.
  doc = doc:walk { Image = handle_image }
  doc = doc:walk(inline_filter)
  doc = doc:walk {
    Para = clean_para,
    Plain = clean_para,
    Header = function(h) if is_blank(inlines_text(h.content)) then return {} end end,
  }

  if opt['indented'] == 'Unindent' then
    for _ = 1, 10 do
      local found = false
      doc = doc:walk { BlockQuote = function(bq) found = true; return bq.content end }
      if not found then break end
    end
  end

  -- Tables (callouts created here are block quotes and must survive the unindent pass above).
  doc = doc:walk { Table = handle_table }

  if opt['glyph-bullets'] == 'true' then
    doc = doc:walk { Blocks = convert_glyph_lists }
  end

  doc = doc:walk { Blocks = merge_code_blocks }
  doc = doc:walk { CodeBlock = fence_code_block }
  return doc
end
