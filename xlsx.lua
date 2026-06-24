-- Copyright (C) 2026  Freyermuth Julien
--
-- This program is free software: you can redistribute it and/or modify
-- it under the terms of the GNU General Public License as published by
-- the Free Software Foundation, either version 3 of the License, or
-- (at your option) any later version.
--
-- This program is distributed in the hope that it will be useful,
-- but WITHOUT ANY WARRANTY; without even the implied warranty of
-- MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
-- GNU General Public License for more details.
--
-- You should have received a copy of the GNU General Public License
-- along with this program.  If not, see <https://www.gnu.org/licenses/>.

--- xlsx.lua — lecture/écriture de fichiers .xlsx en Lua pur.
--- Cible : LuaPilot (Lua 5.5). Compatible Lua 5.3+.
--- v1 : ÉCRITURE seulement. Aucune dépendance externe (pas de unzip, pas de C).
---
--- Un .xlsx est une archive ZIP de fichiers XML. Ici on construit les XML
--- à la main et on les empaquette en ZIP STORED (non compressé) — Excel et
--- LibreOffice lisent ça parfaitement. Le seul morceau "binaire" est le CRC32
--- exigé par le format ZIP, implémenté en Lua pur ci-dessous.

local xlsx = {}

-- ----------------------------------------------------------------------------
-- CRC32 (table-based, polynôme ZIP standard 0xEDB88320)
-- ----------------------------------------------------------------------------
local crc_table
local function init_crc()
  crc_table = {}
  for i = 0, 255 do
    local c = i
    for _ = 1, 8 do
      if (c & 1) ~= 0 then
        c = 0xEDB88320 ~ (c >> 1)
      else
        c = c >> 1
      end
    end
    crc_table[i] = c
  end
end

local function crc32(s)
  if not crc_table then init_crc() end
  local crc = 0xFFFFFFFF
  for i = 1, #s do
    crc = crc_table[(crc ~ s:byte(i)) & 0xFF] ~ (crc >> 8)
  end
  return (crc ~ 0xFFFFFFFF) & 0xFFFFFFFF
end

-- ----------------------------------------------------------------------------
-- Helpers XML / colonnes / nombres
-- ----------------------------------------------------------------------------
local XML_ESC = { ["&"] = "&amp;", ["<"] = "&lt;", [">"] = "&gt;", ['"'] = "&quot;" }
local function esc(s)
  return (s:gsub('[&<>"]', XML_ESC))
end

-- colonne 0-indexée -> lettres ("A", "B", ..., "AA", ...)
local function col_ref(c0)
  local n = c0 + 1
  local out = ""
  while n > 0 do
    local rem = (n - 1) % 26
    out = string.char(65 + rem) .. out
    n = (n - 1) // 26
  end
  return out
end

-- nombre -> string round-trip (le plus court qui se relit exactement)
local function num2str(v)
  if math.type(v) == "integer" then
    return tostring(v)
  end
  local s = string.format("%.15g", v)
  if tonumber(s) ~= v then
    s = string.format("%.17g", v)
  end
  return s
end

-- ----------------------------------------------------------------------------
-- Dates (système 1900 d'Excel ; epoch effectif 1899-12-30, comme openpyxl)
-- ----------------------------------------------------------------------------
-- Conversion calendaire entière (algorithme de Howard Hinnant), indépendante
-- de os.time/os.date donc sans souci de fuseau ni de plage.
local function days_from_civil(y, m, d)
  if m <= 2 then y = y - 1 end
  local era = (y >= 0 and y or y - 399) // 400
  local yoe = y - era * 400
  local mp = (m > 2) and (m - 3) or (m + 9)
  local doy = (153 * mp + 2) // 5 + d - 1
  local doe = yoe * 365 + yoe // 4 - yoe // 100 + doy
  return era * 146097 + doe - 719468
end

local function civil_from_days(z)
  z = z + 719468
  local era = (z >= 0 and z or z - 146096) // 146097
  local doe = z - era * 146097
  local yoe = (doe - doe // 1460 + doe // 36524 - doe // 146096) // 365
  local y = yoe + era * 400
  local doy = doe - (365 * yoe + yoe // 4 - yoe // 100)
  local mp = (5 * doy + 2) // 153
  local d = doy - (153 * mp + 2) // 5 + 1
  local m = (mp < 10) and (mp + 3) or (mp - 9)
  if m <= 2 then y = y + 1 end
  return y, m, d
end

local EXCEL_EPOCH = days_from_civil(1899, 12, 30) -- = -25569

local function ymd_to_serial(y, m, d, h, mi, s)
  local days = days_from_civil(y, m, d) - EXCEL_EPOCH
  return days + (h * 3600 + mi * 60 + s) / 86400
end

-- numéro de série Excel -> chaîne ISO 8601 (cohérent avec luapilot.toml)
local function serial_to_iso(serial, withtime)
  local whole = math.floor(serial)
  local frac = serial - whole
  if not withtime then
    local y, m, d = civil_from_days(whole + EXCEL_EPOCH)
    return string.format("%04d-%02d-%02d", y, m, d)
  end
  local secs = math.floor(frac * 86400 + 0.5)
  if secs >= 86400 then secs = secs - 86400; whole = whole + 1 end
  local y, m, d = civil_from_days(whole + EXCEL_EPOCH)
  return string.format("%04d-%02d-%02dT%02d:%02d:%02d",
    y, m, d, secs // 3600, (secs % 3600) // 60, secs % 60)
end

-- valeur "date" tagguée, reconnue par Sheet:write
local DATE_MT = {}

--- Crée une valeur date (affichée yyyy-mm-dd dans Excel).
function xlsx.date(y, m, d)
  if math.type(y) ~= "integer" or math.type(m) ~= "integer" or math.type(d) ~= "integer" then
    error("xlsx.date attend des entiers (année, mois, jour)", 2)
  end
  return setmetatable({ serial = ymd_to_serial(y, m, d, 0, 0, 0), kind = "date" }, DATE_MT)
end

--- Crée une valeur date-heure (affichée yyyy-mm-dd hh:mm:ss dans Excel).
function xlsx.datetime(y, m, d, h, mi, s)
  return setmetatable(
    { serial = ymd_to_serial(y, m, d, h or 0, mi or 0, s or 0), kind = "datetime" },
    DATE_MT)
end

-- exposés pour conversions manuelles
xlsx.serial_to_iso = serial_to_iso
function xlsx.date_to_serial(y, m, d, h, mi, s)
  return ymd_to_serial(y, m, d, h or 0, mi or 0, s or 0)
end

-- ----------------------------------------------------------------------------
-- Worksheet
-- ----------------------------------------------------------------------------
local Sheet = {}
Sheet.__index = Sheet

local function new_sheet(name)
  return setmetatable({
    name   = name,
    cells  = {},   -- cells[row][col] = valeur Lua brute
    maxrow = -1,
    maxcol = -1,
  }, Sheet)
end

local function check_index(v, what, level)
  if math.type(v) ~= "integer" or v < 0 then
    error("xlsx: " .. what .. " doit être un entier >= 0", level)
  end
end

--- Écrit une cellule. row/col sont 0-indexés. value : string|number|boolean|nil.
function Sheet:write(row, col, value)
  check_index(row, "row", 3)
  check_index(col, "col", 3)
  local t = type(value)
  if t == "table" and getmetatable(value) == DATE_MT then
    -- valeur date : acceptée telle quelle
  elseif value ~= nil and t ~= "string" and t ~= "number" and t ~= "boolean" then
    error("xlsx: type de cellule non supporté : " .. t, 2)
  end
  if t == "number" and (value ~= value or value == math.huge or value == -math.huge) then
    error("xlsx: impossible d'écrire NaN/Inf dans une cellule", 2)
  end
  local r = self.cells[row]
  if not r then r = {}; self.cells[row] = r end
  r[col] = value
  if value ~= nil then
    if row > self.maxrow then self.maxrow = row end
    if col > self.maxcol then self.maxcol = col end
  end
  return self
end

--- Ajoute une ligne complète à la suite (les nil en fin de table sont ignorés).
function Sheet:append_row(values)
  if type(values) ~= "table" then
    error("xlsx: append_row attend une table", 2)
  end
  local row = self.maxrow + 1
  local n = #values
  for i = 1, n do
    local v = values[i]
    if v ~= nil then self:write(row, i - 1, v) end
  end
  -- garantir que la ligne existe même si entièrement vide
  if not self.cells[row] then self.cells[row] = {} end
  if row > self.maxrow then self.maxrow = row end
  return self
end

--- Écrit une matrice (table de tables) à partir de la position courante.
function Sheet:write_rows(matrix)
  if type(matrix) ~= "table" then
    error("xlsx: write_rows attend une table de lignes", 2)
  end
  for i = 1, #matrix do
    self:append_row(matrix[i])
  end
  return self
end

-- sérialise la feuille en XML ; alimente sst (liste) et sst_index (map)
function Sheet:_xml(sst, sst_index)
  local b = {}
  b[#b + 1] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  b[#b + 1] = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
  if self.maxrow >= 0 and self.maxcol >= 0 then
    b[#b + 1] = '<dimension ref="A1:' .. col_ref(self.maxcol) .. (self.maxrow + 1) .. '"/>'
  end
  b[#b + 1] = '<sheetData>'
  for r = 0, self.maxrow do
    local rowcells = self.cells[r]
    if rowcells then
      b[#b + 1] = '<row r="' .. (r + 1) .. '">'
      for c = 0, self.maxcol do
        local v = rowcells[c]
        if v ~= nil then
          local ref = col_ref(c) .. (r + 1)
          local tv = type(v)
          if tv == "table" then -- valeur date taggée
            local style = (v.kind == "datetime") and 2 or 1
            b[#b + 1] = '<c r="' .. ref .. '" s="' .. style .. '"><v>' ..
              num2str(v.serial) .. '</v></c>'
          elseif tv == "string" then
            local idx = sst_index[v]
            if not idx then
              idx = #sst
              sst[#sst + 1] = v
              sst_index[v] = idx
            end
            b[#b + 1] = '<c r="' .. ref .. '" t="s"><v>' .. idx .. '</v></c>'
          elseif tv == "boolean" then
            b[#b + 1] = '<c r="' .. ref .. '" t="b"><v>' .. (v and "1" or "0") .. '</v></c>'
          else
            b[#b + 1] = '<c r="' .. ref .. '"><v>' .. num2str(v) .. '</v></c>'
          end
        end
      end
      b[#b + 1] = '</row>'
    end
  end
  b[#b + 1] = '</sheetData></worksheet>'
  return table.concat(b)
end

-- ----------------------------------------------------------------------------
-- Workbook
-- ----------------------------------------------------------------------------
local Workbook = {}
Workbook.__index = Workbook

function xlsx.new()
  return setmetatable({ sheets = {} }, Workbook)
end

function Workbook:add_sheet(name)
  if name == nil then
    name = "Sheet" .. (#self.sheets + 1)
  end
  if type(name) ~= "string" then
    error("xlsx: le nom de feuille doit être une string", 2)
  end
  if #name == 0 or #name > 31 then
    error("xlsx: nom de feuille invalide (1 à 31 caractères)", 2)
  end
  if name:find('[:\\/?*%[%]]') then
    error("xlsx: le nom de feuille contient un caractère interdit ( : \\ / ? * [ ] )", 2)
  end
  local sh = new_sheet(name)
  self.sheets[#self.sheets + 1] = sh
  return sh
end

-- construit la liste des parties ZIP {name=, data=}
function Workbook:_parts()
  local nsheets = #self.sheets
  local sst, sst_index = {}, {}
  local sheet_xmls = {}
  for i = 1, nsheets do
    sheet_xmls[i] = self.sheets[i]:_xml(sst, sst_index)
  end

  local parts = {}
  local function add(name, data) parts[#parts + 1] = { name = name, data = data } end

  -- [Content_Types].xml
  do
    local t = {}
    t[#t + 1] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    t[#t + 1] = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    t[#t + 1] = '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    t[#t + 1] = '<Default Extension="xml" ContentType="application/xml"/>'
    t[#t + 1] = '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    for i = 1, nsheets do
      t[#t + 1] = '<Override PartName="/xl/worksheets/sheet' .. i ..
        '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    end
    t[#t + 1] = '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    t[#t + 1] = '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    t[#t + 1] = '</Types>'
    add("[Content_Types].xml", table.concat(t))
  end

  -- _rels/.rels
  add("_rels/.rels", table.concat({
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    '</Relationships>',
  }))

  -- xl/workbook.xml
  do
    local t = {}
    t[#t + 1] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    t[#t + 1] = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    t[#t + 1] = '<sheets>'
    for i = 1, nsheets do
      t[#t + 1] = '<sheet name="' .. esc(self.sheets[i].name) ..
        '" sheetId="' .. i .. '" r:id="rId' .. i .. '"/>'
    end
    t[#t + 1] = '</sheets></workbook>'
    add("xl/workbook.xml", table.concat(t))
  end

  -- xl/_rels/workbook.xml.rels  (sheets rId1..N, sharedStrings rId(N+1))
  do
    local t = {}
    t[#t + 1] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    t[#t + 1] = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    for i = 1, nsheets do
      t[#t + 1] = '<Relationship Id="rId' .. i ..
        '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' ..
        i .. '.xml"/>'
    end
    t[#t + 1] = '<Relationship Id="rId' .. (nsheets + 1) ..
      '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    t[#t + 1] = '<Relationship Id="rId' .. (nsheets + 2) ..
      '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    t[#t + 1] = '</Relationships>'
    add("xl/_rels/workbook.xml.rels", table.concat(t))
  end

  -- xl/worksheets/sheetN.xml
  for i = 1, nsheets do
    add("xl/worksheets/sheet" .. i .. ".xml", sheet_xmls[i])
  end

  -- xl/sharedStrings.xml
  do
    local t = {}
    t[#t + 1] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    t[#t + 1] = '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' ..
      #sst .. '" uniqueCount="' .. #sst .. '">'
    for i = 1, #sst do
      t[#t + 1] = '<si><t xml:space="preserve">' .. esc(sst[i]) .. '</t></si>'
    end
    t[#t + 1] = '</sst>'
    add("xl/sharedStrings.xml", table.concat(t))
  end

  -- xl/styles.xml  (3 styles : 0 = général, 1 = date, 2 = date-heure)
  add("xl/styles.xml", table.concat({
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<numFmts count="2">',
    '<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>',
    '<numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm:ss"/>',
    '</numFmts>',
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>',
    '<fills count="2"><fill><patternFill patternType="none"/></fill>',
    '<fill><patternFill patternType="gray125"/></fill></fills>',
    '<borders count="1"><border/></borders>',
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
    '<cellXfs count="3">',
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
    '</cellXfs>',
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
    '</styleSheet>',
  }))

  return parts
end

-- ----------------------------------------------------------------------------
-- Empaquetage ZIP (STORED) + écriture fichier
-- ----------------------------------------------------------------------------
local SIG_LOCAL   = 0x04034b50
local SIG_CENTRAL = 0x02014b50
local SIG_EOCD    = 0x06054b50
local DOS_TIME    = 0          -- 00:00:00
local DOS_DATE    = 0x21       -- 1980-01-01

local function build_zip(parts)
  local out, central = {}, {}
  local offset = 0
  local function emit(s) out[#out + 1] = s; offset = offset + #s end

  for i = 1, #parts do
    local p = parts[i]
    local data, name = p.data, p.name
    local crc = crc32(data)
    local local_off = offset

    emit(string.pack("<I4 I2 I2 I2 I2 I2 I4 I4 I4 I2 I2",
      SIG_LOCAL, 20, 0, 0, DOS_TIME, DOS_DATE, crc, #data, #data, #name, 0))
    emit(name)
    emit(data)

    central[#central + 1] = string.pack("<I4 I2 I2 I2 I2 I2 I2 I4 I4 I4 I2 I2 I2 I2 I2 I4 I4",
      SIG_CENTRAL, 20, 20, 0, 0, DOS_TIME, DOS_DATE, crc, #data, #data,
      #name, 0, 0, 0, 0, 0, local_off)
    central[#central + 1] = name
  end

  local cd_start = offset
  local cd = table.concat(central)
  emit(cd)
  emit(string.pack("<I4 I2 I2 I2 I2 I4 I4 I2",
    SIG_EOCD, 0, 0, #parts, #parts, #cd, cd_start, 0))

  return table.concat(out)
end

--- Écrit le classeur dans `path`. Renvoie true, ou (nil, err) en cas d'échec I/O.
function Workbook:save(path)
  if type(path) ~= "string" then
    error("xlsx: le chemin de sauvegarde doit être une string", 2)
  end
  if #self.sheets == 0 then self:add_sheet("Sheet1") end

  local bytes = build_zip(self:_parts())

  local f, err = io.open(path, "wb")
  if not f then
    return nil, "xlsx: impossible d'ouvrir '" .. path .. "' : " .. tostring(err)
  end
  f:write(bytes)
  f:close()
  return true
end

-- ----------------------------------------------------------------------------
-- Raccourci pratique
-- ----------------------------------------------------------------------------
--- Écrit une matrice directement dans un fichier. opts.sheet = nom de feuille.
function xlsx.write_rows(path, matrix, opts)
  opts = opts or {}
  local wb = xlsx.new()
  local sh = wb:add_sheet(opts.sheet or "Sheet1")
  sh:write_rows(matrix)
  return wb:save(path)
end

-- ============================================================================
-- LECTURE
-- ============================================================================
-- v1 : valeurs (string / number / boolean), sharedStrings, inlineStr, plusieurs
-- feuilles, cellules éparses. Hors v1 lecture : styles, dates typées, formules
-- (la valeur cachée d'une formule EST lue), formats de nombre.

-- ----------------------------------------------------------------------------
-- INFLATE (RFC 1951) — décompression DEFLATE brute, en Lua pur
-- ----------------------------------------------------------------------------
local LEN_BASE = {3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,
  115,131,163,195,227,258}
local LEN_EXTRA = {0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0}
local DIST_BASE = {1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,
  1025,1537,2049,3073,4097,6145,8193,12289,16385,24577}
local DIST_EXTRA = {0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,
  12,13,13}
local CLC_ORDER = {16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15}

-- construit un arbre Huffman canonique à partir des longueurs (clés 0..num-1)
local function huff_build(lengths, num)
  local counts = {}
  for i = 0, 15 do counts[i] = 0 end
  for sym = 0, num - 1 do
    local l = lengths[sym] or 0
    counts[l] = counts[l] + 1
  end
  counts[0] = 0
  local offs = { [1] = 0 }
  for i = 1, 15 do offs[i + 1] = offs[i] + counts[i] end
  local symbols = {}
  for sym = 0, num - 1 do
    local l = lengths[sym] or 0
    if l ~= 0 then
      symbols[offs[l]] = sym
      offs[l] = offs[l] + 1
    end
  end
  return counts, symbols
end

local function inflate(input)
  local byte = string.byte
  local n = #input
  local pos = 1
  local bitbuf, bitcnt = 0, 0

  local function getbits(num)
    while bitcnt < num do
      if pos > n then error("xlsx: inflate: données tronquées", 0) end
      bitbuf = bitbuf | (byte(input, pos) << bitcnt)
      pos = pos + 1
      bitcnt = bitcnt + 8
    end
    local v = bitbuf & ((1 << num) - 1)
    bitbuf = bitbuf >> num
    bitcnt = bitcnt - num
    return v
  end

  -- décode un symbole (algorithme "puff")
  local function decode(counts, symbols)
    local code, first, index = 0, 0, 0
    for len = 1, 15 do
      code = code | getbits(1)
      local count = counts[len]
      if code - count < first then
        return symbols[index + (code - first)]
      end
      index = index + count
      first = (first + count) << 1
      code = code << 1
    end
    error("xlsx: inflate: code Huffman invalide", 0)
  end

  local out, olen = {}, 0

  local function inflate_block(lc, ls, dc, ds)
    while true do
      local sym = decode(lc, ls)
      if sym == 256 then return end
      if sym < 256 then
        olen = olen + 1
        out[olen] = sym
      else
        local li = sym - 256
        local length = LEN_BASE[li] + getbits(LEN_EXTRA[li])
        local dsym = decode(dc, ds)
        local distance = DIST_BASE[dsym + 1] + getbits(DIST_EXTRA[dsym + 1])
        local start = olen - distance
        for k = 1, length do
          olen = olen + 1
          out[olen] = out[start + k]
        end
      end
    end
  end

  -- tables Huffman fixes (construites une fois par appel)
  local fixed_lit = {}
  for s = 0, 143 do fixed_lit[s] = 8 end
  for s = 144, 255 do fixed_lit[s] = 9 end
  for s = 256, 279 do fixed_lit[s] = 7 end
  for s = 280, 287 do fixed_lit[s] = 8 end
  local fixed_dist = {}
  for s = 0, 29 do fixed_dist[s] = 5 end

  repeat
    local bfinal = getbits(1)
    local btype = getbits(2)
    if btype == 0 then
      -- bloc stocké : aligner sur l'octet
      local drop = bitcnt % 8
      bitbuf = bitbuf >> drop
      bitcnt = bitcnt - drop
      local len = getbits(16)
      local nlen = getbits(16)
      if (len ~ 0xFFFF) ~= nlen then
        error("xlsx: inflate: bloc stocké corrompu", 0)
      end
      for _ = 1, len do
        olen = olen + 1
        out[olen] = getbits(8)
      end
    elseif btype == 1 then
      local lc, ls = huff_build(fixed_lit, 288)
      local dc, ds = huff_build(fixed_dist, 30)
      inflate_block(lc, ls, dc, ds)
    elseif btype == 2 then
      local hlit = getbits(5) + 257
      local hdist = getbits(5) + 1
      local hclen = getbits(4) + 4
      local cll = {}
      for i = 0, 18 do cll[i] = 0 end
      for i = 1, hclen do cll[CLC_ORDER[i]] = getbits(3) end
      local clc, cls = huff_build(cll, 19)
      local all = {}
      local i = 0
      local total = hlit + hdist
      while i < total do
        local sym = decode(clc, cls)
        if sym < 16 then
          all[i] = sym; i = i + 1
        elseif sym == 16 then
          local prev = all[i - 1]
          for _ = 1, getbits(2) + 3 do all[i] = prev; i = i + 1 end
        elseif sym == 17 then
          for _ = 1, getbits(3) + 3 do all[i] = 0; i = i + 1 end
        else
          for _ = 1, getbits(7) + 11 do all[i] = 0; i = i + 1 end
        end
      end
      local litl = {}
      for k = 0, hlit - 1 do litl[k] = all[k] end
      local distl = {}
      for k = 0, hdist - 1 do distl[k] = all[hlit + k] end
      local lc, ls = huff_build(litl, hlit)
      local dc, ds = huff_build(distl, hdist)
      inflate_block(lc, ls, dc, ds)
    else
      error("xlsx: inflate: type de bloc réservé", 0)
    end
  until bfinal == 1

  -- octets -> string (par paquets pour rester sous la limite de table.unpack)
  local parts = {}
  local CHUNK = 2000
  for i = 1, olen, CHUNK do
    local j = i + CHUNK - 1
    if j > olen then j = olen end
    parts[#parts + 1] = string.char(table.unpack(out, i, j))
  end
  return table.concat(parts)
end

-- ----------------------------------------------------------------------------
-- Lecture ZIP (répertoire central -> extraction d'une entrée)
-- ----------------------------------------------------------------------------
local function find_eocd(s)
  local n = #s
  if n < 22 then error("xlsx: fichier trop court pour être un zip", 0) end
  local lo = n - 65557
  if lo < 1 then lo = 1 end
  for i = n - 21, lo, -1 do
    if s:byte(i) == 0x50 and s:byte(i + 1) == 0x4b
      and s:byte(i + 2) == 0x05 and s:byte(i + 3) == 0x06 then
      return i
    end
  end
  error("xlsx: EOCD zip introuvable (pas un xlsx ?)", 0)
end

local function read_zip(s)
  local eocd = find_eocd(s)
  local _, _, _, _, nTotal, _, cdOff = string.unpack("<I4 I2 I2 I2 I2 I4 I4", s, eocd)
  local byname = {}
  local pos = cdOff + 1
  for _ = 1, nTotal do
    local sig, _vm, _vn, _fl, method, _mt, _md, _crc, csize, _us,
          namelen, extralen, commentlen, _ds, _ia, _ea, lhoff, np =
      string.unpack("<I4 I2 I2 I2 I2 I2 I2 I4 I4 I4 I2 I2 I2 I2 I2 I4 I4", s, pos)
    if sig ~= 0x02014b50 then error("xlsx: répertoire central zip corrompu", 0) end
    local name = s:sub(np, np + namelen - 1)
    byname[name] = { method = method, csize = csize, lhoff = lhoff }
    pos = np + namelen + extralen + commentlen
  end
  return byname
end

local function zip_extract(s, e)
  local sig, _v, _f, _m, _t, _d, _c, _cs, _us, lnamelen, lextralen, dp =
    string.unpack("<I4 I2 I2 I2 I2 I2 I4 I4 I4 I2 I2", s, e.lhoff + 1)
  if sig ~= 0x04034b50 then error("xlsx: en-tête local zip corrompu", 0) end
  local start = dp + lnamelen + lextralen
  local comp = s:sub(start, start + e.csize - 1)
  if e.method == 0 then
    return comp
  elseif e.method == 8 then
    return inflate(comp)
  end
  error("xlsx: méthode de compression zip non supportée : " .. e.method, 0)
end

-- ----------------------------------------------------------------------------
-- Helpers XML (parsing léger adapté au xlsx, pas un parseur XML général)
-- ----------------------------------------------------------------------------
local ENT = { lt = "<", gt = ">", amp = "&", quot = '"', apos = "'" }
local function unescape(s)
  if not s or not s:find("&", 1, true) then return s end
  return (s:gsub("&(#?[%w]+);", function(e)
    local known = ENT[e]
    if known then return known end
    local hex = e:match("^#[xX](%x+)$")
    if hex then return utf8.char(tonumber(hex, 16)) end
    local dec = e:match("^#(%d+)$")
    if dec then return utf8.char(tonumber(dec)) end
    return "&" .. e .. ";"
  end))
end

local function ref_to_rowcol(ref)
  local letters, digits = ref:match("^(%a+)(%d+)$")
  if not letters then return nil end
  local col = 0
  for i = 1, #letters do
    col = col * 26 + (string.byte(letters, i) - 64)
  end
  return tonumber(digits) - 1, col - 1
end

local function parse_shared_strings(xml)
  local res = {}
  local idx = 0
  xml = xml:gsub("<si%s*/>", "<si></si>")
  for si in xml:gmatch("<si[%s>].-</si>") do
    local buf = {}
    for t in si:gmatch("<t[^>]*>(.-)</t>") do
      buf[#buf + 1] = unescape(t)
    end
    res[idx] = table.concat(buf)
    idx = idx + 1
  end
  return res  -- 0-indexé
end

local function decode_cell(inner, ty, shared)
  if ty == "s" then
    local v = inner:match("<v>(.-)</v>")
    if not v then return nil end
    return shared[tonumber(v)]
  elseif ty == "b" then
    return inner:match("<v>(.-)</v>") == "1"
  elseif ty == "inlineStr" then
    local buf = {}
    for t in inner:gmatch("<t[^>]*>(.-)</t>") do buf[#buf + 1] = unescape(t) end
    return table.concat(buf)
  elseif ty == "str" then
    local v = inner:match("<v>(.-)</v>")
    return v and unescape(v) or nil
  else
    local v = inner:match("<v>(.-)</v>")
    return v and tonumber(v) or nil
  end
end

-- analyse styles.xml -> map index_de_style(0-based) -> "date" | "datetime"
local BUILTIN_DATE = {
  [14] = "date", [15] = "date", [16] = "date", [17] = "date",
  [18] = "datetime", [19] = "datetime", [20] = "datetime", [21] = "datetime",
  [22] = "datetime", [45] = "datetime", [46] = "datetime", [47] = "datetime",
}
local function parse_styles(xml)
  local datestyle = {}
  if not xml then return datestyle end
  -- formats personnalisés (numFmtId >= 164)
  local fmtcode = {}
  for tag in xml:gmatch("<numFmt%s.-/>") do
    local id = tag:match('numFmtId="(%d+)"')
    local code = tag:match('formatCode="([^"]*)"')
    if id and code then fmtcode[tonumber(id)] = unescape(code) end
  end
  local function classify(id)
    local b = BUILTIN_DATE[id]
    if b then return b end
    local code = fmtcode[id]
    if not code then return nil end
    -- retire littéraux entre guillemets, sections [..] et échappements
    local c = code:gsub('"[^"]*"', ""):gsub("%[[^%]]*%]", ""):gsub("\\.", ""):lower()
    local has_date = c:find("[yd]") or c:find("mmm")
    local has_time = c:find("[hs]")
    if has_time then return "datetime" end
    if has_date then return "date" end
    return nil
  end
  -- cellXfs : un xf par index, dans l'ordre
  local cellxfs = xml:match("<cellXfs[^>]*>(.-)</cellXfs>")
  if cellxfs then
    local idx = 0
    for tag in cellxfs:gmatch("<xf%s[^>]*>") do
      local cls = classify(tonumber(tag:match('numFmtId="(%d+)"') or "0"))
      if cls then datestyle[idx] = cls end
      idx = idx + 1
    end
  end
  return datestyle
end

local function parse_sheet(xml, shared, datestyle)
  local cells = {}
  local maxrow, maxcol = -1, -1
  for crow in xml:gmatch("<row[%s>].-</row>") do
    local p = 1
    while true do
      local cs = crow:find("<c[%s/>]", p)
      if not cs then break end
      local tagend = crow:find(">", cs)
      local tag = crow:sub(cs, tagend)
      local selfclose = tag:sub(-2) == "/>"
      local ref = tag:match('r="([^"]*)"')
      local ty = tag:match('t="([^"]*)"')
      local st = tag:match('s="(%d+)"')
      local value, nextp
      if selfclose then
        value, nextp = nil, tagend + 1
      else
        local close = crow:find("</c>", tagend, true)
        local inner = crow:sub(tagend + 1, close - 1)
        value = decode_cell(inner, ty, shared)
        nextp = close + 4
      end
      -- conversion date : cellule numérique portant un style de date
      if type(value) == "number" and st and datestyle then
        local cls = datestyle[tonumber(st)]
        if cls then value = serial_to_iso(value, cls == "datetime") end
      end
      if ref and value ~= nil then
        local r0, c0 = ref_to_rowcol(ref)
        if r0 then
          local row = cells[r0]
          if not row then row = {}; cells[r0] = row end
          row[c0] = value
          if r0 > maxrow then maxrow = r0 end
          if c0 > maxcol then maxcol = c0 end
        end
      end
      p = nextp
    end
  end
  return cells, maxrow, maxcol
end

local function parse_workbook(wbxml, relsxml)
  local rid2target = {}
  for tag in relsxml:gmatch("<Relationship%s.-/>") do
    local id = tag:match('Id="([^"]*)"')
    local target = tag:match('Target="([^"]*)"')
    if id and target then rid2target[id] = target end
  end
  local sheets = {}
  for tag in wbxml:gmatch("<sheet%s.-/>") do
    local name = tag:match('name="([^"]*)"')
    local rid = tag:match('r:id="([^"]*)"')
    sheets[#sheets + 1] = { name = unescape(name or ""), rid = rid }
  end
  for _, sh in ipairs(sheets) do
    local t = rid2target[sh.rid]
    if t then
      if t:sub(1, 1) == "/" then
        sh.path = t:sub(2)
      else
        sh.path = "xl/" .. t
      end
    end
  end
  return sheets
end

-- ----------------------------------------------------------------------------
-- Objets de lecture
-- ----------------------------------------------------------------------------
local ReadSheet = {}
ReadSheet.__index = ReadSheet

--- Lit une cellule (row, col 0-indexés, comme à l'écriture). nil si vide.
function ReadSheet:read(r, c)
  if math.type(r) ~= "integer" or math.type(c) ~= "integer" then
    error("xlsx: read attend des entiers (row, col)", 2)
  end
  local row = self.cells[r]
  if not row then return nil end
  return row[c]
end

--- Renvoie maxrow, maxcol (0-indexés ; -1, -1 si la feuille est vide).
function ReadSheet:dims()
  return self.maxrow, self.maxcol
end

--- Itérateur de lignes. Chaque `row` est un tableau 1-indexé (row[1] = col A),
--- avec row.n = nombre de colonnes. Les cellules vides sont nil.
function ReadSheet:rows()
  local r = -1
  return function()
    r = r + 1
    if r > self.maxrow then return nil end
    local src = self.cells[r]
    local row = { n = self.maxcol + 1 }
    if src then
      for c = 0, self.maxcol do row[c + 1] = src[c] end
    end
    return row
  end
end

local ReadWB = {}
ReadWB.__index = ReadWB

--- Liste ordonnée des noms de feuilles.
function ReadWB:sheet_names()
  local t = {}
  for i, sh in ipairs(self._sheets) do t[i] = sh.name end
  return t
end

--- Récupère une feuille par nom (string) ou index 1-based (number).
function ReadWB:sheet(which)
  local def
  if type(which) == "number" then
    def = self._sheets[which]
  elseif type(which) == "string" then
    for _, sh in ipairs(self._sheets) do
      if sh.name == which then def = sh; break end
    end
  else
    error("xlsx: sheet attend un nom (string) ou un index (number)", 2)
  end
  if not def then return nil end
  if self._cache[def] then return self._cache[def] end
  if not def.path then return nil, "xlsx: cible de feuille introuvable" end
  local e = self._byname[def.path]
  if not e then return nil, "xlsx: partie manquante : " .. def.path end
  local cells, maxrow, maxcol = parse_sheet(zip_extract(self._data, e), self._shared, self._datestyle)
  local rs = setmetatable(
    { name = def.name, cells = cells, maxrow = maxrow, maxcol = maxcol },
    ReadSheet)
  self._cache[def] = rs
  return rs
end

--- Ouvre un .xlsx en lecture. Renvoie un workbook, ou (nil, err).
function xlsx.open(path)
  if type(path) ~= "string" then
    error("xlsx: open attend un chemin (string)", 2)
  end
  local f, err = io.open(path, "rb")
  if not f then
    return nil, "xlsx: impossible d'ouvrir '" .. path .. "' : " .. tostring(err)
  end
  local data = f:read("a")
  f:close()

  local ok, byname = pcall(read_zip, data)
  if not ok then return nil, byname end

  local function part(name)
    local e = byname[name]
    return e and zip_extract(data, e) or nil
  end

  local ok2, wbxml = pcall(part, "xl/workbook.xml")
  if not ok2 then return nil, wbxml end
  if not wbxml then return nil, "xlsx: xl/workbook.xml manquant (pas un xlsx ?)" end

  local relsxml = part("xl/_rels/workbook.xml.rels") or ""
  local ssxml = part("xl/sharedStrings.xml")
  local shared = ssxml and parse_shared_strings(ssxml) or {}
  local sheets = parse_workbook(wbxml, relsxml)
  local datestyle = parse_styles(part("xl/styles.xml"))

  return setmetatable({
    _data = data, _byname = byname, _shared = shared,
    _sheets = sheets, _cache = {}, _datestyle = datestyle,
  }, ReadWB)
end

return xlsx
