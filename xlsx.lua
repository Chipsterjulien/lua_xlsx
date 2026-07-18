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
--- Cible principale : Babet (Lua 5.5). Compatible avec Lua standard 5.3+.
---
--- Le cœur XLSX reste autonome. Lorsque la globale `babet` est disponible,
--- le module utilise automatiquement ses primitives de CRC-32, de validation
--- d'archive et d'écriture atomique. Sans Babet, un chemin de repli Lua pur
--- conserve l'API et applique ses propres contrôles de taille, CRC et ZIP.

local xlsx = { VERSION = "1.2.0" }

local function get_babet()
  local runtime = rawget(_G, "babet")
  return type(runtime) == "table" and runtime or nil
end

local function babet_function(name)
  local runtime = get_babet()
  local fn = runtime and runtime[name]
  return type(fn) == "function" and fn or nil
end

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
  local native = babet_function("crc32")
  if native then
    local ok, digest = pcall(native, s)
    if ok and type(digest) == "string" and digest:match("^%x%x%x%x%x%x%x%x$") then
      return assert(tonumber(digest, 16))
    end
  end
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

local function is_xml_codepoint(cp)
  return cp == 0x09 or cp == 0x0A or cp == 0x0D
    or (cp >= 0x20 and cp <= 0xD7FF)
    or (cp >= 0xE000 and cp <= 0xFFFD)
    or (cp >= 0x10000 and cp <= 0x10FFFF)
end

local function validate_xml_text(s, what, level)
  local count, bad = utf8.len(s)
  if not count then
    error("xlsx: " .. what .. " n'est pas une chaîne UTF-8 valide (octet " .. bad .. ")", level or 3)
  end
  for _, cp in utf8.codes(s) do
    if not is_xml_codepoint(cp) then
      error(string.format("xlsx: %s contient un caractère interdit par XML 1.0 (U+%04X)", what, cp), level or 3)
    end
  end
  return count
end

local function validate_xml_document(s, what)
  validate_xml_text(s, what or "document XML", 0)
  return s
end

local function esc_xml(s, what, level)
  validate_xml_text(s, what or "texte XML", level or 3)
  return (s:gsub('[&<>"]', XML_ESC))
end

local function esc(s)
  return esc_xml(s, "texte de cellule", 3)
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

local EXCEL_EPOCH_1900 = days_from_civil(1899, 12, 30)
local EXCEL_EPOCH_1904 = days_from_civil(1904, 1, 1)

local function epoch_for(date1904)
  return date1904 and EXCEL_EPOCH_1904 or EXCEL_EPOCH_1900
end

local function is_leap_year(y)
  return y % 4 == 0 and (y % 100 ~= 0 or y % 400 == 0)
end

local function validate_date_parts(y, m, d, h, mi, sec, level)
  local values = { y, m, d, h, mi, sec }
  local names = { "année", "mois", "jour", "heure", "minute", "seconde" }
  for i = 1, 6 do
    if math.type(values[i]) ~= "integer" then
      error("xlsx: " .. names[i] .. " doit être un entier", level or 3)
    end
  end
  if y < 1 or y > 9999 then error("xlsx: année hors plage (1..9999)", level or 3) end
  if m < 1 or m > 12 then error("xlsx: mois hors plage (1..12)", level or 3) end
  local mdays = { 31, is_leap_year(y) and 29 or 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 }
  if d < 1 or d > mdays[m] then error("xlsx: jour invalide pour ce mois", level or 3) end
  if h < 0 or h > 23 then error("xlsx: heure hors plage (0..23)", level or 3) end
  if mi < 0 or mi > 59 then error("xlsx: minute hors plage (0..59)", level or 3) end
  if sec < 0 or sec > 59 then error("xlsx: seconde hors plage (0..59)", level or 3) end
end

local function ymd_to_serial(y, m, d, h, mi, sec, date1904)
  local days = days_from_civil(y, m, d) - epoch_for(date1904)
  return days + (h * 3600 + mi * 60 + sec) / 86400
end

-- numéro de série Excel -> chaîne ISO 8601
local function serial_to_iso(serial, withtime, date1904)
  if type(serial) ~= "number" or serial ~= serial or serial == math.huge or serial == -math.huge then
    error("xlsx: serial_to_iso attend un nombre fini", 2)
  end
  local whole = math.floor(serial)
  local frac = serial - whole
  if not withtime then
    local y, m, d = civil_from_days(whole + epoch_for(date1904))
    return string.format("%04d-%02d-%02d", y, m, d)
  end
  local secs = math.floor(frac * 86400 + 0.5)
  if secs >= 86400 then secs = secs - 86400; whole = whole + 1 end
  local y, m, d = civil_from_days(whole + epoch_for(date1904))
  return string.format("%04d-%02d-%02dT%02d:%02d:%02d",
    y, m, d, secs // 3600, (secs % 3600) // 60, secs % 60)
end

local DATE_MT = {}

local function make_date(kind, y, m, d, h, mi, sec, level)
  validate_date_parts(y, m, d, h, mi, sec, level)
  return setmetatable({
    serial = ymd_to_serial(y, m, d, h, mi, sec, false),
    kind = kind, y = y, m = m, d = d, h = h, mi = mi, s = sec,
  }, DATE_MT)
end

--- Crée une valeur date (affichée yyyy-mm-dd dans Excel).
function xlsx.date(y, m, d)
  return make_date("date", y, m, d, 0, 0, 0, 3)
end

--- Crée une valeur date-heure (affichée yyyy-mm-dd hh:mm:ss dans Excel).
function xlsx.datetime(y, m, d, h, mi, sec)
  h, mi, sec = h or 0, mi or 0, sec or 0
  return make_date("datetime", y, m, d, h, mi, sec, 3)
end

xlsx.serial_to_iso = serial_to_iso
function xlsx.date_to_serial(y, m, d, h, mi, sec, date1904)
  h, mi, sec = h or 0, mi or 0, sec or 0
  validate_date_parts(y, m, d, h, mi, sec, 3)
  if date1904 ~= nil and type(date1904) ~= "boolean" then
    error("xlsx: date1904 doit être un booléen", 2)
  end
  return ymd_to_serial(y, m, d, h, mi, sec, date1904)
end

-- ----------------------------------------------------------------------------
-- Formules, hyperliens et styles d'écriture
-- ----------------------------------------------------------------------------
local FORMULA_MT, HYPERLINK_MT, STYLE_MT = {}, {}, {}
local FORMULA_DATA = setmetatable({}, { __mode = "k" })
local HYPERLINK_DATA = setmetatable({}, { __mode = "k" })
local STYLE_DATA = setmetatable({}, { __mode = "k" })

local function copy_border(border)
  if not border then return nil end
  local out = {}
  for _, side in ipairs({ "left", "right", "top", "bottom" }) do
    local src = border[side]
    if src then out[side] = { style = src.style, color = src.color } end
  end
  return next(out) and out or nil
end

FORMULA_MT.__index = function(self, key)
  local data = FORMULA_DATA[self]
  return data and data[key] or nil
end
FORMULA_MT.__newindex = function()
  error("xlsx: une formule est immuable", 2)
end

HYPERLINK_MT.__index = function(self, key)
  local data = HYPERLINK_DATA[self]
  return data and data[key] or nil
end
HYPERLINK_MT.__newindex = function()
  error("xlsx: un hyperlien est immuable", 2)
end

STYLE_MT.__index = function(self, key)
  local data = STYLE_DATA[self]
  if not data then return nil end
  if key == "border" then return copy_border(data.border) end
  return data[key]
end
STYLE_MT.__newindex = function()
  error("xlsx: un style est immuable", 2)
end

local function validate_cached_value(value, level)
  if value == nil then return end
  local t = type(value)
  if t ~= "string" and t ~= "number" and t ~= "boolean" then
    error("xlsx: la valeur mise en cache d'une formule doit être string, number, boolean ou nil", level or 3)
  end
  if t == "string" then validate_xml_text(value, "valeur mise en cache", level or 3) end
  if t == "number" and (value ~= value or value == math.huge or value == -math.huge) then
    error("xlsx: la valeur mise en cache d'une formule doit être finie", level or 3)
  end
end

local function new_formula(data)
  local obj = setmetatable({}, FORMULA_MT)
  FORMULA_DATA[obj] = data
  return obj
end

--- Crée une formule. Le signe = initial est facultatif.
--- cached_value est facultatif et permet d'écrire une valeur mise en cache.
function xlsx.formula(expression, cached_value)
  if type(expression) ~= "string" then error("xlsx: formula attend une string", 2) end
  validate_xml_text(expression, "formule", 3)
  if expression:sub(1, 1) == "=" then expression = expression:sub(2) end
  if expression == "" then error("xlsx: une formule ne peut pas être vide", 2) end
  validate_cached_value(cached_value, 3)
  return new_formula({ expression = expression, cached_value = cached_value, formula_type = "normal" })
end

function xlsx.is_formula(value)
  return getmetatable(value) == FORMULA_MT and FORMULA_DATA[value] ~= nil
end

local function new_hyperlink(data)
  local obj = setmetatable({}, HYPERLINK_MT)
  HYPERLINK_DATA[obj] = data
  return obj
end

local function normalize_hyperlink(target, text, opts, level)
  if type(target) ~= "string" or target == "" then
    error("xlsx: la cible d'un hyperlien doit être une string non vide", level or 3)
  end
  validate_xml_text(target, "cible d'hyperlien", level or 3)
  if #target > 32767 then error("xlsx: la cible d'hyperlien dépasse 32767 octets", level or 3) end
  opts = opts or {}
  if type(opts) ~= "table" then error("xlsx: les options d'hyperlien doivent être une table", level or 3) end
  local allowed = { internal=true, tooltip=true }
  for k in pairs(opts) do
    if not allowed[k] then error("xlsx: option d'hyperlien inconnue : " .. tostring(k), level or 3) end
  end
  local internal = opts.internal == true
  if opts.internal ~= nil and type(opts.internal) ~= "boolean" then
    error("xlsx: internal doit être un booléen", level or 3)
  end
  if text == nil then text = target end
  if type(text) ~= "string" then error("xlsx: le texte d'un hyperlien doit être une string", level or 3) end
  validate_xml_text(text, "texte d'hyperlien", level or 3)
  local tooltip = opts.tooltip
  if tooltip ~= nil then
    if type(tooltip) ~= "string" then error("xlsx: tooltip doit être une string", level or 3) end
    validate_xml_text(tooltip, "infobulle d'hyperlien", level or 3)
    if #tooltip > 255 then error("xlsx: tooltip dépasse 255 octets", level or 3) end
  end
  return { target=target, text=text, internal=internal, tooltip=tooltip }
end

--- Crée une valeur hyperlien pouvant être passée directement à write().
function xlsx.hyperlink(target, text, opts)
  return new_hyperlink(normalize_hyperlink(target, text, opts, 3))
end

function xlsx.is_hyperlink(value)
  return getmetatable(value) == HYPERLINK_MT and HYPERLINK_DATA[value] ~= nil
end

local NUMBER_FORMAT_ALIASES = {
  general = nil,
  integer = "0",
  decimal = "0.00",
  percent = "0.00%",
  currency_eur = '#,##0.00 "€"',
  currency_usd = '$#,##0.00',
  date = "yyyy-mm-dd",
  datetime = "yyyy-mm-dd hh:mm:ss",
}

local BORDER_STYLES = {
  thin=true, medium=true, thick=true, dashed=true, dotted=true, double=true,
}

local function normalize_color(value, what, level)
  if value == nil then return nil end
  if type(value) ~= "string" then error("xlsx: " .. what .. " doit être une couleur RGB/ARGB", level or 3) end
  if value:match("^#%x%x%x%x%x%x$") then value = value:sub(2) end
  if value:match("^%x%x%x%x%x%x$") then value = "FF" .. value end
  if not value:match("^%x%x%x%x%x%x%x%x$") then
    error("xlsx: " .. what .. " doit contenir 6 chiffres RGB ou 8 chiffres ARGB", level or 3)
  end
  return value:upper()
end

local function normalize_border(border, level)
  if border == nil then return nil end
  if type(border) ~= "table" then error("xlsx: border doit être une table", level or 3) end
  local out = {}
  for key in pairs(border) do
    if key ~= "left" and key ~= "right" and key ~= "top" and key ~= "bottom" then
      error("xlsx: côté de bordure inconnu : " .. tostring(key), level or 3)
    end
  end
  for _, side in ipairs({ "left", "right", "top", "bottom" }) do
    local src = border[side]
    if src ~= nil then
      if type(src) ~= "table" then error("xlsx: border." .. side .. " doit être une table", level or 3) end
      for key in pairs(src) do
        if key ~= "style" and key ~= "color" then
          error("xlsx: option de bordure inconnue : " .. side .. "." .. tostring(key), level or 3)
        end
      end
      local style = src.style
      if type(style) ~= "string" or not BORDER_STYLES[style] then
        error("xlsx: border." .. side .. ".style doit valoir thin, medium, thick, dashed, dotted ou double", level or 3)
      end
      out[side] = {
        style = style,
        color = normalize_color(src.color, "border." .. side .. ".color", level or 3),
      }
    end
  end
  return next(out) and out or nil
end

local function normalize_style(opts, level)
  if type(opts) ~= "table" then error("xlsx: style attend une table d'options", level or 3) end
  local allowed = {
    bold=true, italic=true, underline=true, strike=true,
    font_name=true, font_size=true, font_color=true, fill_color=true,
    horizontal=true, vertical=true, wrap_text=true, number_format=true, border=true,
  }
  for k in pairs(opts) do
    if not allowed[k] then error("xlsx: option de style inconnue : " .. tostring(k), level or 3) end
  end
  local out = {}
  for _, key in ipairs({ "bold", "italic", "strike", "wrap_text" }) do
    local value = opts[key]
    if value ~= nil and type(value) ~= "boolean" then
      error("xlsx: " .. key .. " doit être un booléen", level or 3)
    end
    out[key] = value == true
  end
  if opts.underline ~= nil then
    if type(opts.underline) ~= "string" or
        (opts.underline ~= "none" and opts.underline ~= "single" and opts.underline ~= "double") then
      error("xlsx: underline doit valoir none, single ou double", level or 3)
    end
    if opts.underline ~= "none" then out.underline = opts.underline end
  end
  if opts.font_name ~= nil then
    if type(opts.font_name) ~= "string" or opts.font_name == "" then
      error("xlsx: font_name doit être une string non vide", level or 3)
    end
    validate_xml_text(opts.font_name, "nom de police", level or 3)
    if #opts.font_name > 255 then error("xlsx: font_name dépasse 255 octets", level or 3) end
    out.font_name = opts.font_name
  end
  if opts.font_size ~= nil then
    local value = opts.font_size
    if type(value) ~= "number" or value ~= value or value == math.huge or value == -math.huge
        or value < 1 or value > 409 then
      error("xlsx: taille de police doit être un nombre fini entre 1 et 409", level or 3)
    end
    out.font_size = value
  end
  out.font_color = normalize_color(opts.font_color, "font_color", level or 3)
  out.fill_color = normalize_color(opts.fill_color, "fill_color", level or 3)
  if opts.horizontal ~= nil then
    local allowed_h = { left=true, center=true, right=true, justify=true }
    if type(opts.horizontal) ~= "string" or not allowed_h[opts.horizontal] then
      error("xlsx: horizontal doit valoir left, center, right ou justify", level or 3)
    end
    out.horizontal = opts.horizontal
  end
  if opts.vertical ~= nil then
    local allowed_v = { top=true, center=true, bottom=true, justify=true }
    if type(opts.vertical) ~= "string" or not allowed_v[opts.vertical] then
      error("xlsx: vertical doit valoir top, center, bottom ou justify", level or 3)
    end
    out.vertical = opts.vertical
  end
  if opts.number_format ~= nil then
    if type(opts.number_format) ~= "string" or opts.number_format == "" then
      error("xlsx: number_format doit être une string non vide", level or 3)
    end
    validate_xml_text(opts.number_format, "format numérique", level or 3)
    if #opts.number_format > 255 then error("xlsx: number_format dépasse 255 octets", level or 3) end
    if NUMBER_FORMAT_ALIASES[opts.number_format] ~= nil or opts.number_format == "general" then
      out.number_format = NUMBER_FORMAT_ALIASES[opts.number_format]
    else
      out.number_format = opts.number_format
    end
  end
  out.border = normalize_border(opts.border, level or 3)
  return out
end

local function copy_style_data(data)
  return {
    bold = data.bold == true, italic = data.italic == true,
    underline = data.underline, strike = data.strike == true,
    font_name = data.font_name, font_size = data.font_size,
    font_color = data.font_color, fill_color = data.fill_color,
    horizontal = data.horizontal, vertical = data.vertical,
    wrap_text = data.wrap_text == true, number_format = data.number_format,
    border = copy_border(data.border),
  }
end

local function style_object_from_data(data)
  local obj = setmetatable({}, STYLE_MT)
  STYLE_DATA[obj] = copy_style_data(data)
  return obj
end

--- Crée un style réutilisable et immuable.
function xlsx.style(opts)
  return style_object_from_data(normalize_style(opts or {}, 3))
end

function xlsx.is_style(value)
  return getmetatable(value) == STYLE_MT and STYLE_DATA[value] ~= nil
end

--- Renvoie une copie des options d'un style.
function xlsx.style_options(style)
  if not xlsx.is_style(style) then error("xlsx: style_options attend un style xlsx", 2) end
  return copy_style_data(STYLE_DATA[style])
end

local function require_style(style, level)
  if style == nil then return nil end
  if not xlsx.is_style(style) then
    error("xlsx: le style doit être créé avec xlsx.style", level or 3)
  end
  return STYLE_DATA[style]
end

local function key_part(value)
  value = value == nil and "" or tostring(value)
  return #value .. ":" .. value
end

local function border_key(border)
  if not border then return "" end
  local parts = {}
  for _, side in ipairs({ "left", "right", "top", "bottom" }) do
    local item = border[side]
    parts[#parts + 1] = side .. "=" .. (item and (key_part(item.style) .. key_part(item.color)) or "")
  end
  return table.concat(parts, ";")
end

local function style_key(data)
  return table.concat({
    data.bold and "1" or "0", data.italic and "1" or "0",
    key_part(data.underline), data.strike and "1" or "0",
    key_part(data.font_name), key_part(data.font_size),
    key_part(data.font_color), key_part(data.fill_color),
    key_part(data.horizontal), key_part(data.vertical),
    data.wrap_text and "1" or "0", key_part(data.number_format),
    key_part(border_key(data.border)),
  }, "|")
end

local function new_style_registry()
  local reg = {
    fonts = { { bold=false, italic=false, underline=nil, strike=false,
      name="Calibri", size=11, color=nil } },
    font_map = {},
    fills = { { kind="none" }, { kind="gray125" } },
    fill_map = {},
    borders = { {} }, border_map = { [""] = 0 },
    numfmts = {}, numfmt_map = {},
    xfs = { { fontId=0, fillId=0, borderId=0, numFmtId=0 } },
    xf_map = {},
  }
  reg.font_map[table.concat({"0","0",key_part(nil),"0",key_part("Calibri"),key_part(11),key_part(nil)}, "|")] = 0
  reg.xf_map[style_key({ bold=false, italic=false, strike=false, wrap_text=false })] = 0
  return reg
end

local function font_key(data)
  return table.concat({
    data.bold and "1" or "0", data.italic and "1" or "0", key_part(data.underline),
    data.strike and "1" or "0", key_part(data.font_name or "Calibri"),
    key_part(data.font_size or 11), key_part(data.font_color),
  }, "|")
end

local function register_font(reg, data)
  local key = font_key(data)
  local id = reg.font_map[key]
  if id ~= nil then return id end
  id = #reg.fonts
  reg.fonts[#reg.fonts + 1] = {
    bold=data.bold, italic=data.italic, underline=data.underline, strike=data.strike,
    name=data.font_name or "Calibri", size=data.font_size or 11, color=data.font_color,
  }
  reg.font_map[key] = id
  return id
end

local function register_fill(reg, color)
  if not color then return 0 end
  local id = reg.fill_map[color]
  if id ~= nil then return id end
  id = #reg.fills
  reg.fills[#reg.fills + 1] = { kind="solid", color=color }
  reg.fill_map[color] = id
  return id
end

local function register_border(reg, border)
  local key = border_key(border)
  local id = reg.border_map[key]
  if id ~= nil then return id end
  id = #reg.borders
  reg.borders[#reg.borders + 1] = copy_border(border) or {}
  reg.border_map[key] = id
  return id
end

local function register_numfmt(reg, code)
  if not code then return 0 end
  local id = reg.numfmt_map[code]
  if id then return id end
  id = 163 + #reg.numfmts + 1
  reg.numfmts[#reg.numfmts + 1] = { id=id, code=code }
  reg.numfmt_map[code] = id
  return id
end

local function register_style(reg, data)
  local key = style_key(data)
  local known = reg.xf_map[key]
  if known ~= nil then return known end
  local fontId = register_font(reg, data)
  local fillId = register_fill(reg, data.fill_color)
  local borderId = register_border(reg, data.border)
  local numFmtId = register_numfmt(reg, data.number_format)
  local id = #reg.xfs
  reg.xfs[#reg.xfs + 1] = {
    fontId=fontId, fillId=fillId, borderId=borderId, numFmtId=numFmtId,
    horizontal=data.horizontal, vertical=data.vertical, wrap_text=data.wrap_text,
  }
  reg.xf_map[key] = id
  return id
end

local function effective_style_data(style, value)
  local source = style and STYLE_DATA[style] or nil
  local data = source and copy_style_data(source)
    or { bold=false, italic=false, strike=false, wrap_text=false }
  if getmetatable(value) == DATE_MT and not data.number_format then
    data.number_format = value.kind == "datetime"
      and NUMBER_FORMAT_ALIASES.datetime or NUMBER_FORMAT_ALIASES.date
  end
  return data
end

local function build_styles_xml(reg)
  local b = {}
  b[#b + 1] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  b[#b + 1] = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
  if #reg.numfmts > 0 then
    b[#b + 1] = '<numFmts count="' .. #reg.numfmts .. '">'
    for _, fmt in ipairs(reg.numfmts) do
      b[#b + 1] = '<numFmt numFmtId="' .. fmt.id .. '" formatCode="' ..
        esc_xml(fmt.code, "format numérique", 0) .. '"/>'
    end
    b[#b + 1] = '</numFmts>'
  end
  b[#b + 1] = '<fonts count="' .. #reg.fonts .. '">'
  for _, font in ipairs(reg.fonts) do
    b[#b + 1] = '<font>'
    if font.bold then b[#b + 1] = '<b/>' end
    if font.italic then b[#b + 1] = '<i/>' end
    if font.underline == "single" then b[#b + 1] = '<u/>' end
    if font.underline == "double" then b[#b + 1] = '<u val="double"/>' end
    if font.strike then b[#b + 1] = '<strike/>' end
    if font.color then b[#b + 1] = '<color rgb="' .. font.color .. '"/>' end
    b[#b + 1] = '<sz val="' .. num2str(font.size) .. '"/><name val="' ..
      esc_xml(font.name, "nom de police", 0) .. '"/></font>'
  end
  b[#b + 1] = '</fonts>'
  b[#b + 1] = '<fills count="' .. #reg.fills .. '">'
  for _, fill in ipairs(reg.fills) do
    if fill.kind == "none" then
      b[#b + 1] = '<fill><patternFill patternType="none"/></fill>'
    elseif fill.kind == "gray125" then
      b[#b + 1] = '<fill><patternFill patternType="gray125"/></fill>'
    else
      b[#b + 1] = '<fill><patternFill patternType="solid"><fgColor rgb="' .. fill.color ..
        '"/><bgColor indexed="64"/></patternFill></fill>'
    end
  end
  b[#b + 1] = '</fills>'
  b[#b + 1] = '<borders count="' .. #reg.borders .. '">'
  for _, border in ipairs(reg.borders) do
    b[#b + 1] = '<border>'
    for _, side in ipairs({ "left", "right", "top", "bottom" }) do
      local item = border[side]
      if item then
        b[#b + 1] = '<' .. side .. ' style="' .. item.style .. '">'
        if item.color then b[#b + 1] = '<color rgb="' .. item.color .. '"/>' end
        b[#b + 1] = '</' .. side .. '>'
      else
        b[#b + 1] = '<' .. side .. '/>'
      end
    end
    b[#b + 1] = '<diagonal/></border>'
  end
  b[#b + 1] = '</borders>'
  b[#b + 1] = '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
  b[#b + 1] = '<cellXfs count="' .. #reg.xfs .. '">'
  for _, xf in ipairs(reg.xfs) do
    local attrs = {
      'numFmtId="' .. xf.numFmtId .. '"', 'fontId="' .. xf.fontId .. '"',
      'fillId="' .. xf.fillId .. '"', 'borderId="' .. xf.borderId .. '"', 'xfId="0"',
    }
    if xf.numFmtId ~= 0 then attrs[#attrs + 1] = 'applyNumberFormat="1"' end
    if xf.fontId ~= 0 then attrs[#attrs + 1] = 'applyFont="1"' end
    if xf.fillId ~= 0 then attrs[#attrs + 1] = 'applyFill="1"' end
    if xf.borderId ~= 0 then attrs[#attrs + 1] = 'applyBorder="1"' end
    local has_alignment = xf.horizontal or xf.vertical or xf.wrap_text
    if has_alignment then attrs[#attrs + 1] = 'applyAlignment="1"' end
    if has_alignment then
      b[#b + 1] = '<xf ' .. table.concat(attrs, " ") .. '><alignment'
      if xf.horizontal then b[#b + 1] = ' horizontal="' .. xf.horizontal .. '"' end
      if xf.vertical then b[#b + 1] = ' vertical="' .. xf.vertical .. '"' end
      if xf.wrap_text then b[#b + 1] = ' wrapText="1"' end
      b[#b + 1] = '/></xf>'
    else
      b[#b + 1] = '<xf ' .. table.concat(attrs, " ") .. '/>'
    end
  end
  b[#b + 1] = '</cellXfs>'
  b[#b + 1] = '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
  b[#b + 1] = '</styleSheet>'
  return table.concat(b)
end

-- ----------------------------------------------------------------------------
-- Worksheet
-- ----------------------------------------------------------------------------
local Sheet = {}
Sheet.__index = Sheet

local MAX_ROW = 1048575
local MAX_COL = 16383
local function check_index(v, what, level)
  local max = what == "row" and MAX_ROW or MAX_COL
  if math.type(v) ~= "integer" or v < 0 or v > max then
    error("xlsx: " .. what .. " doit être un entier entre 0 et " .. max, level)
  end
end

local function check_finite_number(value, what, min, max, level)
  if type(value) ~= "number" or value ~= value or value == math.huge or value == -math.huge
      or value < min or value > max then
    error(string.format("xlsx: %s doit être un nombre fini entre %s et %s", what, min, max), level or 3)
  end
end

local function col_number(letters)
  local n = 0
  letters = letters:gsub("%$", ""):upper()
  if letters == "" then return nil end
  for i = 1, #letters do
    local byte = letters:byte(i)
    if byte < 65 or byte > 90 then return nil end
    n = n * 26 + byte - 64
  end
  return n - 1
end

local function parse_a1_range(ref, what, allow_single, level)
  if type(ref) ~= "string" then error("xlsx: " .. what .. " doit être une string A1:B2", level or 3) end
  local c1, r1, c2, r2 = ref:match("^%$?([A-Za-z]+)%$?(%d+):%$?([A-Za-z]+)%$?(%d+)$")
  if not c1 and allow_single then
    c1, r1 = ref:match("^%$?([A-Za-z]+)%$?(%d+)$")
    c2, r2 = c1, r1
  end
  if not c1 then error("xlsx: " .. what .. " invalide : " .. ref, level or 3) end
  local a, b = col_number(c1), col_number(c2)
  local x, y = tonumber(r1) - 1, tonumber(r2) - 1
  if not a or not b or a < 0 or b > MAX_COL or x < 0 or y > MAX_ROW or a > b or x > y then
    error("xlsx: " .. what .. " hors limites ou inversée : " .. ref, level or 3)
  end
  return x, a, y, b, col_ref(a) .. (x + 1) .. ":" .. col_ref(b) .. (y + 1)
end

local function ensure_cell_extent(sheet, row, col)
  if row > sheet.maxrow then sheet.maxrow = row end
  if col > sheet.maxcol then sheet.maxcol = col end
end

local function new_sheet(name, workbook)
  return setmetatable({
    name   = name,
    cells  = {},
    styles = {},
    hyperlinks = {},
    row_heights = {},
    column_widths = {},
    merged_ranges = {},
    maxrow = -1,
    maxcol = -1,
    _workbook = workbook,
    _freeze_rows = 0,
    _freeze_cols = 0,
    _auto_filter = false,
  }, Sheet)
end

local function set_hyperlink_data(sheet, row, col, data)
  local rr = sheet.hyperlinks[row]
  if not rr then rr = {}; sheet.hyperlinks[row] = rr end
  rr[col] = { target=data.target, internal=data.internal, tooltip=data.tooltip }
  ensure_cell_extent(sheet, row, col)
end

--- Écrit une cellule. row/col sont 0-indexés.
--- value : string|number|boolean|date|formula|hyperlink|nil. style est facultatif.
function Sheet:write(row, col, value, style)
  check_index(row, "row", 3)
  check_index(col, "col", 3)
  require_style(style, 3)
  local t = type(value)
  local mt = t == "table" and getmetatable(value) or nil
  if mt == HYPERLINK_MT then
    local link = HYPERLINK_DATA[value]
    value = link.text
    t = "string"
    mt = nil
    set_hyperlink_data(self, row, col, link)
  end
  if mt == DATE_MT or mt == FORMULA_MT then
    -- valeur taggée : acceptée telle quelle
  elseif value ~= nil and t ~= "string" and t ~= "number" and t ~= "boolean" then
    error("xlsx: type de cellule non supporté : " .. t, 2)
  elseif t == "string" then
    validate_xml_text(value, "texte de cellule", 3)
  end
  if t == "number" and (value ~= value or value == math.huge or value == -math.huge) then
    error("xlsx: impossible d'écrire NaN/Inf dans une cellule", 2)
  end
  local r = self.cells[row]
  if not r then r = {}; self.cells[row] = r end
  r[col] = value
  if style ~= nil then
    local sr = self.styles[row]
    if not sr then sr = {}; self.styles[row] = sr end
    sr[col] = style
  end
  if value ~= nil or style ~= nil then ensure_cell_extent(self, row, col) end
  return self
end

--- Applique ou retire le style d'une cellule, y compris une cellule vide.
function Sheet:set_style(row, col, style)
  check_index(row, "row", 3)
  check_index(col, "col", 3)
  require_style(style, 3)
  local sr = self.styles[row]
  if style == nil then
    if sr then sr[col] = nil end
  else
    if not sr then sr = {}; self.styles[row] = sr end
    sr[col] = style
    ensure_cell_extent(self, row, col)
  end
  return self
end

--- Ajoute ou remplace un hyperlien sur une cellule existante.
function Sheet:set_hyperlink(row, col, target, opts)
  check_index(row, "row", 3)
  check_index(col, "col", 3)
  local data = normalize_hyperlink(target, target, opts, 3)
  set_hyperlink_data(self, row, col, data)
  return self
end

--- Supprime l'hyperlien d'une cellule sans modifier sa valeur.
function Sheet:remove_hyperlink(row, col)
  check_index(row, "row", 3)
  check_index(col, "col", 3)
  local rr = self.hyperlinks[row]
  if rr then rr[col] = nil end
  return self
end

--- Définit la largeur d'une colonne 0-indexée, en unités Excel.
function Sheet:set_column_width(col, width)
  check_index(col, "col", 3)
  check_finite_number(width, "largeur de colonne", 0.1, 255, 3)
  self.column_widths[col] = width
  return self
end

--- Définit la hauteur d'une ligne 0-indexée, en points.
function Sheet:set_row_height(row, height)
  check_index(row, "row", 3)
  check_finite_number(height, "hauteur de ligne", 0.1, 409.5, 3)
  self.row_heights[row] = height
  if row > self.maxrow then self.maxrow = row end
  return self
end

--- Fige un nombre de lignes et de colonnes à partir du coin supérieur gauche.
function Sheet:freeze_panes(rows, cols)
  rows, cols = rows or 0, cols or 0
  if math.type(rows) ~= "integer" or rows < 0 or rows > MAX_ROW then
    error("xlsx: rows doit être un entier entre 0 et " .. MAX_ROW, 2)
  end
  if math.type(cols) ~= "integer" or cols < 0 or cols > MAX_COL then
    error("xlsx: cols doit être un entier entre 0 et " .. MAX_COL, 2)
  end
  self._freeze_rows, self._freeze_cols = rows, cols
  return self
end

local function validate_filter_ref(ref)
  local _, _, _, _, normalized = parse_a1_range(ref, "plage de filtre", false, 3)
  return normalized
end

--- Active un filtre automatique. Sans plage, utilise la zone de données courante.
--- Passer false désactive le filtre.
function Sheet:set_auto_filter(ref)
  if ref == false then
    self._auto_filter = false
  elseif ref == nil then
    self._auto_filter = true
  else
    self._auto_filter = validate_filter_ref(ref)
  end
  return self
end

local function ranges_overlap(a, b)
  return not (a.r2 < b.r1 or b.r2 < a.r1 or a.c2 < b.c1 or b.c2 < a.c1)
end

--- Fusionne une plage. Accepte "A1:D1" ou quatre coordonnées 0-indexées.
function Sheet:merge_cells(a, b, c, d)
  local r1, c1, r2, c2, ref
  if type(a) == "string" and b == nil then
    r1, c1, r2, c2, ref = parse_a1_range(a, "plage fusionnée", false, 3)
  else
    if a == nil or b == nil or c == nil or d == nil then
      error("xlsx: merge_cells attend une plage ou quatre coordonnées", 2)
    end
    check_index(a, "row", 3); check_index(b, "col", 3)
    check_index(c, "row", 3); check_index(d, "col", 3)
    if a > c or b > d then error("xlsx: plage fusionnée inversée", 2) end
    r1, c1, r2, c2 = a, b, c, d
    ref = col_ref(c1) .. (r1 + 1) .. ":" .. col_ref(c2) .. (r2 + 1)
  end
  if r1 == r2 and c1 == c2 then error("xlsx: une fusion doit couvrir au moins deux cellules", 2) end
  local candidate = { r1=r1, c1=c1, r2=r2, c2=c2, ref=ref }
  for _, current in ipairs(self.merged_ranges) do
    if ranges_overlap(candidate, current) then
      error("xlsx: la fusion " .. ref .. " chevauche " .. current.ref, 2)
    end
  end
  self.merged_ranges[#self.merged_ranges + 1] = candidate
  table.sort(self.merged_ranges, function(x, y)
    return x.r1 ~= y.r1 and x.r1 < y.r1 or (x.c1 ~= y.c1 and x.c1 < y.c1 or x.ref < y.ref)
  end)
  for row = r1, r2 do
    for col = c1, c2 do
      if row ~= r1 or col ~= c1 then
        if self.cells[row] then self.cells[row][col] = nil end
        if self.styles[row] then self.styles[row][col] = nil end
        if self.hyperlinks[row] then self.hyperlinks[row][col] = nil end
      end
    end
  end
  ensure_cell_extent(self, r2, c2)
  return self
end

--- Retire une fusion exacte. Les anciennes valeurs des cellules ne sont pas restaurées.
function Sheet:unmerge_cells(ref)
  local _, _, _, _, normalized = parse_a1_range(ref, "plage fusionnée", false, 3)
  for i, current in ipairs(self.merged_ranges) do
    if current.ref == normalized then
      table.remove(self.merged_ranges, i)
      return self
    end
  end
  error("xlsx: fusion introuvable : " .. normalized, 2)
end

--- Ajoute une ligne complète à la suite. Un style facultatif s'applique aux valeurs présentes.
function Sheet:append_row(values, style)
  if type(values) ~= "table" then error("xlsx: append_row attend une table", 2) end
  require_style(style, 3)
  local row = self.maxrow + 1
  for i = 1, #values do
    local v = values[i]
    if v ~= nil then self:write(row, i - 1, v, style) end
  end
  if not self.cells[row] then self.cells[row] = {} end
  if row > self.maxrow then self.maxrow = row end
  return self
end

--- Écrit une matrice. Un style facultatif s'applique à toutes les valeurs présentes.
function Sheet:write_rows(matrix, style)
  if type(matrix) ~= "table" then error("xlsx: write_rows attend une table de lignes", 2) end
  require_style(style, 3)
  for i = 1, #matrix do self:append_row(matrix[i], style) end
  return self
end

local function style_attr(id)
  return id ~= 0 and (' s="' .. id .. '"') or ""
end

local function freeze_pane_xml(rows, cols)
  if rows == 0 and cols == 0 then return nil end
  local attrs = {}
  if cols > 0 then attrs[#attrs + 1] = 'xSplit="' .. cols .. '"' end
  if rows > 0 then attrs[#attrs + 1] = 'ySplit="' .. rows .. '"' end
  attrs[#attrs + 1] = 'topLeftCell="' .. col_ref(cols) .. (rows + 1) .. '"'
  local pane = rows > 0 and cols > 0 and "bottomRight" or (rows > 0 and "bottomLeft" or "topRight")
  attrs[#attrs + 1] = 'activePane="' .. pane .. '"'
  attrs[#attrs + 1] = 'state="frozen"'
  return '<sheetViews><sheetView workbookViewId="0"><pane ' ..
    table.concat(attrs, " ") .. '/></sheetView></sheetViews>'
end

local function cached_formula_xml(data)
  local value = data.cached_value
  if value == nil then return "", "" end
  local t = type(value)
  if t == "string" then return ' t="str"', '<v>' .. esc_xml(value, "valeur de formule", 0) .. '</v>' end
  if t == "boolean" then return ' t="b"', '<v>' .. (value and "1" or "0") .. '</v>' end
  return "", '<v>' .. num2str(value) .. '</v>'
end

local function worksheet_relationships_xml(rels)
  if #rels == 0 then return nil end
  local b = {
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
  }
  for _, rel in ipairs(rels) do
    b[#b + 1] = '<Relationship Id="' .. rel.id ..
      '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' ..
      esc_xml(rel.target, "cible d'hyperlien", 0) .. '" TargetMode="External"/>'
  end
  b[#b + 1] = '</Relationships>'
  return table.concat(b)
end

-- sérialise la feuille en XML ; alimente sst et renvoie aussi ses relations.
function Sheet:_xml(sst, sst_index)
  local b = {}
  b[#b + 1] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  b[#b + 1] = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
  if self.maxrow >= 0 and self.maxcol >= 0 then
    b[#b + 1] = '<dimension ref="A1:' .. col_ref(self.maxcol) .. (self.maxrow + 1) .. '"/>'
  end
  local pane = freeze_pane_xml(self._freeze_rows, self._freeze_cols)
  if pane then b[#b + 1] = pane end
  if next(self.column_widths) then
    b[#b + 1] = '<cols>'
    local cols = {}
    for col in pairs(self.column_widths) do cols[#cols + 1] = col end
    table.sort(cols)
    for _, col in ipairs(cols) do
      b[#b + 1] = '<col min="' .. (col + 1) .. '" max="' .. (col + 1) ..
        '" width="' .. num2str(self.column_widths[col]) .. '" customWidth="1"/>'
    end
    b[#b + 1] = '</cols>'
  end
  b[#b + 1] = '<sheetData>'
  for r = 0, self.maxrow do
    local rowcells, rowstyles = self.cells[r], self.styles[r]
    if rowcells or rowstyles or self.row_heights[r] then
      local rowattrs = { 'r="' .. (r + 1) .. '"' }
      if self.row_heights[r] then
        rowattrs[#rowattrs + 1] = 'ht="' .. num2str(self.row_heights[r]) .. '"'
        rowattrs[#rowattrs + 1] = 'customHeight="1"'
      end
      b[#b + 1] = '<row ' .. table.concat(rowattrs, " ") .. '>'
      for c = 0, self.maxcol do
        local v, style
        if rowcells then v = rowcells[c] end
        if rowstyles then style = rowstyles[c] end
        if v ~= nil or style ~= nil then
          local ref = col_ref(c) .. (r + 1)
          local sid = self._workbook:_style_id(style, v)
          local sattr = style_attr(sid)
          local tv = type(v)
          local mt = tv == "table" and getmetatable(v) or nil
          if mt == DATE_MT then
            local date1904 = self._workbook and self._workbook._date1904 or false
            local serial = ymd_to_serial(v.y, v.m, v.d, v.h, v.mi, v.s, date1904)
            b[#b + 1] = '<c r="' .. ref .. '"' .. sattr .. '><v>' .. num2str(serial) .. '</v></c>'
          elseif mt == FORMULA_MT then
            local formula = FORMULA_DATA[v]
            if not formula.expression or formula.expression == "" then
              error("xlsx: impossible d'écrire une formule partagée sans expression résolue", 0)
            end
            local tattr, cached = cached_formula_xml(formula)
            local fattrs = ""
            if formula.formula_type and formula.formula_type ~= "normal" then
              fattrs = ' t="' .. formula.formula_type .. '"'
              if formula.ref then fattrs = fattrs .. ' ref="' .. esc_xml(formula.ref, "référence de formule", 0) .. '"' end
              if formula.shared_index ~= nil then fattrs = fattrs .. ' si="' .. formula.shared_index .. '"' end
            end
            b[#b + 1] = '<c r="' .. ref .. '"' .. sattr .. tattr .. '><f' .. fattrs .. '>' ..
              esc_xml(formula.expression, "formule", 0) .. '</f>' .. cached .. '</c>'
          elseif tv == "string" then
            local idx = sst_index[v]
            if idx == nil then idx = #sst; sst[#sst + 1] = v; sst_index[v] = idx end
            b[#b + 1] = '<c r="' .. ref .. '"' .. sattr .. ' t="s"><v>' .. idx .. '</v></c>'
          elseif tv == "boolean" then
            b[#b + 1] = '<c r="' .. ref .. '"' .. sattr .. ' t="b"><v>' .. (v and "1" or "0") .. '</v></c>'
          elseif tv == "number" then
            b[#b + 1] = '<c r="' .. ref .. '"' .. sattr .. '><v>' .. num2str(v) .. '</v></c>'
          else
            b[#b + 1] = '<c r="' .. ref .. '"' .. sattr .. '/>'
          end
        end
      end
      b[#b + 1] = '</row>'
    end
  end
  b[#b + 1] = '</sheetData>'
  local filter_ref
  if self._auto_filter == true and self.maxrow >= 0 and self.maxcol >= 0 then
    filter_ref = 'A1:' .. col_ref(self.maxcol) .. (self.maxrow + 1)
  elseif type(self._auto_filter) == "string" then
    filter_ref = self._auto_filter
  end
  if filter_ref then b[#b + 1] = '<autoFilter ref="' .. filter_ref .. '"/>' end
  if #self.merged_ranges > 0 then
    b[#b + 1] = '<mergeCells count="' .. #self.merged_ranges .. '">'
    for _, range in ipairs(self.merged_ranges) do b[#b + 1] = '<mergeCell ref="' .. range.ref .. '"/>' end
    b[#b + 1] = '</mergeCells>'
  end

  local rels, links = {}, {}
  local rows = {}
  for row in pairs(self.hyperlinks) do rows[#rows + 1] = row end
  table.sort(rows)
  for _, row in ipairs(rows) do
    local cols = {}
    for col in pairs(self.hyperlinks[row]) do cols[#cols + 1] = col end
    table.sort(cols)
    for _, col in ipairs(cols) do
      local link = self.hyperlinks[row][col]
      local attrs = { 'ref="' .. col_ref(col) .. (row + 1) .. '"' }
      if link.tooltip then attrs[#attrs + 1] = 'tooltip="' .. esc_xml(link.tooltip, "infobulle", 0) .. '"' end
      if link.internal then
        attrs[#attrs + 1] = 'location="' .. esc_xml(link.target, "cible interne", 0) .. '"'
      else
        local id = "rId" .. (#rels + 1)
        rels[#rels + 1] = { id=id, target=link.target }
        attrs[#attrs + 1] = 'r:id="' .. id .. '"'
      end
      links[#links + 1] = '<hyperlink ' .. table.concat(attrs, " ") .. '/>'
    end
  end
  if #links > 0 then b[#b + 1] = '<hyperlinks>' .. table.concat(links) .. '</hyperlinks>' end
  b[#b + 1] = '</worksheet>'
  return table.concat(b), worksheet_relationships_xml(rels)
end

-- ----------------------------------------------------------------------------
-- Workbook
-- ----------------------------------------------------------------------------
local Workbook = {}
Workbook.__index = Workbook

function Workbook:_style_id(style, value)
  require_style(style, 3)
  if not self._style_registry then error("xlsx: registre de styles indisponible", 2) end
  return register_style(self._style_registry, effective_style_data(style, value))
end

function xlsx.new(opts)
  opts = opts or {}
  if type(opts) ~= "table" then error("xlsx: new attend une table d'options", 2) end
  for k in pairs(opts) do
    if k ~= "date_system" then error("xlsx: option inconnue pour new : " .. tostring(k), 2) end
  end
  local date_system = opts.date_system or "1900"
  if date_system ~= "1900" and date_system ~= "1904" then
    error("xlsx: date_system doit valoir '1900' ou '1904'", 2)
  end
  return setmetatable({
    sheets = {}, _date1904 = date_system == "1904", _sheet_names = {},
    _style_registry = nil,
  }, Workbook)
end

function Workbook:add_sheet(name)
  if name == nil then name = "Sheet" .. (#self.sheets + 1) end
  if type(name) ~= "string" then error("xlsx: le nom de feuille doit être une string", 2) end
  local count = validate_xml_text(name, "nom de feuille", 3)
  if count == 0 or count > 31 then error("xlsx: nom de feuille invalide (1 à 31 caractères Unicode)", 2) end
  if name:find('[:\\/?*%[%]]') then
    error("xlsx: le nom de feuille contient un caractère interdit ( : \\ / ? * [ ] )", 2)
  end
  if name:sub(1, 1) == "'" or name:sub(-1) == "'" then
    error("xlsx: le nom de feuille ne peut pas commencer ou finir par une apostrophe", 2)
  end
  local key = name:lower()
  if self._sheet_names[key] then error("xlsx: nom de feuille dupliqué : " .. name, 2) end
  local sh = new_sheet(name, self)
  self.sheets[#self.sheets + 1] = sh
  self._sheet_names[key] = true
  return sh
end

-- construit la liste des parties ZIP {name=, data=}
function Workbook:_parts()
  local nsheets = #self.sheets
  self._style_registry = new_style_registry()
  local sst, sst_index = {}, {}
  local sheet_xmls, sheet_rels = {}, {}
  for i = 1, nsheets do
    sheet_xmls[i], sheet_rels[i] = self.sheets[i]:_xml(sst, sst_index)
  end

  local parts = {}
  local function add(name, data) parts[#parts + 1] = { name = name, data = data } end

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

  add("_rels/.rels", table.concat({
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    '</Relationships>',
  }))

  do
    local t = {}
    t[#t + 1] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    t[#t + 1] = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    if self._date1904 then t[#t + 1] = '<workbookPr date1904="1"/>' end
    t[#t + 1] = '<sheets>'
    for i = 1, nsheets do
      t[#t + 1] = '<sheet name="' .. esc_xml(self.sheets[i].name, "nom de feuille", 0) ..
        '" sheetId="' .. i .. '" r:id="rId' .. i .. '"/>'
    end
    t[#t + 1] = '</sheets></workbook>'
    add("xl/workbook.xml", table.concat(t))
  end

  do
    local t = {}
    t[#t + 1] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    t[#t + 1] = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    for i = 1, nsheets do
      t[#t + 1] = '<Relationship Id="rId' .. i ..
        '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' .. i .. '.xml"/>'
    end
    t[#t + 1] = '<Relationship Id="rId' .. (nsheets + 1) ..
      '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    t[#t + 1] = '<Relationship Id="rId' .. (nsheets + 2) ..
      '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    t[#t + 1] = '</Relationships>'
    add("xl/_rels/workbook.xml.rels", table.concat(t))
  end

  for i = 1, nsheets do
    add("xl/worksheets/sheet" .. i .. ".xml", sheet_xmls[i])
    if sheet_rels[i] then add("xl/worksheets/_rels/sheet" .. i .. ".xml.rels", sheet_rels[i]) end
  end

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

  add("xl/styles.xml", build_styles_xml(self._style_registry))
  self._style_registry = nil
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

local temp_counter = 0

local function write_all(file, bytes)
  local ok, err = file:write(bytes)
  if not ok then return nil, err or "écriture incomplète" end
  local flushed, flush_err = file:flush()
  if not flushed then return nil, flush_err or "échec de flush" end
  return true
end

local function fallback_atomic_write(path, bytes, overwrite)
  if not overwrite then
    local existing = io.open(path, "rb")
    if existing then existing:close(); return nil, "destination existante" end
  end
  temp_counter = temp_counter + 1
  local tmp = string.format("%s.xlsx-tmp-%d-%d-%d", path, os.time(), temp_counter, math.random(0, 0x7fffffff))
  local f, err = io.open(tmp, "wb")
  if not f then return nil, "impossible de créer le temporaire : " .. tostring(err) end
  local ok, write_err = write_all(f, bytes)
  local closed, close_err = f:close()
  if not ok or not closed then
    os.remove(tmp)
    return nil, tostring(write_err or close_err or "échec de fermeture")
  end
  local renamed, rename_err = os.rename(tmp, path)
  if not renamed then
    os.remove(tmp)
    return nil, "publication atomique impossible : " .. tostring(rename_err)
  end
  return true
end

--- Écrit le classeur dans `path`.
--- opts : overwrite=true, durable=true, permissions=0644, use_babet=true.
function Workbook:save(path, opts)
  if type(path) ~= "string" then error("xlsx: le chemin de sauvegarde doit être une string", 2) end
  opts = opts or {}
  if type(opts) ~= "table" then error("xlsx: save attend une table d'options", 2) end
  local allowed = { overwrite=true, durable=true, permissions=true, use_babet=true }
  for k in pairs(opts) do if not allowed[k] then error("xlsx: option inconnue pour save : " .. tostring(k), 2) end end
  local overwrite = opts.overwrite
  if overwrite == nil then overwrite = true elseif type(overwrite) ~= "boolean" then error("xlsx: overwrite doit être un booléen", 2) end
  local durable = opts.durable
  if durable == nil then durable = true elseif type(durable) ~= "boolean" then error("xlsx: durable doit être un booléen", 2) end
  local permissions = opts.permissions or tonumber("644", 8)
  if math.type(permissions) ~= "integer" or permissions < 0 or permissions > tonumber("777", 8) then
    error("xlsx: permissions doit être un entier entre 0000 et 0777", 2)
  end
  local use_babet = opts.use_babet
  if use_babet == nil then use_babet = true elseif type(use_babet) ~= "boolean" then error("xlsx: use_babet doit être un booléen", 2) end
  if #self.sheets == 0 then self:add_sheet("Sheet1") end

  local ok_build, bytes = pcall(function() return build_zip(self:_parts()) end)
  if not ok_build then return nil, tostring(bytes) end

  local native = use_babet and babet_function("writeFileAtomic") or nil
  if native then
    local ok_call, ok, err = pcall(native, path, bytes, {
      overwrite = overwrite, permissions = permissions, durable = durable,
    })
    if not ok_call then return nil, "xlsx: writeFileAtomic a levé une erreur : " .. tostring(ok) end
    if not ok then return nil, "xlsx: " .. tostring(err) end
    return true
  end

  local ok, err = fallback_atomic_write(path, bytes, overwrite)
  if not ok then return nil, "xlsx: impossible d'écrire '" .. path .. "' : " .. tostring(err) end
  return true
end

-- ----------------------------------------------------------------------------
-- Raccourci pratique
-- ----------------------------------------------------------------------------
--- Écrit une matrice directement dans un fichier. opts.sheet = nom de feuille.
function xlsx.write_rows(path, matrix, opts)
  opts = opts or {}
  if type(opts) ~= "table" then error("xlsx: write_rows attend une table d'options", 2) end
  local allowed = { sheet=true, date_system=true, overwrite=true, durable=true, permissions=true, use_babet=true }
  for k in pairs(opts) do if not allowed[k] then error("xlsx: option inconnue pour write_rows : " .. tostring(k), 2) end end
  local wb = xlsx.new({ date_system = opts.date_system or "1900" })
  local sh = wb:add_sheet(opts.sheet or "Sheet1")
  sh:write_rows(matrix)
  return wb:save(path, {
    overwrite = opts.overwrite, durable = opts.durable,
    permissions = opts.permissions, use_babet = opts.use_babet,
  })
end

-- ============================================================================
-- LECTURE
-- ============================================================================
-- Lecture : valeurs, dates, formules, styles, dimensions, fusions, volets,
-- filtres et hyperliens. read() conserve son contrat historique et renvoie la
-- valeur mise en cache ; get_formula() expose séparément l'expression.

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

local function inflate(input, max_output)
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

  local function decode(counts, symbols)
    local code, first, index = 0, 0, 0
    for len = 1, 15 do
      code = code | getbits(1)
      local count = counts[len]
      if code - count < first then return symbols[index + (code - first)] end
      index = index + count
      first = (first + count) << 1
      code = code << 1
    end
    error("xlsx: inflate: code Huffman invalide", 0)
  end

  local out, olen = {}, 0
  local function reserve(count)
    if count < 0 or olen + count > max_output then
      error("xlsx: entrée décompressée trop volumineuse", 0)
    end
  end

  local function inflate_block(lc, ls, dc, ds)
    while true do
      local sym = decode(lc, ls)
      if sym == nil then error("xlsx: inflate: symbole absent", 0) end
      if sym == 256 then return end
      if sym < 256 then
        reserve(1); olen = olen + 1; out[olen] = sym
      else
        if sym < 257 or sym > 285 then error("xlsx: inflate: longueur invalide", 0) end
        local li = sym - 256
        local length = LEN_BASE[li] + getbits(LEN_EXTRA[li])
        local dsym = decode(dc, ds)
        if dsym == nil or dsym < 0 or dsym > 29 then error("xlsx: inflate: distance invalide", 0) end
        local distance = DIST_BASE[dsym + 1] + getbits(DIST_EXTRA[dsym + 1])
        if distance < 1 or distance > olen then error("xlsx: inflate: référence arrière invalide", 0) end
        reserve(length)
        local start = olen - distance
        for k = 1, length do olen = olen + 1; out[olen] = out[start + k] end
      end
    end
  end

  local fixed_lit = {}
  for sym = 0, 143 do fixed_lit[sym] = 8 end
  for sym = 144, 255 do fixed_lit[sym] = 9 end
  for sym = 256, 279 do fixed_lit[sym] = 7 end
  for sym = 280, 287 do fixed_lit[sym] = 8 end
  local fixed_dist = {}
  for sym = 0, 29 do fixed_dist[sym] = 5 end

  repeat
    local bfinal = getbits(1)
    local btype = getbits(2)
    if btype == 0 then
      local drop = bitcnt % 8
      bitbuf = bitbuf >> drop; bitcnt = bitcnt - drop
      local len = getbits(16)
      local nlen = getbits(16)
      if (len ~ 0xFFFF) ~= nlen then error("xlsx: inflate: bloc stocké corrompu", 0) end
      reserve(len)
      for _ = 1, len do olen = olen + 1; out[olen] = getbits(8) end
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
      local all, i, total = {}, 0, hlit + hdist
      while i < total do
        local sym = decode(clc, cls)
        if sym < 16 then
          all[i] = sym; i = i + 1
        elseif sym == 16 then
          if i == 0 then error("xlsx: inflate: répétition sans longueur précédente", 0) end
          local prev, count = all[i - 1], getbits(2) + 3
          if i + count > total then error("xlsx: inflate: longueurs débordantes", 0) end
          for _ = 1, count do all[i] = prev; i = i + 1 end
        elseif sym == 17 then
          local count = getbits(3) + 3
          if i + count > total then error("xlsx: inflate: longueurs débordantes", 0) end
          for _ = 1, count do all[i] = 0; i = i + 1 end
        elseif sym == 18 then
          local count = getbits(7) + 11
          if i + count > total then error("xlsx: inflate: longueurs débordantes", 0) end
          for _ = 1, count do all[i] = 0; i = i + 1 end
        else
          error("xlsx: inflate: code de longueur invalide", 0)
        end
      end
      local litl, distl = {}, {}
      for k = 0, hlit - 1 do litl[k] = all[k] end
      for k = 0, hdist - 1 do distl[k] = all[hlit + k] end
      if (litl[256] or 0) == 0 then error("xlsx: inflate: symbole de fin absent", 0) end
      local lc, ls = huff_build(litl, hlit)
      local dc, ds = huff_build(distl, hdist)
      inflate_block(lc, ls, dc, ds)
    else
      error("xlsx: inflate: type de bloc réservé", 0)
    end
  until bfinal == 1

  if pos <= n then error("xlsx: inflate: octets finaux étrangers", 0) end
  local parts, CHUNK = {}, 2000
  for i = 1, olen, CHUNK do
    local j = math.min(i + CHUNK - 1, olen)
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

local function require_range(s, pos, count, what)
  if pos < 1 or count < 0 or pos + count - 1 > #s then
    error("xlsx: " .. what .. " hors limites", 0)
  end
end

local function safe_zip_name(name, max_path_length)
  if #name == 0 or #name > max_path_length or not utf8.len(name) or name:find("\0", 1, true)
      or name:sub(1, 1) == "/" or name:find("\\", 1, true) then return false end
  local first = name:match("^([^/]+)") or ""
  if first:match("^[A-Za-z]:") then return false end
  for part in name:gmatch("[^/]+") do
    if part == "." or part == ".." then return false end
  end
  return not name:find("//", 1, true)
end

local function read_zip(s, limits)
  local eocd = find_eocd(s)
  require_range(s, eocd, 22, "EOCD ZIP")
  local sig, disk, cd_disk, n_disk, n_total, cd_size, cd_off, comment_len =
    string.unpack("<I4 I2 I2 I2 I2 I4 I4 I2", s, eocd)
  if sig ~= SIG_EOCD then error("xlsx: EOCD ZIP invalide", 0) end
  if disk ~= 0 or cd_disk ~= 0 or n_disk ~= n_total then error("xlsx: ZIP multidisque non supporté", 0) end
  if n_total == 0xFFFF or cd_size == 0xFFFFFFFF or cd_off == 0xFFFFFFFF then
    error("xlsx: ZIP64 non supporté par le lecteur Lua", 0)
  end
  if n_total > limits.max_entries then error("xlsx: trop d'entrées ZIP", 0) end
  if eocd + 22 + comment_len - 1 ~= #s then error("xlsx: commentaire ou données finales ZIP incohérents", 0) end
  local cd_start, cd_end = cd_off + 1, cd_off + cd_size
  if cd_start < 1 or cd_end >= eocd or cd_end > #s then error("xlsx: répertoire central ZIP hors limites", 0) end

  local byname, entries, total_size, total_name_bytes, pos = {}, {}, 0, 0, cd_start
  for index = 1, n_total do
    require_range(s, pos, 46, "entrée du répertoire central")
    local csig, _vm, _vn, flags, method, _mt, _md, crc, csize, usize,
          namelen, extralen, commentlen, diskstart, _ia, _ea, lhoff, np =
      string.unpack("<I4 I2 I2 I2 I2 I2 I2 I4 I4 I4 I2 I2 I2 I2 I2 I4 I4", s, pos)
    if csig ~= SIG_CENTRAL then error("xlsx: répertoire central ZIP corrompu", 0) end
    require_range(s, np, namelen + extralen + commentlen, "nom/métadonnées ZIP")
    local name = s:sub(np, np + namelen - 1)
    if not safe_zip_name(name, limits.max_path_length) then error("xlsx: nom d'entrée ZIP dangereux : " .. name, 0) end
    if byname[name] then error("xlsx: entrée ZIP dupliquée : " .. name, 0) end
    total_name_bytes = total_name_bytes + #name
    if total_name_bytes > limits.max_total_name_bytes then error("xlsx: budget cumulé des noms ZIP dépassé", 0) end
    if diskstart ~= 0 then error("xlsx: entrée ZIP située sur un autre disque", 0) end
    if (flags & 0x0001) ~= 0 then error("xlsx: entrée ZIP chiffrée non supportée", 0) end
    if method ~= 0 and method ~= 8 then error("xlsx: méthode de compression ZIP non supportée : " .. method, 0) end
    if usize > limits.max_entry_size then error("xlsx: entrée ZIP trop volumineuse : " .. name, 0) end
    total_size = total_size + usize
    if total_size > limits.max_total_size then error("xlsx: taille totale décompressée trop grande", 0) end
    if usize > 0 and csize == 0 then error("xlsx: taille compressée nulle pour une entrée non vide", 0) end
    if usize > 0 and usize / math.max(csize, 1) > limits.max_compression_ratio then
      error("xlsx: rapport de compression excessif : " .. name, 0)
    end
    local e = { name=name, flags=flags, method=method, crc=crc, csize=csize, usize=usize,
      lhoff=lhoff, index=index }
    byname[name], entries[#entries + 1] = e, e
    pos = np + namelen + extralen + commentlen
  end
  if pos ~= cd_end + 1 then error("xlsx: taille du répertoire central ZIP incohérente", 0) end

  local ranges = {}
  for _, e in ipairs(entries) do
    local lp = e.lhoff + 1
    require_range(s, lp, 30, "en-tête local ZIP")
    local lsig, _v, lflags, lmethod, _t, _d, lcrc, lcsize, lusize, lnamelen, lextralen, dp =
      string.unpack("<I4 I2 I2 I2 I2 I2 I4 I4 I4 I2 I2", s, lp)
    if lsig ~= SIG_LOCAL then error("xlsx: en-tête local ZIP corrompu : " .. e.name, 0) end
    require_range(s, dp, lnamelen + lextralen, "nom local ZIP")
    local lname = s:sub(dp, dp + lnamelen - 1)
    if lname ~= e.name or lmethod ~= e.method or lflags ~= e.flags then
      error("xlsx: incohérence entre en-tête local et central : " .. e.name, 0)
    end
    if (e.flags & 0x0008) == 0 and (lcrc ~= e.crc or lcsize ~= e.csize or lusize ~= e.usize) then
      error("xlsx: tailles ou CRC locaux incohérents : " .. e.name, 0)
    end
    local data_start = dp + lnamelen + lextralen
    local data_end = data_start + e.csize - 1
    if e.csize > 0 then require_range(s, data_start, e.csize, "données ZIP") end
    local local_end = math.max(dp + lnamelen + lextralen - 1, data_end)
    if (e.flags & 0x0008) ~= 0 then
      local desc = data_end + 1
      require_range(s, desc, 12, "data descriptor ZIP")
      local first = string.unpack("<I4", s, desc)
      local dcrc, dcsize, dusize
      if first == 0x08074b50 then
        require_range(s, desc, 16, "data descriptor ZIP")
        local descriptor_sig
        descriptor_sig, dcrc, dcsize, dusize = string.unpack("<I4 I4 I4 I4", s, desc)
        local_end = desc + 15
      else
        dcrc, dcsize, dusize = string.unpack("<I4 I4 I4", s, desc)
        local_end = desc + 11
      end
      if dcrc ~= e.crc or dcsize ~= e.csize or dusize ~= e.usize then
        error("xlsx: data descriptor ZIP incohérent : " .. e.name, 0)
      end
    end
    if local_end >= cd_start then error("xlsx: entrée ZIP chevauchant le répertoire central", 0) end
    e.data_start, e.data_end, e.local_start, e.local_end = data_start, data_end, lp, local_end
    ranges[#ranges + 1] = e
  end
  table.sort(ranges, function(a,b) return a.local_start < b.local_start end)
  for i = 2, #ranges do
    if ranges[i].local_start <= ranges[i - 1].local_end then error("xlsx: entrées ZIP chevauchantes", 0) end
  end
  return byname
end

local function zip_extract(s, e, limits)
  local comp = e.csize == 0 and "" or s:sub(e.data_start, e.data_end)
  local data
  if e.method == 0 then
    if e.csize ~= e.usize then error("xlsx: entrée STORED avec tailles différentes : " .. e.name, 0) end
    data = comp
  else
    data = inflate(comp, math.min(e.usize, limits.max_entry_size))
  end
  if #data ~= e.usize then error("xlsx: taille décompressée incohérente : " .. e.name, 0) end
  if crc32(data) ~= e.crc then error("xlsx: CRC-32 invalide : " .. e.name, 0) end
  return data
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
    local dec = e:match("^#(%d+)$")
    local cp = hex and tonumber(hex, 16) or (dec and tonumber(dec) or nil)
    if cp then
      if not is_xml_codepoint(cp) then error("xlsx: entité XML vers un caractère interdit", 0) end
      return utf8.char(cp)
    end
    return "&" .. e .. ";"
  end))
end

local function ref_to_rowcol(ref)
  local letters, digits = ref:match("^(%a+)(%d+)$")
  if not letters then return nil end
  letters = letters:upper()
  local col = 0
  for i = 1, #letters do col = col * 26 + (string.byte(letters, i) - 64) end
  local row = tonumber(digits) - 1
  col = col - 1
  if row < 0 or row > MAX_ROW or col < 0 or col > MAX_COL then
    error("xlsx: référence de cellule hors limites : " .. ref, 0)
  end
  return row, col
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
    local raw = inner:match("<v>(.-)</v>")
    if not raw then return nil end
    local index = tonumber(raw)
    if not index or math.type(index) ~= "integer" or index < 0 or shared[index] == nil then
      error("xlsx: index sharedStrings invalide : " .. tostring(raw), 0)
    end
    return shared[index]
  elseif ty == "b" then
    local raw = inner:match("<v>(.-)</v>")
    if raw == "1" then return true end
    if raw == "0" then return false end
    error("xlsx: booléen de cellule invalide", 0)
  elseif ty == "inlineStr" then
    local buf = {}
    for text in inner:gmatch("<t[^>]*>(.-)</t>") do buf[#buf + 1] = unescape(text) end
    return table.concat(buf)
  elseif ty == "str" or ty == "d" or ty == "e" then
    local raw = inner:match("<v>(.-)</v>")
    return raw and unescape(raw) or nil
  else
    local raw = inner:match("<v>(.-)</v>")
    if raw == nil or raw == "" then return nil end
    local value = tonumber(raw)
    if value == nil then error("xlsx: valeur numérique invalide : " .. raw, 0) end
    return value
  end
end

-- analyse styles.xml -> styles pris en charge et classification date/datetime
local BUILTIN_DATE = {
  [14] = "date", [15] = "date", [16] = "date", [17] = "date",
  [18] = "datetime", [19] = "datetime", [20] = "datetime", [21] = "datetime",
  [22] = "datetime", [45] = "datetime", [46] = "datetime", [47] = "datetime",
}
local BUILTIN_NUMFMT = {
  [0]="General", [1]="0", [2]="0.00", [3]="#,##0", [4]="#,##0.00",
  [9]="0%", [10]="0.00%", [11]="0.00E+00", [12]="# ?/?", [13]="# ??/??",
  [14]="mm-dd-yy", [15]="d-mmm-yy", [16]="d-mmm", [17]="mmm-yy",
  [18]="h:mm AM/PM", [19]="h:mm:ss AM/PM", [20]="h:mm", [21]="h:mm:ss",
  [22]="m/d/yy h:mm", [45]="mm:ss", [46]="[h]:mm:ss", [47]="mmss.0",
  [49]="@",
}

local function xml_attr(attrs, name)
  return attrs and unescape(attrs:match(name .. '="([^"]*)"')) or nil
end

local function parse_repeated_elements(section, name)
  local out, p = {}, 1
  while true do
    local a, b = section:find("<" .. name .. "[%s>/]", p)
    if not a then break end
    local tagend = assert(section:find(">", b))
    local tag = section:sub(a, tagend)
    if tag:sub(-2) == "/>" then
      out[#out + 1] = { attrs=tag:match("^<" .. name .. "%s*(.-)/>$") or "", body="" }
      p = tagend + 1
    else
      local close_a, close_b = section:find("</" .. name .. ">", tagend + 1, true)
      if not close_a then error("xlsx: élément XML non fermé : " .. name, 0) end
      out[#out + 1] = {
        attrs=tag:match("^<" .. name .. "%s*(.-)>$") or "",
        body=section:sub(tagend + 1, close_a - 1),
      }
      p = close_b + 1
    end
  end
  return out
end

local function parse_border_side(body, side)
  local attrs, inner = body:match("<" .. side .. "%s*(.-)>(.-)</" .. side .. ">")
  if not attrs then attrs = body:match("<" .. side .. "%s*(.-)/>"); inner = "" end
  if attrs == nil then return nil end
  local style = xml_attr(attrs, "style")
  if not style or not BORDER_STYLES[style] then return nil end
  local color = inner:match('<color[^>]-rgb="([%x]+)"')
  return { style=style, color=color and normalize_color(color, "couleur de bordure", 0) or nil }
end

local function parse_styles(xml)
  local datestyle, style_objects = {}, { [0] = false }
  if not xml then return datestyle, style_objects end
  local fmtcode = {}
  for tag in xml:gmatch("<numFmt%s.-/>") do
    local id = tonumber(tag:match('numFmtId="(%d+)"'))
    local code = tag:match('formatCode="([^"]*)"')
    if id and code then fmtcode[id] = unescape(code) end
  end
  local function classify(id)
    local builtin = BUILTIN_DATE[id]
    if builtin then return builtin end
    local code = fmtcode[id]
    if not code then return nil end
    local c = code:gsub('"[^"]*"', ""):gsub("%[[^%]]*%]", ""):gsub("\\.", ""):lower()
    local has_date = c:find("[yd]") or c:find("mmm")
    local has_time = c:find("[hs]")
    if has_time then return "datetime" end
    if has_date then return "date" end
    return nil
  end

  local fonts, fills, borders = {}, {}, {}
  local font_section = xml:match("<fonts[^>]*>(.-)</fonts>") or ""
  for i, element in ipairs(parse_repeated_elements(font_section, "font")) do
    local body = element.body
    local uattrs = body:match("<u%s*(.-)/>")
    local underline
    if uattrs ~= nil then underline = xml_attr(uattrs, "val") == "double" and "double" or "single" end
    local color = body:match('<color[^>]-rgb="([%x]+)"')
    fonts[i - 1] = {
      bold=body:find("<b[%s/]", 1) ~= nil,
      italic=body:find("<i[%s/]", 1) ~= nil,
      underline=underline,
      strike=body:find("<strike[%s/]", 1) ~= nil,
      font_name=unescape(body:match('<name[^>]-val="([^"]*)"') or "Calibri"),
      font_size=tonumber(body:match('<sz[^>]-val="([^"]*)"') or "11"),
      font_color=color and normalize_color(color, "couleur de police", 0) or nil,
    }
  end
  local fill_section = xml:match("<fills[^>]*>(.-)</fills>") or ""
  for i, element in ipairs(parse_repeated_elements(fill_section, "fill")) do
    local body = element.body
    local color = body:match('<fgColor[^>]-rgb="([%x]+)"')
    fills[i - 1] = color and normalize_color(color, "couleur de fond", 0) or nil
  end
  local border_section = xml:match("<borders[^>]*>(.-)</borders>") or ""
  for i, element in ipairs(parse_repeated_elements(border_section, "border")) do
    local border = {}
    for _, side in ipairs({ "left", "right", "top", "bottom" }) do
      border[side] = parse_border_side(element.body, side)
    end
    borders[i - 1] = next(border) and border or nil
  end

  local cellxfs = xml:match("<cellXfs[^>]*>(.-)</cellXfs>") or ""
  for idx, element in ipairs(parse_repeated_elements(cellxfs, "xf")) do
    local attrs, body = element.attrs, element.body
    local numFmtId = tonumber(xml_attr(attrs, "numFmtId") or "0")
    local font = fonts[tonumber(xml_attr(attrs, "fontId") or "0")] or {}
    local fill = fills[tonumber(xml_attr(attrs, "fillId") or "0")]
    local border = borders[tonumber(xml_attr(attrs, "borderId") or "0")]
    local alignment_attrs = body:match("<alignment%s*(.-)/>") or body:match("<alignment%s*(.-)>")
    local data = {
      bold=font.bold == true, italic=font.italic == true, underline=font.underline,
      strike=font.strike == true, font_name=font.font_name, font_size=font.font_size,
      font_color=font.font_color, fill_color=fill,
      horizontal=xml_attr(alignment_attrs, "horizontal"),
      vertical=xml_attr(alignment_attrs, "vertical"),
      wrap_text=xml_attr(alignment_attrs, "wrapText") == "1" or xml_attr(alignment_attrs, "wrapText") == "true",
      number_format=fmtcode[numFmtId] or (numFmtId ~= 0 and BUILTIN_NUMFMT[numFmtId] or nil),
      border=copy_border(border),
    }
    local style_index = idx - 1
    if classify(numFmtId) then datestyle[style_index] = classify(numFmtId) end
    local is_default = style_index == 0 or (not data.bold and not data.italic and not data.underline and
      not data.strike and (not data.font_name or data.font_name == "Calibri") and
      (not data.font_size or data.font_size == 11) and not data.font_color and not data.fill_color and
      not data.horizontal and not data.vertical and not data.wrap_text and not data.number_format and not data.border)
    style_objects[style_index] = is_default and false or style_object_from_data(data)
  end
  return datestyle, style_objects
end

local function parse_relationships(xml)
  local result = {}
  if not xml then return result end
  for tag in xml:gmatch("<Relationship%s.-/>") do
    local id = tag:match('Id="([^"]*)"')
    if id then
      result[id] = {
        target=unescape(tag:match('Target="([^"]*)"') or ""),
        target_mode=tag:match('TargetMode="([^"]*)"'),
        rel_type=tag:match('Type="([^"]*)"'),
      }
    end
  end
  return result
end

local function parse_sheet(xml, shared, datestyle, style_objects, date1904, relationships)
  local cells, formulas, styles = {}, {}, {}
  local row_heights, column_widths = {}, {}
  local maxrow, maxcol = -1, -1

  for tag in xml:gmatch("<col%s.-/>") do
    local first, last = tonumber(tag:match('min="(%d+)"')), tonumber(tag:match('max="(%d+)"'))
    local width = tonumber(tag:match('width="([^"]+)"'))
    if first and last and width and first >= 1 and last <= MAX_COL + 1 and first <= last then
      for col = first - 1, last - 1 do column_widths[col] = width end
    end
  end

  for crow in xml:gmatch("<row[%s>].-</row>") do
    local rowtag = crow:match("^(<row[^>]*>)") or ""
    local rownum = tonumber(rowtag:match('r="(%d+)"'))
    local height = tonumber(rowtag:match('ht="([^"]+)"'))
    if rownum and height and rownum >= 1 and rownum <= MAX_ROW + 1 then row_heights[rownum - 1] = height end
    local p = 1
    while true do
      local cs = crow:find("<c[%s/>]", p)
      if not cs then break end
      local tagend = assert(crow:find(">", cs))
      local tag = crow:sub(cs, tagend)
      local selfclose = tag:sub(-2) == "/>"
      local ref = tag:match('r="([^"]*)"')
      local ty = tag:match('t="([^"]*)"')
      local st = tonumber(tag:match('s="(%d+)"'))
      local inner, nextp
      if selfclose then inner, nextp = "", tagend + 1
      else
        local close = crow:find("</c>", tagend, true)
        if not close then error("xlsx: cellule XML non fermée", 0) end
        inner, nextp = crow:sub(tagend + 1, close - 1), close + 4
      end
      if ref then
        local r0, c0 = ref_to_rowcol(ref)
        if r0 then
          if r0 > maxrow then maxrow = r0 end
          if c0 > maxcol then maxcol = c0 end
          if st and st ~= 0 then
            local sr = styles[r0]; if not sr then sr = {}; styles[r0] = sr end
            sr[c0] = style_objects[st] or false
          end
          local raw_value = decode_cell(inner, ty, shared)
          local display_value = raw_value
          if type(display_value) == "number" and st then
            local cls = datestyle[st]
            if cls then display_value = serial_to_iso(display_value, cls == "datetime", date1904) end
          end
          local fattrs, fexpr = inner:match("<f%s*(.-)>(.-)</f>")
          if not fattrs then fattrs = inner:match("<f%s*(.-)/>") end
          if fattrs ~= nil then
            local ftype = xml_attr(fattrs, "t") or "normal"
            local expression = fexpr and unescape(fexpr) or nil
            local formula = new_formula({
              expression=expression, cached_value=raw_value, formula_type=ftype,
              ref=xml_attr(fattrs, "ref"), shared_index=tonumber(xml_attr(fattrs, "si")),
            })
            local fr = formulas[r0]; if not fr then fr = {}; formulas[r0] = fr end
            fr[c0] = formula
          end
          if display_value ~= nil then
            local row = cells[r0]; if not row then row = {}; cells[r0] = row end
            row[c0] = display_value
          end
        end
      end
      p = nextp
    end
  end

  local freeze_rows, freeze_cols = 0, 0
  local pane = xml:match("<pane%s.-/>")
  if pane and (pane:match('state="([^"]+)"') == "frozen" or pane:match('state="([^"]+)"') == "frozenSplit") then
    freeze_cols = tonumber(pane:match('xSplit="([^"]+)"')) or 0
    freeze_rows = tonumber(pane:match('ySplit="([^"]+)"')) or 0
    freeze_cols, freeze_rows = math.floor(freeze_cols), math.floor(freeze_rows)
  end
  local auto_filter = xml:match('<autoFilter[^>]-ref="([^"]+)"')
  if auto_filter then auto_filter = unescape(auto_filter) end
  local merged = {}
  for tag in xml:gmatch("<mergeCell%s.-/>") do
    local ref = tag:match('ref="([^"]+)"')
    if ref then
      local r1, c1, r2, c2, normalized = parse_a1_range(unescape(ref), "fusion lue", false, 0)
      merged[#merged + 1] = { ref=normalized, r1=r1, c1=c1, r2=r2, c2=c2 }
    end
  end
  local hyperlink_ranges = {}
  for tag in xml:gmatch("<hyperlink%s.-/>") do
    local ref = tag:match('ref="([^"]+)"')
    if ref then
      local r1, c1, r2, c2, normalized = parse_a1_range(unescape(ref), "hyperlien lu", true, 0)
      local rid = tag:match('r:id="([^"]+)"')
      local location = tag:match('location="([^"]*)"')
      local relation = rid and relationships[rid] or nil
      local target = location and unescape(location) or (relation and relation.target or nil)
      if target then
        hyperlink_ranges[#hyperlink_ranges + 1] = {
          ref=normalized, r1=r1, c1=c1, r2=r2, c2=c2, target=target,
          internal=location ~= nil, tooltip=unescape(tag:match('tooltip="([^"]*)"') or ""),
        }
        if hyperlink_ranges[#hyperlink_ranges].tooltip == "" then hyperlink_ranges[#hyperlink_ranges].tooltip = nil end
      end
    end
  end
  return cells, formulas, styles, maxrow, maxcol, row_heights, column_widths,
    freeze_rows, freeze_cols, auto_filter, merged, hyperlink_ranges
end

local function package_path(base, target)
  if target:find("\\", 1, true) or target:find("\0", 1, true) then return nil end
  local raw = target:sub(1,1) == "/" and target:sub(2) or (base .. target)
  local out = {}
  for part in raw:gmatch("[^/]+") do
    if part == "." then
      -- ignore
    elseif part == ".." then
      if #out == 0 then return nil end
      out[#out] = nil
    elseif part ~= "" then
      out[#out + 1] = part
    end
  end
  return table.concat(out, "/")
end

local function parse_workbook(wbxml, relsxml)
  local relationships = parse_relationships(relsxml)
  local rid2target = {}
  for id, rel in pairs(relationships) do rid2target[id] = rel.target end
  local sheets, seen = {}, {}
  for tag in wbxml:gmatch("<sheet%s.-/>") do
    local name = unescape(tag:match('name="([^"]*)"') or "")
    local name_count = validate_xml_text(name, "nom de feuille lu", 0)
    if name_count == 0 or name_count > 31 or name:find('[:\\/?*%[%]]')
        or name:sub(1,1) == "'" or name:sub(-1) == "'" then
      error("xlsx: nom de feuille invalide dans le classeur : " .. name, 0)
    end
    local name_key = name:lower()
    if seen[name_key] then error("xlsx: nom de feuille dupliqué dans le classeur : " .. name, 0) end
    seen[name_key] = true
    local rid = tag:match('r:id="([^"]*)"')
    sheets[#sheets + 1] = { name = name, rid = rid }
  end
  for _, sh in ipairs(sheets) do
    local target = rid2target[sh.rid]
    if target then
      sh.path = package_path("xl/", target)
      if not sh.path then error("xlsx: cible de feuille dangereuse", 0) end
    end
  end
  local workbook_pr = wbxml:match("<workbookPr%s.-/>") or wbxml:match("<workbookPr%s.-</workbookPr>") or ""
  local date1904 = workbook_pr:match('date1904="([^"]+)"')
  date1904 = date1904 == "1" or date1904 == "true"
  return sheets, date1904
end

-- ----------------------------------------------------------------------------
-- Objets de lecture
-- ----------------------------------------------------------------------------
local ReadSheet = {}
ReadSheet.__index = ReadSheet

--- Lit la valeur mise en cache d'une cellule. nil si vide ou sans cache.
function ReadSheet:read(r, c)
  check_index(r, "row", 3); check_index(c, "col", 3)
  local row = self.cells[r]
  if not row then return nil end
  return row[c]
end

--- Renvoie l'objet formule d'une cellule, ou nil.
function ReadSheet:get_formula(r, c)
  check_index(r, "row", 3); check_index(c, "col", 3)
  local row = self.formulas[r]
  return row and row[c] or nil
end

--- Renvoie le style pris en charge de la cellule, ou nil pour le style par défaut.
function ReadSheet:get_style(r, c)
  check_index(r, "row", 3); check_index(c, "col", 3)
  local row = self.styles[r]
  local style = row and row[c] or nil
  return style == false and nil or style
end

function ReadSheet:get_column_width(col)
  check_index(col, "col", 3)
  return self.column_widths[col]
end

function ReadSheet:get_row_height(row)
  check_index(row, "row", 3)
  return self.row_heights[row]
end

function ReadSheet:get_frozen_panes()
  return self.freeze_rows, self.freeze_cols
end

function ReadSheet:get_auto_filter()
  return self.auto_filter
end

function ReadSheet:merged_cells()
  local out = {}
  for i, range in ipairs(self.merged_ranges) do out[i] = range.ref end
  return out
end

function ReadSheet:get_hyperlink(row, col)
  check_index(row, "row", 3); check_index(col, "col", 3)
  for _, link in ipairs(self.hyperlink_ranges) do
    if row >= link.r1 and row <= link.r2 and col >= link.c1 and col <= link.c2 then
      return { target=link.target, internal=link.internal, tooltip=link.tooltip, ref=link.ref }
    end
  end
  return nil
end

function ReadSheet:hyperlinks()
  local out = {}
  for i, link in ipairs(self.hyperlink_ranges) do
    out[i] = { target=link.target, internal=link.internal, tooltip=link.tooltip, ref=link.ref }
  end
  return out
end

--- Renvoie maxrow, maxcol (0-indexés ; -1, -1 si la feuille est vide).
function ReadSheet:dims()
  return self.maxrow, self.maxcol
end

--- Itérateur de lignes sur les valeurs mises en cache.
function ReadSheet:rows()
  local r = -1
  return function()
    r = r + 1
    if r > self.maxrow then return nil end
    local src = self.cells[r]
    local row = { n = self.maxcol + 1 }
    if src then for c = 0, self.maxcol do row[c + 1] = src[c] end end
    return row
  end
end

local ReadWB = {}
ReadWB.__index = ReadWB

function ReadWB:date_system()
  return self._date1904 and "1904" or "1900"
end

function ReadWB:sheet_names()
  local t = {}
  for i, sh in ipairs(self._sheets) do t[i] = sh.name end
  return t
end

local function rels_path_for_part(part)
  local dir, file = part:match("^(.-)([^/]+)$")
  if not dir then return "_rels/" .. part .. ".rels" end
  return dir .. "_rels/" .. file .. ".rels"
end

function ReadWB:sheet(which)
  local def
  if type(which) == "number" then
    if math.type(which) ~= "integer" or which < 1 then error("xlsx: l'index de feuille doit être un entier >= 1", 2) end
    def = self._sheets[which]
  elseif type(which) == "string" then
    for _, sh in ipairs(self._sheets) do if sh.name == which then def = sh; break end end
  else
    error("xlsx: sheet attend un nom (string) ou un index (number)", 2)
  end
  if not def then return nil end
  if self._cache[def] then return self._cache[def] end
  if not def.path then return nil, "xlsx: cible de feuille introuvable" end
  local e = self._byname[def.path]
  if not e then return nil, "xlsx: partie manquante : " .. def.path end
  local ok, result = pcall(function()
    local xml = validate_xml_document(zip_extract(self._data, e, self._limits), def.path)
    local rel_entry = self._byname[rels_path_for_part(def.path)]
    local rel_xml = rel_entry and validate_xml_document(zip_extract(self._data, rel_entry, self._limits), rel_entry.name) or nil
    local relationships = parse_relationships(rel_xml)
    return { parse_sheet(xml, self._shared, self._datestyle, self._style_objects, self._date1904, relationships) }
  end)
  if not ok then return nil, tostring(result) end
  local values = result
  local rs = setmetatable({
    name=def.name, cells=values[1], formulas=values[2], styles=values[3],
    maxrow=values[4], maxcol=values[5], row_heights=values[6], column_widths=values[7],
    freeze_rows=values[8], freeze_cols=values[9], auto_filter=values[10],
    merged_ranges=values[11], hyperlink_ranges=values[12],
  }, ReadSheet)
  self._cache[def] = rs
  return rs
end

local OPEN_DEFAULTS = {
  max_file_size = 256 * 1024 * 1024,
  max_entries = 10000,
  max_entry_size = 64 * 1024 * 1024,
  max_total_size = 512 * 1024 * 1024,
  max_path_length = 4096,
  max_total_name_bytes = 16 * 1024 * 1024,
  max_compression_ratio = 200,
  use_babet = true,
  validate_archive = true,
}

local function open_options(opts)
  opts = opts or {}
  if type(opts) ~= "table" then error("xlsx: open attend une table d'options", 3) end
  local out = {}
  for k, default in pairs(OPEN_DEFAULTS) do out[k] = opts[k] == nil and default or opts[k] end
  for k in pairs(opts) do if OPEN_DEFAULTS[k] == nil then error("xlsx: option inconnue pour open : " .. tostring(k), 3) end end
  for _, k in ipairs({"max_file_size","max_entries","max_entry_size","max_total_size","max_path_length","max_total_name_bytes"}) do
    if math.type(out[k]) ~= "integer" or out[k] <= 0 then error("xlsx: " .. k .. " doit être un entier positif", 3) end
  end
  local ceilings = {
    max_entries = 100000, max_entry_size = 8 * 1024^3,
    max_total_size = 64 * 1024^3, max_path_length = 1024 * 1024,
    max_total_name_bytes = 64 * 1024 * 1024,
  }
  for k, ceiling in pairs(ceilings) do
    if out[k] > ceiling then error("xlsx: " .. k .. " dépasse le plafond " .. ceiling, 3) end
  end
  if type(out.max_compression_ratio) ~= "number" or out.max_compression_ratio < 1
      or out.max_compression_ratio ~= out.max_compression_ratio or out.max_compression_ratio == math.huge
      or out.max_compression_ratio > 1000000000 then
    error("xlsx: max_compression_ratio doit être un nombre fini >= 1", 3)
  end
  for _, k in ipairs({"use_babet","validate_archive"}) do
    if type(out[k]) ~= "boolean" then error("xlsx: " .. k .. " doit être un booléen", 3) end
  end
  return out
end

local function file_size_lua(path)
  local f, err = io.open(path, "rb")
  if not f then return nil, err end
  local size, seek_err = f:seek("end")
  local closed, close_err = f:close()
  if not size then return nil, seek_err end
  if not closed then return nil, close_err end
  return size
end

local function checked_file_size(path, opts)
  local fn = opts.use_babet and babet_function("fileSize") or nil
  if fn then
    local ok, size, err = pcall(fn, path)
    if not ok then return nil, "fileSize a levé une erreur : " .. tostring(size) end
    if not size then return nil, err end
    return size
  end
  return file_size_lua(path)
end

local function validate_with_babet(path, opts)
  local runtime = opts.use_babet and get_babet() or nil
  local archive = runtime and type(runtime.archive) == "table" and runtime.archive or nil
  local fn = archive and archive.test
  if not opts.validate_archive or type(fn) ~= "function" then return true end
  local ok, result, err = pcall(fn, path, {
    max_entries = opts.max_entries,
    max_entry_size = opts.max_entry_size,
    max_total_size = opts.max_total_size,
    max_path_length = opts.max_path_length,
    max_total_name_bytes = opts.max_total_name_bytes,
    max_compression_ratio = opts.max_compression_ratio,
  })
  if not ok then return nil, "archive.test a levé une erreur : " .. tostring(result) end
  if not result then return nil, err end
  if result.format ~= "zip" then return nil, "le conteneur n'est pas un ZIP" end
  return true
end

--- Ouvre un .xlsx en lecture. Renvoie un workbook, ou (nil, err).
function xlsx.open(path, opts)
  if type(path) ~= "string" then error("xlsx: open attend un chemin (string)", 2) end
  local limits = open_options(opts)

  local size, size_err = checked_file_size(path, limits)
  if not size then return nil, "xlsx: impossible d'inspecter '" .. path .. "' : " .. tostring(size_err) end
  if size > limits.max_file_size then return nil, "xlsx: fichier trop volumineux" end

  local valid, valid_err = validate_with_babet(path, limits)
  if not valid then return nil, "xlsx: archive refusée par Babet : " .. tostring(valid_err) end

  local f, err = io.open(path, "rb")
  if not f then return nil, "xlsx: impossible d'ouvrir '" .. path .. "' : " .. tostring(err) end
  local data, read_err = f:read("a")
  local closed, close_err = f:close()
  if not data then return nil, "xlsx: échec de lecture : " .. tostring(read_err) end
  if not closed then return nil, "xlsx: échec de fermeture : " .. tostring(close_err) end
  if #data > limits.max_file_size then return nil, "xlsx: fichier modifié ou trop volumineux pendant la lecture" end

  local ok, result = pcall(function()
    local byname = read_zip(data, limits)
    local function part(name)
      local e = byname[name]
      if not e then return nil end
      return validate_xml_document(zip_extract(data, e, limits), name)
    end
    local wbxml = part("xl/workbook.xml")
    if not wbxml then error("xlsx: xl/workbook.xml manquant (pas un xlsx ?)", 0) end
    local relsxml = part("xl/_rels/workbook.xml.rels") or ""
    local ssxml = part("xl/sharedStrings.xml")
    local shared = ssxml and parse_shared_strings(ssxml) or {}
    local sheets, date1904 = parse_workbook(wbxml, relsxml)
    local datestyle, style_objects = parse_styles(part("xl/styles.xml"))
    return setmetatable({
      _data=data, _byname=byname, _shared=shared, _sheets=sheets,
      _cache={}, _datestyle=datestyle, _style_objects=style_objects,
      _date1904=date1904, _limits=limits,
    }, ReadWB)
  end)
  if not ok then return nil, tostring(result) end
  return result
end

return xlsx
