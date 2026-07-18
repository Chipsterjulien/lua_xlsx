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

local xlsx = { VERSION = "1.5.0" }

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
local FORMULA_MT, HYPERLINK_MT, STYLE_MT, RICH_TEXT_MT = {}, {}, {}, {}
local FORMULA_DATA = setmetatable({}, { __mode = "k" })
local HYPERLINK_DATA = setmetatable({}, { __mode = "k" })
local STYLE_DATA = setmetatable({}, { __mode = "k" })
local RICH_TEXT_DATA = setmetatable({}, { __mode = "k" })

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

RICH_TEXT_MT.__index = function(self, key)
  local data = RICH_TEXT_DATA[self]
  if not data then return nil end
  if key == "runs" then
    local out = {}
    for i, run in ipairs(data.runs) do local copy = {}; for k, v in pairs(run) do copy[k] = v end; out[i] = copy end
    return out
  end
  return data[key]
end
RICH_TEXT_MT.__newindex = function()
  error("xlsx: un texte enrichi est immuable", 2)
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

local function shallow_copy(src)
  local out = {}
  for k, v in pairs(src or {}) do out[k] = v end
  return out
end

local function dense_array(value, what, level)
  if type(value) ~= "table" then error("xlsx: " .. what .. " doit être une table dense", level or 3) end
  local n = #value
  for k in pairs(value) do
    if math.type(k) ~= "integer" or k < 1 or k > n then
      error("xlsx: " .. what .. " doit être une table dense 1..n", level or 3)
    end
  end
  return n
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
    horizontal=true, vertical=true, wrap_text=true, number_format=true, border=true, locked=true, hidden=true,
  }
  for k in pairs(opts) do
    if not allowed[k] then error("xlsx: option de style inconnue : " .. tostring(k), level or 3) end
  end
  local out = {}
  for _, key in ipairs({ "bold", "italic", "strike", "wrap_text" }) do
    local value = opts[key]
    if value ~= nil and type(value) ~= "boolean" then error("xlsx: " .. key .. " doit être un booléen", level or 3) end
    out[key] = value == true
  end
  for _, key in ipairs({ "locked", "hidden" }) do
    local value = opts[key]
    if value ~= nil and type(value) ~= "boolean" then error("xlsx: " .. key .. " doit être un booléen", level or 3) end
    out[key] = value
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
    locked = data.locked, hidden = data.hidden, border = copy_border(data.border),
  }
end

local function style_object_from_data(data)
  local obj = setmetatable({}, STYLE_MT)
  STYLE_DATA[obj] = copy_style_data(data)
  return obj
end

local function copy_rich_runs(runs)
  local out = {}
  for i, run in ipairs(runs or {}) do out[i] = shallow_copy(run) end
  return out
end

local function normalize_rich_text(runs, level)
  local n = dense_array(runs, "segments de texte enrichi", level or 3)
  if n < 1 or n > 255 then error("xlsx: un texte enrichi doit contenir entre 1 et 255 segments", level or 3) end
  local out, plain = {}, {}
  for i, src in ipairs(runs) do
    if type(src) ~= "table" then error("xlsx: chaque segment enrichi doit être une table", level or 3) end
    local allowed = { text=true, bold=true, italic=true, underline=true, strike=true,
      font_name=true, font_size=true, font_color=true }
    for k in pairs(src) do if not allowed[k] then error("xlsx: option de segment enrichi inconnue : " .. tostring(k), level or 3) end end
    if type(src.text) ~= "string" or src.text == "" then error("xlsx: chaque segment enrichi exige text non vide", level or 3) end
    validate_xml_text(src.text, "segment enrichi", level or 3)
    local style = normalize_style({
      bold=src.bold, italic=src.italic, underline=src.underline, strike=src.strike,
      font_name=src.font_name, font_size=src.font_size, font_color=src.font_color,
    }, level or 3)
    out[i] = {
      text=src.text, bold=style.bold, italic=style.italic, underline=style.underline,
      strike=style.strike, font_name=style.font_name, font_size=style.font_size,
      font_color=style.font_color,
    }
    plain[#plain + 1] = src.text
  end
  return { runs=out, text=table.concat(plain) }
end

--- Crée un texte enrichi immuable, utilisable avec write().
function xlsx.rich_text(runs)
  local data = normalize_rich_text(runs, 3)
  local obj = setmetatable({}, RICH_TEXT_MT)
  RICH_TEXT_DATA[obj] = data
  return obj
end

function xlsx.is_rich_text(value)
  return getmetatable(value) == RICH_TEXT_MT and RICH_TEXT_DATA[value] ~= nil
end

--- Renvoie une copie des segments d'un texte enrichi.
function xlsx.rich_text_runs(value)
  if not xlsx.is_rich_text(value) then error("xlsx: rich_text_runs attend un texte enrichi xlsx", 2) end
  return copy_rich_runs(RICH_TEXT_DATA[value].runs)
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
    key_part(data.locked), key_part(data.hidden), key_part(border_key(data.border)),
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
    dxfs = {}, dxf_map = {},
  }
  reg.font_map[table.concat({"0","0",key_part(nil),"0",key_part("Calibri"),key_part(11),key_part(nil)}, "|")] = 0
  reg.xf_map[style_key({ bold=false, italic=false, strike=false, wrap_text=false, locked=nil, hidden=nil })] = 0
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
    locked=data.locked, hidden=data.hidden,
  }
  reg.xf_map[key] = id
  return id
end

local function register_dxf(reg, data)
  local key = style_key(data)
  local known = reg.dxf_map[key]
  if known ~= nil then return known end
  local id = #reg.dxfs
  reg.dxfs[#reg.dxfs + 1] = copy_style_data(data)
  reg.dxf_map[key] = id
  return id
end

local function effective_style_data(style, value)
  local source = style and STYLE_DATA[style] or nil
  local data = source and copy_style_data(source)
    or { bold=false, italic=false, strike=false, wrap_text=false, locked=nil, hidden=nil }
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
    local has_protection = xf.locked ~= nil or xf.hidden ~= nil
    if has_alignment then attrs[#attrs + 1] = 'applyAlignment="1"' end
    if has_protection then attrs[#attrs + 1] = 'applyProtection="1"' end
    if has_alignment or has_protection then
      b[#b + 1] = '<xf ' .. table.concat(attrs, " ") .. '>'
      if has_alignment then
        b[#b + 1] = '<alignment'
        if xf.horizontal then b[#b + 1] = ' horizontal="' .. xf.horizontal .. '"' end
        if xf.vertical then b[#b + 1] = ' vertical="' .. xf.vertical .. '"' end
        if xf.wrap_text then b[#b + 1] = ' wrapText="1"' end
        b[#b + 1] = '/>'
      end
      if has_protection then
        b[#b + 1] = '<protection'
        if xf.locked ~= nil then b[#b + 1] = ' locked="' .. (xf.locked and '1' or '0') .. '"' end
        if xf.hidden ~= nil then b[#b + 1] = ' hidden="' .. (xf.hidden and '1' or '0') .. '"' end
        b[#b + 1] = '/>'
      end
      b[#b + 1] = '</xf>'
    else
      b[#b + 1] = '<xf ' .. table.concat(attrs, " ") .. '/>'
    end
  end
  b[#b + 1] = '</cellXfs>'
  b[#b + 1] = '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
  if #reg.dxfs > 0 then
    b[#b + 1] = '<dxfs count="' .. #reg.dxfs .. '">'
    for _, data in ipairs(reg.dxfs) do
      b[#b + 1] = '<dxf>'
      if data.bold or data.italic or data.underline or data.strike or data.font_name or data.font_size or data.font_color then
        b[#b + 1] = '<font>'
        if data.bold then b[#b + 1] = '<b/>' end
        if data.italic then b[#b + 1] = '<i/>' end
        if data.underline == "single" then b[#b + 1] = '<u/>' end
        if data.underline == "double" then b[#b + 1] = '<u val="double"/>' end
        if data.strike then b[#b + 1] = '<strike/>' end
        if data.font_color then b[#b + 1] = '<color rgb="' .. data.font_color .. '"/>' end
        if data.font_size then b[#b + 1] = '<sz val="' .. num2str(data.font_size) .. '"/>' end
        if data.font_name then b[#b + 1] = '<name val="' .. esc_xml(data.font_name, "nom de police", 0) .. '"/>' end
        b[#b + 1] = '</font>'
      end
      if data.fill_color then
        b[#b + 1] = '<fill><patternFill patternType="solid"><fgColor rgb="' .. data.fill_color ..
          '"/><bgColor indexed="64"/></patternFill></fill>'
      end
      if data.border then
        b[#b + 1] = '<border>'
        for _, side in ipairs({ "left", "right", "top", "bottom" }) do
          local item = data.border[side]
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
      if data.number_format then
        local id = register_numfmt(reg, data.number_format)
        b[#b + 1] = '<numFmt numFmtId="' .. id .. '" formatCode="' ..
          esc_xml(data.number_format, "format conditionnel", 0) .. '"/>'
      end
      if data.horizontal or data.vertical or data.wrap_text then
        b[#b + 1] = '<alignment'
        if data.horizontal then b[#b + 1] = ' horizontal="' .. data.horizontal .. '"' end
        if data.vertical then b[#b + 1] = ' vertical="' .. data.vertical .. '"' end
        if data.wrap_text then b[#b + 1] = ' wrapText="1"' end
        b[#b + 1] = '/>'
      end
      b[#b + 1] = '</dxf>'
    end
    b[#b + 1] = '</dxfs>'
  end
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
    row_hidden = {},
    column_hidden = {},
    merged_ranges = {},
    data_validations = {},
    conditional_formats = {},
    comments = {},
    images = {},
    charts = {},
    sparklines = {},
    tables = {},
    sheet_protection = nil,
    row_page_breaks = {},
    column_page_breaks = {},
    page_margins = nil,
    page_setup = nil,
    print_options = nil,
    header_footer = nil,
    _print_area = nil,
    _repeat_rows = nil,
    _repeat_cols = nil,
    maxrow = -1,
    maxcol = -1,
    _workbook = workbook,
    _freeze_rows = 0,
    _freeze_cols = 0,
    _auto_filter = false,
    _tab_color = nil,
    _visibility = "visible",
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
  if mt == DATE_MT or mt == FORMULA_MT or mt == RICH_TEXT_MT then
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

--- Écrit un texte enrichi dans une cellule.
function Sheet:write_rich_text(row, col, runs, style)
  return self:write(row, col, xlsx.rich_text(runs), style)
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


local validate_optional_text

-- ----------------------------------------------------------------------------
-- Images, graphiques, tableaux et mise en page d'impression
-- ----------------------------------------------------------------------------
local function read_binary_file(path, level)
  if type(path) ~= "string" or path == "" then error("xlsx: le chemin d'image doit être une string non vide", level or 3) end
  local f, err = io.open(path, "rb")
  if not f then error("xlsx: impossible d'ouvrir l'image : " .. tostring(err), level or 3) end
  local data, rerr = f:read("a")
  local ok, cerr = f:close()
  if data == nil then error("xlsx: impossible de lire l'image : " .. tostring(rerr), level or 3) end
  if ok == nil then error("xlsx: impossible de fermer l'image : " .. tostring(cerr), level or 3) end
  return data
end

local function image_info(data, requested, level)
  if type(data) ~= "string" or #data == 0 then error("xlsx: les données d'image doivent être une string binaire non vide", level or 3) end
  local format, width, height
  if data:sub(1,8) == "\137PNG\r\n\26\n" and #data >= 24 then
    format = "png"
    width = string.unpack(">I4", data, 17)
    height = string.unpack(">I4", data, 21)
  elseif data:sub(1,2) == "\255\216" then
    format = "jpeg"
    local pos = 3
    while pos + 8 <= #data do
      if data:byte(pos) ~= 0xFF then pos = pos + 1 else
        while pos <= #data and data:byte(pos) == 0xFF do pos = pos + 1 end
        local marker_byte = data:byte(pos); pos = pos + 1
        if not marker_byte then break end
        if marker_byte ~= 0xD8 and marker_byte ~= 0xD9 and marker_byte ~= 0x01 and not (marker_byte >= 0xD0 and marker_byte <= 0xD7) then
          if pos + 1 > #data then break end
          local length = string.unpack(">I2", data, pos)
          if length < 2 or pos + length - 1 > #data then break end
          if (marker_byte >= 0xC0 and marker_byte <= 0xC3) or (marker_byte >= 0xC5 and marker_byte <= 0xC7)
              or (marker_byte >= 0xC9 and marker_byte <= 0xCB) or (marker_byte >= 0xCD and marker_byte <= 0xCF) then
            height = string.unpack(">I2", data, pos + 3)
            width = string.unpack(">I2", data, pos + 5)
            break
          end
          pos = pos + length
        end
      end
    end
  end
  if requested ~= nil then
    if type(requested) ~= "string" then error("xlsx: format d'image doit être une string", level or 3) end
    requested = requested:lower():gsub("^jpg$", "jpeg")
    if requested ~= "png" and requested ~= "jpeg" then error("xlsx: format d'image doit valoir png, jpeg ou jpg", level or 3) end
    if format and requested ~= format then error("xlsx: le format annoncé ne correspond pas aux données de l'image", level or 3) end
    format = requested
  end
  if not format then error("xlsx: seules les images PNG et JPEG sont prises en charge", level or 3) end
  if not width or not height or width < 1 or height < 1 then error("xlsx: dimensions d'image invalides ou introuvables", level or 3) end
  return format, width, height
end

local function normalize_image_opts(opts, natural_width, natural_height, level)
  opts = opts or {}
  if type(opts) ~= "table" then error("xlsx: les options d'image doivent être une table", level or 3) end
  local allowed = { width=true, height=true, alt_text=true, name=true }
  for k in pairs(opts) do if not allowed[k] then error("xlsx: option d'image inconnue : " .. tostring(k), level or 3) end end
  local width, height = opts.width, opts.height
  if width ~= nil then check_finite_number(width, "largeur d'image", 1, 10000, level or 3) end
  if height ~= nil then check_finite_number(height, "hauteur d'image", 1, 10000, level or 3) end
  if width and not height then height = natural_height * width / natural_width end
  if height and not width then width = natural_width * height / natural_height end
  width, height = width or natural_width, height or natural_height
  local alt_text = validate_optional_text(opts.alt_text, "texte alternatif", 1000, level or 3)
  local name = validate_optional_text(opts.name, "nom d'image", 255, level or 3)
  return { width=width, height=height, alt_text=alt_text, name=name }
end

--- Ajoute une image PNG/JPEG lue depuis un fichier, ancrée à une cellule 0-indexée.
function Sheet:add_image(path, row, col, opts)
  check_index(row, "row", 3); check_index(col, "col", 3)
  local data = read_binary_file(path, 3)
  local format, natural_width, natural_height = image_info(data, nil, 3)
  local normalized = normalize_image_opts(opts, natural_width, natural_height, 3)
  normalized.data, normalized.format, normalized.row, normalized.col = data, format, row, col
  self.images[#self.images + 1] = normalized
  return self
end

--- Ajoute directement une image binaire PNG/JPEG.
function Sheet:add_image_data(data, format, row, col, opts)
  check_index(row, "row", 3); check_index(col, "col", 3)
  local detected, natural_width, natural_height = image_info(data, format, 3)
  local normalized = normalize_image_opts(opts, natural_width, natural_height, 3)
  normalized.data, normalized.format, normalized.row, normalized.col = data, detected, row, col
  self.images[#self.images + 1] = normalized
  return self
end

function Sheet:remove_image(index)
  if math.type(index) ~= "integer" or index < 1 or not self.images[index] then error("xlsx: index d'image invalide", 2) end
  table.remove(self.images, index); return self
end

local CHART_TYPES = { line=true, column=true, bar=true, pie=true, doughnut=true, area=true, scatter=true }
local CHART_GROUPINGS = { standard="standard", stacked="stacked", percent_stacked="percentStacked" }
local LEGEND_POSITIONS = { right="r", left="l", top="t", bottom="b", top_right="tr" }
local LEGEND_POSITIONS_FROM_XML = {}; for k,v in pairs(LEGEND_POSITIONS) do LEGEND_POSITIONS_FROM_XML[v]=k end
local MARKERS = { none=true, circle=true, dash=true, diamond=true, dot=true, picture=true, plus=true, square=true, star=true, triangle=true, x=true }
local function absolute_a1_range(ref, what, level)
  local r1, c1, r2, c2 = parse_a1_range(ref, what, true, level or 3)
  return "$" .. col_ref(c1) .. "$" .. (r1 + 1) .. ":$" .. col_ref(c2) .. "$" .. (r2 + 1)
end
local function quote_sheet_name(name)
  return "'" .. name:gsub("'", "''") .. "'"
end
local function chart_formula(sheet, ref, what, level)
  return quote_sheet_name(sheet.name) .. "!" .. absolute_a1_range(ref, what, level or 3)
end
local function normalize_axis(opts, what, level)
  if opts==nil then return {} end
  if type(opts)~="table" then error("xlsx: "..what.." doit être une table",level or 3) end
  local allowed={title=true,min=true,max=true,number_format=true,major_gridlines=true}
  for k in pairs(opts) do if not allowed[k] then error("xlsx: option d'axe inconnue : "..tostring(k),level or 3) end end
  local out={ title=validate_optional_text(opts.title,what..".title",1000,level or 3) }
  for _,key in ipairs({"min","max"}) do if opts[key]~=nil then check_finite_number(opts[key],what.."."..key,-1e307,1e307,level or 3); out[key]=opts[key] end end
  if out.min and out.max and out.min>=out.max then error("xlsx: le minimum d'axe doit être inférieur au maximum",level or 3) end
  if opts.number_format~=nil then out.number_format=validate_optional_text(opts.number_format,what..".number_format",255,level or 3) end
  if opts.major_gridlines~=nil and type(opts.major_gridlines)~="boolean" then error("xlsx: major_gridlines doit être un booléen",level or 3) end
  out.major_gridlines=opts.major_gridlines~=false
  return out
end
local function normalize_data_labels(opts, level)
  if opts==nil or opts==false then return nil end
  if opts==true then return {show_value=true} end
  if type(opts)~="table" then error("xlsx: data_labels doit être un booléen ou une table",level or 3) end
  local allowed={show_value=true,show_percent=true,show_category=true,show_series_name=true,position=true}
  for k in pairs(opts) do if not allowed[k] then error("xlsx: option d'étiquette inconnue : "..tostring(k),level or 3) end end
  local out={}
  for _,key in ipairs({"show_value","show_percent","show_category","show_series_name"}) do if opts[key]~=nil and type(opts[key])~="boolean" then error("xlsx: "..key.." doit être un booléen",level or 3) end; out[key]=opts[key]==true end
  local positions={center="ctr",inside_end="inEnd",inside_base="inBase",outside_end="outEnd",best_fit="bestFit",left="l",right="r",top="t",bottom="b"}
  if opts.position~=nil then if not positions[opts.position] then error("xlsx: position d'étiquette inconnue",level or 3) end; out.position=positions[opts.position] end
  return out
end
local function normalize_chart(opts, sheet, level)
  if type(opts) ~= "table" then error("xlsx: add_chart attend une table d'options", level or 3) end
  local allowed = { type=true, title=true, categories=true, x_values=true, series=true, row=true, col=true, width=true, height=true,
    legend=true, legend_position=true, grouping=true, x_axis=true, y_axis=true, hole_size=true, data_labels=true }
  for k in pairs(opts) do if not allowed[k] then error("xlsx: option de graphique inconnue : " .. tostring(k), level or 3) end end
  local kind = opts.type or "column"
  if not CHART_TYPES[kind] then error("xlsx: type de graphique inconnu", level or 3) end
  local row, col = opts.row or 0, opts.col or 0
  check_index(row, "row", level or 3); check_index(col, "col", level or 3)
  local width, height = opts.width or 640, opts.height or 360
  check_finite_number(width, "largeur de graphique", 100, 2000, level or 3)
  check_finite_number(height, "hauteur de graphique", 100, 2000, level or 3)
  local title = validate_optional_text(opts.title, "titre de graphique", 1000, level or 3)
  local categories = kind~="scatter" and chart_formula(sheet, opts.categories, "catégories du graphique", level or 3) or nil
  local default_x = kind=="scatter" and chart_formula(sheet, opts.x_values, "valeurs X", level or 3) or nil
  local n = dense_array(opts.series, "séries du graphique", level or 3)
  if n < 1 or n > 255 then error("xlsx: un graphique doit contenir entre 1 et 255 séries", level or 3) end
  if (kind=="pie" or kind=="doughnut") and n~=1 then error("xlsx: un graphique pie/doughnut exige exactement une série",level or 3) end
  local series = {}
  for i, src in ipairs(opts.series) do
    if type(src) ~= "table" then error("xlsx: chaque série doit être une table", level or 3) end
    local allowed_series={name=true,name_ref=true,values=true,x_values=true,color=true,marker=true,line_width=true,smooth=true}
    for k in pairs(src) do if not allowed_series[k] then error("xlsx: option de série inconnue : " .. tostring(k), level or 3) end end
    if src.name ~= nil and src.name_ref ~= nil then error("xlsx: une série ne peut pas avoir name et name_ref", level or 3) end
    local item = { values=chart_formula(sheet, src.values, "valeurs de série", level or 3) }
    if kind=="scatter" then item.x_values=src.x_values and chart_formula(sheet,src.x_values,"valeurs X de série",level or 3) or default_x; if not item.x_values then error("xlsx: scatter exige x_values",level or 3) end end
    if src.name ~= nil then item.name = validate_optional_text(src.name, "nom de série", 255, level or 3) end
    if src.name_ref ~= nil then item.name_ref = chart_formula(sheet, src.name_ref, "nom de série", level or 3) end
    item.color=normalize_color(src.color,"couleur de série",level or 3)
    if src.marker~=nil then if type(src.marker)~="string" or not MARKERS[src.marker] then error("xlsx: marqueur de série inconnu",level or 3) end; item.marker=src.marker end
    if src.line_width~=nil then check_finite_number(src.line_width,"épaisseur de ligne",0.25,20,level or 3); item.line_width=src.line_width end
    if src.smooth~=nil and type(src.smooth)~="boolean" then error("xlsx: smooth doit être un booléen",level or 3) end; item.smooth=src.smooth==true
    series[i] = item
  end
  if opts.legend ~= nil and type(opts.legend) ~= "boolean" then error("xlsx: legend doit être un booléen", level or 3) end
  local legend_position=opts.legend_position or "right"; if not LEGEND_POSITIONS[legend_position] then error("xlsx: position de légende inconnue",level or 3) end
  local grouping=opts.grouping or "standard"; if not CHART_GROUPINGS[grouping] then error("xlsx: grouping doit valoir standard, stacked ou percent_stacked",level or 3) end
  if (kind=="pie" or kind=="doughnut" or kind=="scatter") and opts.grouping~=nil then error("xlsx: grouping n'est pas applicable à ce type de graphique",level or 3) end
  local hole_size=opts.hole_size or 50; if kind=="doughnut" then check_finite_number(hole_size,"taille du trou",10,90,level or 3) elseif opts.hole_size~=nil then error("xlsx: hole_size est réservé aux graphiques doughnut",level or 3) end
  return { type=kind, title=title, categories=categories, series=series, row=row, col=col, width=width, height=height,
    legend=opts.legend ~= false, legend_position=legend_position, grouping=grouping, x_axis=normalize_axis(opts.x_axis,"x_axis",level),
    y_axis=normalize_axis(opts.y_axis,"y_axis",level), hole_size=hole_size, data_labels=normalize_data_labels(opts.data_labels,level) }
end
--- Ajoute un graphique line, column ou bar.
function Sheet:add_chart(opts)
  self.charts[#self.charts + 1] = normalize_chart(opts, self, 3)
  return self
end
function Sheet:remove_chart(index)
  if math.type(index) ~= "integer" or index < 1 or not self.charts[index] then error("xlsx: index de graphique invalide", 2) end
  table.remove(self.charts, index); return self
end

local function validate_table_name(name, level)
  if type(name) ~= "string" or name == "" then error("xlsx: le nom de tableau doit être une string non vide", level or 3) end
  validate_xml_text(name, "nom de tableau", level or 3)
  if #name > 255 or not name:match("^[A-Za-z_][A-Za-z0-9_.]*$") or name:match("^[A-Za-z]+%d+$") then
    error("xlsx: nom de tableau invalide", level or 3)
  end
  return name
end
local function normalize_table(sheet, ref, opts, level)
  local r1, c1, r2, c2, normalized = parse_a1_range(ref, "plage de tableau", false, level or 3)
  if r2 <= r1 then error("xlsx: un tableau doit contenir un en-tête et au moins une ligne de données", level or 3) end
  opts = opts or {}; if type(opts) ~= "table" then error("xlsx: les options de tableau doivent être une table", level or 3) end
  local allowed = { name=true, style=true, show_first_column=true, show_last_column=true, show_row_stripes=true, show_column_stripes=true }
  for k in pairs(opts) do if not allowed[k] then error("xlsx: option de tableau inconnue : " .. tostring(k), level or 3) end end
  local name = validate_table_name(opts.name or ("Table" .. (#sheet.tables + 1)), level or 3)
  local style = opts.style or "TableStyleMedium2"
  local valid_style = type(style) == "string" and (style:match("^TableStyleLight%d+$") or style:match("^TableStyleMedium%d+$") or style:match("^TableStyleDark%d+$"))
  if not valid_style then error("xlsx: style de tableau invalide", level or 3) end
  local columns, seen = {}, {}
  local header = sheet.cells[r1]
  for col = c1, c2 do
    local value = header and header[col]
    if type(value) ~= "string" or value == "" then error("xlsx: chaque en-tête de tableau doit être une string non vide", level or 3) end
    local key = value:lower(); if seen[key] then error("xlsx: en-tête de tableau dupliqué : " .. value, level or 3) end
    seen[key] = true; columns[#columns + 1] = value
  end
  for _, key in ipairs({"show_first_column","show_last_column","show_row_stripes","show_column_stripes"}) do
    if opts[key] ~= nil and type(opts[key]) ~= "boolean" then error("xlsx: " .. key .. " doit être un booléen", level or 3) end
  end
  return { ref=normalized, r1=r1,c1=c1,r2=r2,c2=c2,name=name,style=style,columns=columns,
    show_first_column=opts.show_first_column == true, show_last_column=opts.show_last_column == true,
    show_row_stripes=opts.show_row_stripes ~= false, show_column_stripes=opts.show_column_stripes == true }
end

--- Ajoute un tableau structuré Excel sur une plage possédant une ligne d'en-tête.
function Sheet:add_table(ref, opts)
  local item = normalize_table(self, ref, opts, 3)
  for _, current in ipairs(self.tables) do if ranges_overlap(item, current) then error("xlsx: le tableau " .. item.ref .. " chevauche " .. current.ref, 2) end end
  self.tables[#self.tables + 1] = item; return self
end
function Sheet:remove_table(name)
  name = validate_table_name(name, 3)
  for i, item in ipairs(self.tables) do if item.name:lower() == name:lower() then table.remove(self.tables, i); return self end end
  return self
end

local function legacy_password_hash(password, level)
  if password == nil then return nil end
  if type(password) ~= "string" then error("xlsx: password doit être une string", level or 3) end
  validate_xml_text(password, "mot de passe", level or 3)
  if #password > 15 then error("xlsx: la protection classique accepte au maximum 15 octets", level or 3) end
  local hash = 0
  for i = 1, #password do
    local value = password:byte(i) << i
    local rotated = value >> 15
    value = value & 0x7fff
    hash = hash ~ (value | rotated)
  end
  hash = hash ~ #password ~ 0xCE4B
  return string.format("%X", hash & 0xFFFF)
end

local SHEET_PROTECTION_OPTIONS = {
  select_locked_cells=true, select_unlocked_cells=true, format_cells=true,
  format_columns=true, format_rows=true, insert_columns=true, insert_rows=true,
  insert_hyperlinks=true, delete_columns=true, delete_rows=true, sort=true,
  auto_filter=true, pivot_tables=true, objects=true, scenarios=true,
}

--- Protège ou déprotège une feuille. La protection XLSX n'est pas un chiffrement.
function Sheet:protect(opts)
  if opts == false then self.sheet_protection = nil; return self end
  opts = opts or {}; if type(opts) ~= "table" then error("xlsx: protect attend une table", 2) end
  for k in pairs(opts) do if k ~= "password" and not SHEET_PROTECTION_OPTIONS[k] then error("xlsx: option de protection inconnue : " .. tostring(k), 2) end end
  local out = { password_hash=legacy_password_hash(opts.password, 3) }
  for key in pairs(SHEET_PROTECTION_OPTIONS) do
    if opts[key] ~= nil and type(opts[key]) ~= "boolean" then error("xlsx: " .. key .. " doit être un booléen", 2) end
    out[key] = opts[key] == true
  end
  self.sheet_protection = out
  return self
end

function Sheet:get_protection()
  return self.sheet_protection and shallow_copy(self.sheet_protection) or nil
end

local function add_page_break(list, value, max, what, level)
  if math.type(value) ~= "integer" or value < 1 or value > max then error("xlsx: " .. what .. " doit être un entier entre 1 et " .. max, level or 3) end
  for _, current in ipairs(list) do if current == value then return end end
  list[#list + 1] = value; table.sort(list)
end
function Sheet:add_row_page_break(row) add_page_break(self.row_page_breaks, row, MAX_ROW, "saut de ligne", 3); return self end
function Sheet:add_column_page_break(col)
  if type(col) == "string" then col = col_number(col); if col ~= nil then col = col + 1 end end
  add_page_break(self.column_page_breaks, col, MAX_COL, "saut de colonne", 3); return self
end
function Sheet:clear_page_breaks() self.row_page_breaks = {}; self.column_page_breaks = {}; return self end
function Sheet:get_row_page_breaks() return { table.unpack(self.row_page_breaks) } end
function Sheet:get_column_page_breaks() return { table.unpack(self.column_page_breaks) } end

local SPARKLINE_TYPES = { line="line", column="column", win_loss="stacked" }
local function normalize_sparkline(sheet, target, source, opts, level)
  local _,_,_,_,target_ref=parse_a1_range(target,"cellule de sparkline",true,level or 3)
  if target_ref:match(":") and target_ref:match("^([^:]+):%1$") then target_ref=target_ref:match("^([^:]+)") end
  local tr1,tc1,tr2,tc2=parse_a1_range(target,"cellule de sparkline",true,level or 3)
  if tr1~=tr2 or tc1~=tc2 then error("xlsx: une sparkline exige une seule cellule cible",level or 3) end
  local _,_,_,_,source_ref=parse_a1_range(source,"source de sparkline",false,level or 3)
  opts=opts or {}; if type(opts)~="table" then error("xlsx: options de sparkline invalides",level or 3) end
  local allowed={type=true,color=true,negative_color=true,high_color=true,low_color=true,first_color=true,last_color=true,
    show_markers=true,show_high=true,show_low=true,show_first=true,show_last=true,show_negative=true,right_to_left=true}
  for k in pairs(opts) do if not allowed[k] then error("xlsx: option de sparkline inconnue : "..tostring(k),level or 3) end end
  local kind=opts.type or "line"; if not SPARKLINE_TYPES[kind] then error("xlsx: type de sparkline doit valoir line, column ou win_loss",level or 3) end
  local out={target=col_ref(tc1)..(tr1+1), source=quote_sheet_name(sheet.name).."!"..absolute_a1_range(source_ref,"source de sparkline",level or 3), type=kind}
  for _,key in ipairs({"color","negative_color","high_color","low_color","first_color","last_color"}) do out[key]=normalize_color(opts[key],key,level or 3) end
  for _,key in ipairs({"show_markers","show_high","show_low","show_first","show_last","show_negative","right_to_left"}) do
    if opts[key]~=nil and type(opts[key])~="boolean" then error("xlsx: "..key.." doit être un booléen",level or 3) end
    out[key]=opts[key]==true
  end
  return out
end
function Sheet:add_sparkline(target, source, opts) self.sparklines[#self.sparklines+1]=normalize_sparkline(self,target,source,opts,3); return self end
function Sheet:remove_sparkline(target)
  local r,c=parse_a1_range(target,"cellule de sparkline",true,3); local ref=col_ref(c)..(r+1)
  for i=#self.sparklines,1,-1 do if self.sparklines[i].target==ref then table.remove(self.sparklines,i) end end
  return self
end
function Sheet:get_sparklines()
  local out={}; for i,v in ipairs(self.sparklines) do out[i]=shallow_copy(v) end; return out
end

local PAPER_SIZES = { letter=1, legal=5, a4=9, a3=8, a5=11 }
--- Configure orientation, papier, mise à l'échelle et centrage à l'impression.
function Sheet:set_page_setup(opts)
  if opts == false then self.page_setup = nil; self.print_options = nil; return self end
  opts = opts or {}; if type(opts) ~= "table" then error("xlsx: set_page_setup attend une table", 2) end
  local allowed = { orientation=true, paper_size=true, scale=true, fit_to_width=true, fit_to_height=true, horizontal_centered=true, vertical_centered=true, grid_lines=true, headings=true }
  for k in pairs(opts) do if not allowed[k] then error("xlsx: option de mise en page inconnue : " .. tostring(k), 2) end end
  local orientation = opts.orientation or "portrait"
  if orientation ~= "portrait" and orientation ~= "landscape" then error("xlsx: orientation doit valoir portrait ou landscape", 2) end
  local paper = opts.paper_size or "a4"
  local paper_id = type(paper) == "string" and PAPER_SIZES[paper:lower()] or paper
  if math.type(paper_id) ~= "integer" or paper_id < 1 or paper_id > 118 then error("xlsx: paper_size invalide", 2) end
  local scale = opts.scale
  if scale ~= nil then check_finite_number(scale, "échelle d'impression", 10, 400, 3) end
  local fitw, fith = opts.fit_to_width, opts.fit_to_height
  if fitw ~= nil and (math.type(fitw) ~= "integer" or fitw < 0 or fitw > 32767) then error("xlsx: fit_to_width doit être un entier 0..32767", 2) end
  if fith ~= nil and (math.type(fith) ~= "integer" or fith < 0 or fith > 32767) then error("xlsx: fit_to_height doit être un entier 0..32767", 2) end
  if scale ~= nil and (fitw ~= nil or fith ~= nil) then error("xlsx: scale est incompatible avec fit_to_width/fit_to_height", 2) end
  for _, key in ipairs({"horizontal_centered","vertical_centered","grid_lines","headings"}) do if opts[key] ~= nil and type(opts[key]) ~= "boolean" then error("xlsx: " .. key .. " doit être un booléen", 2) end end
  self.page_setup = { orientation=orientation,paper_size=paper_id,scale=scale,fit_to_width=fitw,fit_to_height=fith }
  self.print_options = { horizontal_centered=opts.horizontal_centered==true, vertical_centered=opts.vertical_centered==true,
    grid_lines=opts.grid_lines==true, headings=opts.headings==true }
  return self
end

function Sheet:set_page_margins(opts)
  if opts == false then self.page_margins = nil; return self end
  opts = opts or {}; if type(opts) ~= "table" then error("xlsx: set_page_margins attend une table", 2) end
  local defaults = { left=0.7,right=0.7,top=0.75,bottom=0.75,header=0.3,footer=0.3 }
  for k in pairs(opts) do if defaults[k] == nil then error("xlsx: marge inconnue : " .. tostring(k), 2) end end
  local out = {}; for k,v in pairs(defaults) do out[k]=opts[k] == nil and v or opts[k]; check_finite_number(out[k], "marge "..k, 0, 100, 3) end
  self.page_margins = out; return self
end

function Sheet:set_header_footer(opts)
  if opts == false then self.header_footer = nil; return self end
  opts = opts or {}; if type(opts) ~= "table" then error("xlsx: set_header_footer attend une table", 2) end
  local allowed = { header_left=true,header_center=true,header_right=true,footer_left=true,footer_center=true,footer_right=true,different_first=true,different_odd_even=true }
  for k in pairs(opts) do if not allowed[k] then error("xlsx: option d'en-tête/pied inconnue : " .. tostring(k), 2) end end
  local out = {}
  for _, key in ipairs({"header_left","header_center","header_right","footer_left","footer_center","footer_right"}) do out[key]=validate_optional_text(opts[key], key, 255, 3) end
  for _, key in ipairs({"different_first","different_odd_even"}) do if opts[key] ~= nil and type(opts[key]) ~= "boolean" then error("xlsx: "..key.." doit être un booléen",2) end; out[key]=opts[key]==true end
  self.header_footer=out; return self
end

function Sheet:set_print_area(ref)
  if ref == false then self._print_area=nil; return self end
  local _,_,_,_,normalized=parse_a1_range(ref,"zone d'impression",false,3); self._print_area=normalized; return self
end
local function normalize_repeat_rows(value, level)
  if type(value) ~= "string" then error("xlsx: repeat_rows doit être une string comme 1:2", level or 3) end
  local a,b=value:match("^(%d+):(%d+)$"); a,b=tonumber(a),tonumber(b)
  if not a or a<1 or b<a or b>MAX_ROW+1 then error("xlsx: repeat_rows invalide", level or 3) end
  return "$"..a..":$"..b
end
local function normalize_repeat_cols(value, level)
  if type(value) ~= "string" then error("xlsx: repeat_columns doit être une string comme A:B", level or 3) end
  local a,b=value:match("^%$?([A-Za-z]+):%$?([A-Za-z]+)$"); local c1,c2=col_number(a or ""),col_number(b or "")
  if not c1 or not c2 or c1>c2 or c2>MAX_COL then error("xlsx: repeat_columns invalide", level or 3) end
  return "$"..col_ref(c1)..":$"..col_ref(c2)
end
function Sheet:set_print_titles(opts)
  if opts == false then self._repeat_rows=nil; self._repeat_cols=nil; return self end
  opts=opts or {}; if type(opts)~="table" then error("xlsx: set_print_titles attend une table",2) end
  for k in pairs(opts) do if k~="rows" and k~="columns" then error("xlsx: option de titres d'impression inconnue : "..tostring(k),2) end end
  self._repeat_rows=opts.rows and normalize_repeat_rows(opts.rows,3) or nil
  self._repeat_cols=opts.columns and normalize_repeat_cols(opts.columns,3) or nil
  if not self._repeat_rows and not self._repeat_cols then error("xlsx: set_print_titles exige rows et/ou columns",2) end
  return self
end

local DATA_VALIDATION_TYPES = {
  list="list", whole="whole", decimal="decimal", date="date", time="time",
  text_length="textLength", custom="custom",
}
local VALIDATION_OPERATORS = {
  between="between", not_between="notBetween", equal="equal", not_equal="notEqual",
  greater_than="greaterThan", less_than="lessThan",
  greater_or_equal="greaterThanOrEqual", less_or_equal="lessThanOrEqual",
}
local VALIDATION_OPERATOR_FROM_XML = {}
for k, v in pairs(VALIDATION_OPERATORS) do VALIDATION_OPERATOR_FROM_XML[v] = k end

validate_optional_text = function(value, what, max_bytes, level)
  if value == nil then return nil end
  if type(value) ~= "string" then error("xlsx: " .. what .. " doit être une string", level or 3) end
  validate_xml_text(value, what, level or 3)
  if #value > max_bytes then error("xlsx: " .. what .. " dépasse " .. max_bytes .. " octets", level or 3) end
  return value
end

local function validation_operand(value, what, level)
  if value == nil then return nil end
  if getmetatable(value) == DATE_MT then return value end
  if type(value) == "number" then
    if value ~= value or value == math.huge or value == -math.huge then
      error("xlsx: " .. what .. " doit être fini", level or 3)
    end
    return value
  end
  if type(value) ~= "string" or value == "" then
    error("xlsx: " .. what .. " doit être un nombre, une date ou une formule string non vide", level or 3)
  end
  validate_xml_text(value, what, level or 3)
  return value:sub(1, 1) == "=" and value:sub(2) or value
end

local function normalize_data_validation(ref, opts, level)
  local _, _, _, _, normalized = parse_a1_range(ref, "plage de validation", true, level or 3)
  if type(opts) ~= "table" then error("xlsx: les options de validation doivent être une table", level or 3) end
  local allowed = {
    type=true, values=true, formula=true, operator=true, minimum=true, maximum=true, value=true,
    allow_blank=true, show_input_message=true, show_error_message=true, show_dropdown=true,
    prompt_title=true, prompt=true, error_title=true, error=true, error_style=true,
  }
  for k in pairs(opts) do if not allowed[k] then error("xlsx: option de validation inconnue : " .. tostring(k), level or 3) end end
  local kind = opts.type
  if type(kind) ~= "string" or not DATA_VALIDATION_TYPES[kind] then
    error("xlsx: type de validation inconnu", level or 3)
  end
  local out = { ref=normalized, type=kind }
  for _, key in ipairs({"allow_blank", "show_input_message", "show_error_message", "show_dropdown"}) do
    if opts[key] ~= nil and type(opts[key]) ~= "boolean" then error("xlsx: " .. key .. " doit être un booléen", level or 3) end
    out[key] = opts[key]
  end
  out.allow_blank = out.allow_blank == true
  out.show_input_message = out.show_input_message == true
  out.show_error_message = out.show_error_message == true
  out.show_dropdown = out.show_dropdown ~= false
  out.prompt_title = validate_optional_text(opts.prompt_title, "titre d'aide", 32, level)
  out.prompt = validate_optional_text(opts.prompt, "message d'aide", 255, level)
  out.error_title = validate_optional_text(opts.error_title, "titre d'erreur", 32, level)
  out.error = validate_optional_text(opts.error, "message d'erreur", 255, level)
  if opts.error_style ~= nil then
    if opts.error_style ~= "stop" and opts.error_style ~= "warning" and opts.error_style ~= "information" then
      error("xlsx: error_style doit valoir stop, warning ou information", level or 3)
    end
    out.error_style = opts.error_style
  end
  if kind == "list" then
    if opts.values ~= nil and opts.formula ~= nil then error("xlsx: list accepte values ou formula, pas les deux", level or 3) end
    if opts.values ~= nil then
      local n = dense_array(opts.values, "values", level)
      if n == 0 then error("xlsx: values ne peut pas être vide", level or 3) end
      local values = {}
      for i = 1, n do
        local item = opts.values[i]
        if type(item) ~= "string" or item == "" then error("xlsx: chaque choix doit être une string non vide", level or 3) end
        validate_xml_text(item, "choix de validation", level or 3)
        if item:find('[,"]') then error("xlsx: un choix inline ne peut contenir ni virgule ni guillemet", level or 3) end
        values[i] = item
      end
      local joined = table.concat(values, ",")
      if #joined + 2 > 255 then error("xlsx: la liste inline dépasse 255 octets", level or 3) end
      out.values, out.formula1 = values, '"' .. joined .. '"'
    else
      out.formula1 = validation_operand(opts.formula, "formule de liste", level)
      if not out.formula1 then error("xlsx: une validation list exige values ou formula", level or 3) end
    end
  elseif kind == "custom" then
    out.formula1 = validation_operand(opts.formula, "formule personnalisée", level)
    if type(out.formula1) ~= "string" then error("xlsx: custom exige formula string", level or 3) end
  else
    local op = opts.operator or ((opts.minimum ~= nil or opts.maximum ~= nil) and "between" or "equal")
    if not VALIDATION_OPERATORS[op] then error("xlsx: opérateur de validation inconnu", level or 3) end
    out.operator = op
    if op == "between" or op == "not_between" then
      out.formula1 = validation_operand(opts.minimum, "minimum", level)
      out.formula2 = validation_operand(opts.maximum, "maximum", level)
      if out.formula1 == nil or out.formula2 == nil then error("xlsx: between exige minimum et maximum", level or 3) end
    else
      out.formula1 = validation_operand(opts.value, "valeur", level)
      if out.formula1 == nil then error("xlsx: cet opérateur exige value", level or 3) end
    end
  end
  return out
end

function Sheet:add_data_validation(ref, opts)
  self.data_validations[#self.data_validations + 1] = normalize_data_validation(ref, opts, 3)
  return self
end

function Sheet:remove_data_validation(ref)
  local _, _, _, _, normalized = parse_a1_range(ref, "plage de validation", true, 3)
  for i = #self.data_validations, 1, -1 do
    if self.data_validations[i].ref == normalized then table.remove(self.data_validations, i) end
  end
  return self
end

local function copy_validation(v)
  local out = shallow_copy(v)
  if v.values then out.values = { table.unpack(v.values) } end
  return out
end

function Sheet:get_data_validations()
  local out = {}
  for i, v in ipairs(self.data_validations) do out[i] = copy_validation(v) end
  return out
end

local CF_OPERATORS = VALIDATION_OPERATORS
local CF_TYPE_TO_XML = {
  cell="cellIs", contains_text="containsText", blanks="containsBlanks",
  not_blanks="notContainsBlanks", duplicate="duplicateValues", custom="expression",
  color_scale="colorScale", data_bar="dataBar", icon_set="iconSet",
  top="top10", bottom="top10", above_average="aboveAverage", below_average="aboveAverage",
}
local CF_TYPE_FROM_XML = {}
for k, v in pairs(CF_TYPE_TO_XML) do CF_TYPE_FROM_XML[v] = k end

local function top_left_ref(normalized)
  return normalized:match("^([^:]+)")
end

local function excel_quote(text)
  return '"' .. text:gsub('"', '""') .. '"'
end

local CFVO_TYPES = { min=true, max=true, number=true, percent=true, percentile=true, formula=true }
local CFVO_TO_XML = { min="min", max="max", number="num", percent="percent", percentile="percentile", formula="formula" }
local CFVO_FROM_XML = { min="min", max="max", num="number", percent="percent", percentile="percentile", formula="formula" }
local ICON_SETS = {
  ["3_arrows"]="3Arrows", ["3_arrows_gray"]="3ArrowsGray", ["3_flags"]="3Flags",
  ["3_traffic_lights"]="3TrafficLights1", ["3_traffic_lights_rimmed"]="3TrafficLights2",
  ["3_signs"]="3Signs", ["3_symbols"]="3Symbols", ["3_symbols_circled"]="3Symbols2",
  ["4_arrows"]="4Arrows", ["4_arrows_gray"]="4ArrowsGray", ["4_ratings"]="4Rating",
  ["4_traffic_lights"]="4TrafficLights", ["5_arrows"]="5Arrows", ["5_arrows_gray"]="5ArrowsGray",
  ["5_ratings"]="5Rating", ["5_quarters"]="5Quarters",
}
local ICON_SETS_FROM_XML = {}; for k,v in pairs(ICON_SETS) do ICON_SETS_FROM_XML[v]=k end
local function normalize_cfvo(kind, value, what, level)
  kind=kind or what
  if not CFVO_TYPES[kind] then error("xlsx: type de seuil conditionnel inconnu : "..tostring(kind),level or 3) end
  if kind=="min" or kind=="max" then
    if value~=nil then error("xlsx: un seuil min/max ne doit pas avoir de valeur",level or 3) end
    return {type=kind}
  end
  if value==nil then error("xlsx: le seuil "..what.." exige une valeur",level or 3) end
  if kind=="formula" then value=validation_operand(value,"formule de seuil",level); if type(value)~="string" then error("xlsx: une formule de seuil doit être une string",level or 3) end
  elseif type(value)~="number" or value~=value or value==math.huge or value==-math.huge then error("xlsx: valeur de seuil invalide",level or 3) end
  return {type=kind,value=value}
end

local function normalize_conditional_format(ref, opts, level)
  local _, _, _, _, normalized = parse_a1_range(ref, "plage conditionnelle", true, level or 3)
  if type(opts) ~= "table" then error("xlsx: les options conditionnelles doivent être une table", level or 3) end
  local allowed = { type=true, operator=true, value=true, minimum=true, maximum=true, formula=true,
    text=true, style=true, stop_if_true=true,
    min_type=true,min_value=true,min_color=true,mid_type=true,mid_value=true,mid_color=true,max_type=true,max_value=true,max_color=true,
    start_type=true,start_value=true,end_type=true,end_value=true,color=true,show_value=true,
    icons=true,value_type=true,values=true,reverse=true,rank=true,percent=true,
  }
  for k in pairs(opts) do if not allowed[k] then error("xlsx: option conditionnelle inconnue : " .. tostring(k), level or 3) end end
  local kind = opts.type
  if type(kind) ~= "string" or not CF_TYPE_TO_XML[kind] then error("xlsx: type conditionnel inconnu", level or 3) end
  if opts.stop_if_true ~= nil and type(opts.stop_if_true) ~= "boolean" then error("xlsx: stop_if_true doit être un booléen", level or 3) end
  local out = { ref=normalized, type=kind, stop_if_true=opts.stop_if_true == true }
  local first = top_left_ref(normalized)
  if kind=="color_scale" then
    if opts.style~=nil then error("xlsx: color_scale n'accepte pas style",level or 3) end
    out.min=normalize_cfvo(opts.min_type or "min",opts.min_value,"minimum",level)
    out.max=normalize_cfvo(opts.max_type or "max",opts.max_value,"maximum",level)
    out.min_color=normalize_color(opts.min_color or "F8696B","min_color",level)
    out.max_color=normalize_color(opts.max_color or "63BE7B","max_color",level)
    if opts.mid_type~=nil or opts.mid_value~=nil or opts.mid_color~=nil then
      out.mid=normalize_cfvo(opts.mid_type or "percentile",opts.mid_value==nil and 50 or opts.mid_value,"milieu",level)
      out.mid_color=normalize_color(opts.mid_color or "FFEB84","mid_color",level)
    end
    return out
  elseif kind=="data_bar" then
    if opts.style~=nil then error("xlsx: data_bar n'accepte pas style",level or 3) end
    out.start=normalize_cfvo(opts.start_type or "min",opts.start_value,"début",level)
    out.finish=normalize_cfvo(opts.end_type or "max",opts.end_value,"fin",level)
    out.color=normalize_color(opts.color or "5B9BD5","couleur de barre",level)
    if opts.show_value~=nil and type(opts.show_value)~="boolean" then error("xlsx: show_value doit être un booléen",level or 3) end
    out.show_value=opts.show_value~=false
    return out
  elseif kind=="icon_set" then
    if opts.style~=nil then error("xlsx: icon_set n'accepte pas style",level or 3) end
    local icon=opts.icons or "3_traffic_lights"; if not ICON_SETS[icon] then error("xlsx: jeu d'icônes inconnu",level or 3) end
    local count=tonumber(icon:match("^(%d)")); local values=opts.values or {}; dense_array(values,"seuils d'icônes",level)
    if #values~=count then error("xlsx: le jeu d'icônes exige "..count.." seuils",level or 3) end
    local value_type=opts.value_type or "percent"; if not CFVO_TYPES[value_type] or value_type=="min" or value_type=="max" then error("xlsx: value_type d'icônes invalide",level or 3) end
    out.icons=icon; out.thresholds={}; for i,v in ipairs(values) do out.thresholds[i]=normalize_cfvo(value_type,v,"icône",level) end
    for _,key in ipairs({"show_value","reverse"}) do if opts[key]~=nil and type(opts[key])~="boolean" then error("xlsx: "..key.." doit être un booléen",level or 3) end end
    out.show_value=opts.show_value~=false; out.reverse=opts.reverse==true
    return out
  elseif kind=="top" or kind=="bottom" then
    if opts.style==nil then error("xlsx: top/bottom exige style",level or 3) end
    out.rank=opts.rank or 10; if math.type(out.rank)~="integer" or out.rank<1 or out.rank>1000 then error("xlsx: rank doit être un entier 1..1000",level or 3) end
    if opts.percent~=nil and type(opts.percent)~="boolean" then error("xlsx: percent doit être un booléen",level or 3) end
    out.percent=opts.percent==true; out.bottom=kind=="bottom"
  elseif kind=="above_average" or kind=="below_average" then
    if opts.style==nil then error("xlsx: above_average/below_average exige style",level or 3) end
    out.above_average=kind=="above_average"
  end
  if kind~="color_scale" and kind~="data_bar" and kind~="icon_set" then
    local style_data = require_style(opts.style, level or 3)
    if not style_data then error("xlsx: un format conditionnel exige style", level or 3) end
    if style_data.number_format or style_data.horizontal or style_data.vertical or style_data.wrap_text or style_data.locked~=nil or style_data.hidden~=nil then
      error("xlsx: les formats conditionnels acceptent police, fond et bordures uniquement", level or 3)
    end
    out.style=opts.style
  end
  if kind == "cell" then
    local op = opts.operator
    if not CF_OPERATORS[op] then error("xlsx: opérateur conditionnel inconnu", level or 3) end
    out.operator = op
    if op == "between" or op == "not_between" then
      out.formula1 = validation_operand(opts.minimum, "minimum conditionnel", level)
      out.formula2 = validation_operand(opts.maximum, "maximum conditionnel", level)
      if out.formula1 == nil or out.formula2 == nil then error("xlsx: between exige minimum et maximum", level or 3) end
    else
      out.formula1 = validation_operand(opts.value, "valeur conditionnelle", level)
      if out.formula1 == nil then error("xlsx: l'opérateur conditionnel exige value", level or 3) end
    end
  elseif kind == "contains_text" then
    local text = validate_optional_text(opts.text, "texte conditionnel", 255, level)
    if not text or text == "" then error("xlsx: contains_text exige text", level or 3) end
    out.text = text
    out.formula1 = 'NOT(ISERROR(SEARCH(' .. excel_quote(text) .. ',' .. first .. ')))'
  elseif kind == "blanks" then out.formula1 = 'LEN(TRIM(' .. first .. '))=0'
  elseif kind == "not_blanks" then out.formula1 = 'LEN(TRIM(' .. first .. '))>0'
  elseif kind == "custom" then
    out.formula1 = validation_operand(opts.formula, "formule conditionnelle", level)
    if type(out.formula1) ~= "string" then error("xlsx: custom exige formula string", level or 3) end
  end
  return out
end
function Sheet:add_conditional_format(ref, opts)
  self.conditional_formats[#self.conditional_formats + 1] = normalize_conditional_format(ref, opts, 3)
  return self
end

function Sheet:remove_conditional_format(ref)
  local _, _, _, _, normalized = parse_a1_range(ref, "plage conditionnelle", true, 3)
  for i = #self.conditional_formats, 1, -1 do
    if self.conditional_formats[i].ref == normalized then table.remove(self.conditional_formats, i) end
  end
  return self
end

local function copy_conditional_format(v)
  local out=shallow_copy(v)
  if v.min then out.min=shallow_copy(v.min) end; if v.mid then out.mid=shallow_copy(v.mid) end; if v.max then out.max=shallow_copy(v.max) end
  if v.start then out.start=shallow_copy(v.start) end; if v.finish then out.finish=shallow_copy(v.finish) end
  if v.thresholds then out.thresholds={}; for i,x in ipairs(v.thresholds) do out.thresholds[i]=shallow_copy(x) end end
  return out
end

function Sheet:get_conditional_formats()
  local out = {}
  for i, v in ipairs(self.conditional_formats) do out[i] = copy_conditional_format(v) end
  return out
end

function Sheet:set_comment(row, col, comment)
  check_index(row, "row", 3); check_index(col, "col", 3)
  if type(comment) ~= "table" then error("xlsx: comment doit être une table {author,text}", 2) end
  for k in pairs(comment) do if k ~= "author" and k ~= "text" then error("xlsx: option de commentaire inconnue : " .. tostring(k), 2) end end
  local author = validate_optional_text(comment.author, "auteur du commentaire", 255, 3)
  local text = validate_optional_text(comment.text, "texte du commentaire", 32767, 3)
  if not author or author == "" then error("xlsx: author doit être non vide", 2) end
  if not text or text == "" then error("xlsx: text doit être non vide", 2) end
  local rr = self.comments[row]; if not rr then rr = {}; self.comments[row] = rr end
  rr[col] = { author=author, text=text }
  ensure_cell_extent(self, row, col)
  return self
end

function Sheet:remove_comment(row, col)
  check_index(row, "row", 3); check_index(col, "col", 3)
  local rr = self.comments[row]; if rr then rr[col] = nil end
  return self
end

function Sheet:set_row_hidden(row, hidden)
  check_index(row, "row", 3)
  if type(hidden) ~= "boolean" then error("xlsx: hidden doit être un booléen", 2) end
  self.row_hidden[row] = hidden or nil
  if hidden then ensure_cell_extent(self, row, 0) end
  return self
end

function Sheet:set_column_hidden(col, hidden)
  check_index(col, "col", 3)
  if type(hidden) ~= "boolean" then error("xlsx: hidden doit être un booléen", 2) end
  self.column_hidden[col] = hidden or nil
  return self
end

function Sheet:set_tab_color(color)
  self._tab_color = normalize_color(color, "couleur d'onglet", 3)
  return self
end

function Sheet:set_visibility(state)
  if state ~= "visible" and state ~= "hidden" and state ~= "very_hidden" then
    error("xlsx: visibilité doit valoir visible, hidden ou very_hidden", 2)
  end
  self._visibility = state
  return self
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
    local attrs = {
      'Id="' .. rel.id .. '"',
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/' .. rel.kind .. '"',
      'Target="' .. esc_xml(rel.target, "cible de relation", 0) .. '"',
    }
    if rel.external then attrs[#attrs + 1] = 'TargetMode="External"' end
    b[#b + 1] = '<Relationship ' .. table.concat(attrs, " ") .. '/>'
  end
  b[#b + 1] = '</Relationships>'
  return table.concat(b)
end

local function validation_formula_xml(value, date1904)
  if value == nil then return nil end
  if getmetatable(value) == DATE_MT then
    return num2str(ymd_to_serial(value.y, value.m, value.d, value.h, value.mi, value.s, date1904))
  end
  if type(value) == "number" then return num2str(value) end
  return value
end

local function comments_xml(sheet)
  local entries, author_ids, authors = {}, {}, {}
  for row, rr in pairs(sheet.comments) do
    for col, comment in pairs(rr) do entries[#entries + 1] = { row=row, col=col, data=comment } end
  end
  if #entries == 0 then return nil, nil end
  table.sort(entries, function(a, b) return a.row ~= b.row and a.row < b.row or a.col < b.col end)
  for _, entry in ipairs(entries) do
    local author = entry.data.author
    if author_ids[author] == nil then author_ids[author] = #authors; authors[#authors + 1] = author end
  end
  local b = { '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>' }
  for _, author in ipairs(authors) do b[#b + 1] = '<author>' .. esc_xml(author, "auteur", 0) .. '</author>' end
  b[#b + 1] = '</authors><commentList>'
  for i, entry in ipairs(entries) do
    local ref = col_ref(entry.col) .. (entry.row + 1)
    b[#b + 1] = '<comment ref="' .. ref .. '" authorId="' .. author_ids[entry.data.author] ..
      '" shapeId="' .. (i - 1) .. '"><text><t xml:space="preserve">' ..
      esc_xml(entry.data.text, "commentaire", 0) .. '</t></text></comment>'
  end
  b[#b + 1] = '</commentList></comments>'

  local v = { '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">',
    '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>',
    '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' }
  for i, entry in ipairs(entries) do
    v[#v + 1] = '<v:shape type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:144px;height:79px;z-index:' ..
      i .. ';visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto" id="_x0000_s' .. (1025 + i) .. '">' ..
      '<v:fill color2="#ffffe1"/><v:shadow color="black" obscured="t"/><v:path o:connecttype="none"/>' ..
      '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>' ..
      '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill><x:Row>' ..
      entry.row .. '</x:Row><x:Column>' .. entry.col .. '</x:Column></x:ClientData></v:shape>'
  end
  v[#v + 1] = '</xml>'
  return table.concat(b), table.concat(v)
end

-- sérialise la feuille en XML ; alimente sst et renvoie aussi ses relations.
function Sheet:_xml(sst, sst_index)
  local b = {}
  b[#b + 1] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  b[#b + 1] = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
  if self._tab_color or (self.page_setup and (self.page_setup.fit_to_width ~= nil or self.page_setup.fit_to_height ~= nil)) then
    b[#b + 1] = '<sheetPr>'
    if self._tab_color then b[#b + 1] = '<tabColor rgb="' .. self._tab_color .. '"/>' end
    if self.page_setup and (self.page_setup.fit_to_width ~= nil or self.page_setup.fit_to_height ~= nil) then b[#b + 1] = '<pageSetUpPr fitToPage="1"/>' end
    b[#b + 1] = '</sheetPr>'
  end
  if self.maxrow >= 0 and self.maxcol >= 0 then
    b[#b + 1] = '<dimension ref="A1:' .. col_ref(self.maxcol) .. (self.maxrow + 1) .. '"/>'
  end
  local pane = freeze_pane_xml(self._freeze_rows, self._freeze_cols)
  if pane then b[#b + 1] = pane end
  if next(self.column_widths) or next(self.column_hidden) then
    b[#b + 1] = '<cols>'
    local cols, seen_cols = {}, {}
    for col in pairs(self.column_widths) do cols[#cols + 1] = col; seen_cols[col] = true end
    for col in pairs(self.column_hidden) do if not seen_cols[col] then cols[#cols + 1] = col end end
    table.sort(cols)
    for _, col in ipairs(cols) do
      local attrs = { 'min="' .. (col + 1) .. '"', 'max="' .. (col + 1) .. '"' }
      if self.column_widths[col] then
        attrs[#attrs + 1] = 'width="' .. num2str(self.column_widths[col]) .. '"'
        attrs[#attrs + 1] = 'customWidth="1"'
      end
      if self.column_hidden[col] then attrs[#attrs + 1] = 'hidden="1"' end
      b[#b + 1] = '<col ' .. table.concat(attrs, " ") .. '/>'
    end
    b[#b + 1] = '</cols>'
  end
  b[#b + 1] = '<sheetData>'
  for r = 0, self.maxrow do
    local rowcells, rowstyles = self.cells[r], self.styles[r]
    if rowcells or rowstyles or self.row_heights[r] or self.row_hidden[r] then
      local rowattrs = { 'r="' .. (r + 1) .. '"' }
      if self.row_heights[r] then
        rowattrs[#rowattrs + 1] = 'ht="' .. num2str(self.row_heights[r]) .. '"'
        rowattrs[#rowattrs + 1] = 'customHeight="1"'
      end
      if self.row_hidden[r] then rowattrs[#rowattrs + 1] = 'hidden="1"' end
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
          elseif mt == RICH_TEXT_MT then
            local idx = sst_index[v]
            if idx == nil then idx = #sst; sst[#sst + 1] = v; sst_index[v] = idx end
            b[#b + 1] = '<c r="' .. ref .. '"' .. sattr .. ' t="s"><v>' .. idx .. '</v></c>'
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
  if self.sheet_protection then
    local p=self.sheet_protection; local attrs={'sheet="1"'}
    if p.password_hash then attrs[#attrs+1]='password="'..p.password_hash..'"' end
    local names={select_locked_cells="selectLockedCells",select_unlocked_cells="selectUnlockedCells",format_cells="formatCells",format_columns="formatColumns",format_rows="formatRows",insert_columns="insertColumns",insert_rows="insertRows",insert_hyperlinks="insertHyperlinks",delete_columns="deleteColumns",delete_rows="deleteRows",sort="sort",auto_filter="autoFilter",pivot_tables="pivotTables",objects="objects",scenarios="scenarios"}
    for key,attr in pairs(names) do attrs[#attrs+1]=attr..'="'..(p[key] and '1' or '0')..'"' end
    b[#b+1]='<sheetProtection '..table.concat(attrs,' ')..'/>'
  end
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
  local function cfvo_xml(item)
    local attrs={'type="'..CFVO_TO_XML[item.type]..'"'}; if item.value~=nil then attrs[#attrs+1]='val="'..esc_xml(num2str(item.value),"seuil",0)..'"' end
    return '<cfvo '..table.concat(attrs,' ')..'/>'
  end
  for priority, rule in ipairs(self.conditional_formats) do
    local xml_type = CF_TYPE_TO_XML[rule.type]
    local attrs = { 'type="' .. xml_type .. '"', 'priority="' .. priority .. '"' }
    if rule.style then attrs[#attrs+1]='dxfId="'..self._workbook:_dxf_id(rule.style)..'"' end
    if rule.stop_if_true then attrs[#attrs + 1] = 'stopIfTrue="1"' end
    if rule.operator then attrs[#attrs + 1] = 'operator="' .. CF_OPERATORS[rule.operator] .. '"' end
    if rule.text then attrs[#attrs + 1] = 'text="' .. esc_xml(rule.text, "texte conditionnel", 0) .. '"' end
    if rule.type=="top" or rule.type=="bottom" then attrs[#attrs+1]='rank="'..rule.rank..'"'; if rule.percent then attrs[#attrs+1]='percent="1"' end; if rule.bottom then attrs[#attrs+1]='bottom="1"' end end
    if rule.type=="above_average" or rule.type=="below_average" then attrs[#attrs+1]='aboveAverage="'..(rule.above_average and '1' or '0')..'"' end
    b[#b + 1] = '<conditionalFormatting sqref="' .. rule.ref .. '"><cfRule ' .. table.concat(attrs, " ") .. '>'
    if rule.type=="color_scale" then
      b[#b+1]='<colorScale>'..cfvo_xml(rule.min)..(rule.mid and cfvo_xml(rule.mid) or '')..cfvo_xml(rule.max)
      b[#b+1]='<color rgb="'..rule.min_color..'"/>'..(rule.mid_color and '<color rgb="'..rule.mid_color..'"/>' or '')..'<color rgb="'..rule.max_color..'"/></colorScale>'
    elseif rule.type=="data_bar" then
      b[#b+1]='<dataBar showValue="'..(rule.show_value and '1' or '0')..'">'..cfvo_xml(rule.start)..cfvo_xml(rule.finish)..'<color rgb="'..rule.color..'"/></dataBar>'
    elseif rule.type=="icon_set" then
      b[#b+1]='<iconSet iconSet="'..ICON_SETS[rule.icons]..'" showValue="'..(rule.show_value and '1' or '0')..'" reverse="'..(rule.reverse and '1' or '0')..'">'
      for _,threshold in ipairs(rule.thresholds) do b[#b+1]=cfvo_xml(threshold) end; b[#b+1]='</iconSet>'
    else
      local f1 = validation_formula_xml(rule.formula1, self._workbook._date1904)
      local f2 = validation_formula_xml(rule.formula2, self._workbook._date1904)
      if f1 then b[#b + 1] = '<formula>' .. esc_xml(f1, "formule conditionnelle", 0) .. '</formula>' end
      if f2 then b[#b + 1] = '<formula>' .. esc_xml(f2, "formule conditionnelle", 0) .. '</formula>' end
    end
    b[#b + 1] = '</cfRule></conditionalFormatting>'
  end
  if #self.data_validations > 0 then
    b[#b + 1] = '<dataValidations count="' .. #self.data_validations .. '">'
    for _, validation in ipairs(self.data_validations) do
      local attrs = { 'sqref="' .. validation.ref .. '"', 'type="' .. DATA_VALIDATION_TYPES[validation.type] .. '"' }
      if validation.operator then attrs[#attrs + 1] = 'operator="' .. VALIDATION_OPERATORS[validation.operator] .. '"' end
      attrs[#attrs + 1] = 'allowBlank="' .. (validation.allow_blank and '1' or '0') .. '"'
      attrs[#attrs + 1] = 'showInputMessage="' .. (validation.show_input_message and '1' or '0') .. '"'
      attrs[#attrs + 1] = 'showErrorMessage="' .. (validation.show_error_message and '1' or '0') .. '"'
      attrs[#attrs + 1] = 'showDropDown="' .. (validation.show_dropdown and '0' or '1') .. '"'
      for _, key in ipairs({ "prompt_title", "prompt", "error_title", "error", "error_style" }) do
        local attr = ({prompt_title="promptTitle", error_title="errorTitle", error_style="errorStyle"})[key] or key
        if validation[key] then attrs[#attrs + 1] = attr .. '="' .. esc_xml(validation[key], attr, 0) .. '"' end
      end
      b[#b + 1] = '<dataValidation ' .. table.concat(attrs, " ") .. '>'
      local f1 = validation_formula_xml(validation.formula1, self._workbook._date1904)
      local f2 = validation_formula_xml(validation.formula2, self._workbook._date1904)
      if f1 then b[#b + 1] = '<formula1>' .. esc_xml(f1, "formule de validation", 0) .. '</formula1>' end
      if f2 then b[#b + 1] = '<formula2>' .. esc_xml(f2, "formule de validation", 0) .. '</formula2>' end
      b[#b + 1] = '</dataValidation>'
    end
    b[#b + 1] = '</dataValidations>'
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
        rels[#rels + 1] = { id=id, target=link.target, kind="hyperlink", external=true }
        attrs[#attrs + 1] = 'r:id="' .. id .. '"'
      end
      links[#links + 1] = '<hyperlink ' .. table.concat(attrs, " ") .. '/>'
    end
  end
  if #links > 0 then b[#b + 1] = '<hyperlinks>' .. table.concat(links) .. '</hyperlinks>' end
  if self.print_options then
    local attrs = {}
    if self.print_options.horizontal_centered then attrs[#attrs+1]='horizontalCentered="1"' end
    if self.print_options.vertical_centered then attrs[#attrs+1]='verticalCentered="1"' end
    if self.print_options.grid_lines then attrs[#attrs+1]='gridLines="1"' end
    if self.print_options.headings then attrs[#attrs+1]='headings="1"' end
    if #attrs > 0 then b[#b+1]='<printOptions '..table.concat(attrs,' ')..'/>' end
  end
  if self.page_margins then
    local m=self.page_margins; b[#b+1]=string.format('<pageMargins left="%s" right="%s" top="%s" bottom="%s" header="%s" footer="%s"/>',num2str(m.left),num2str(m.right),num2str(m.top),num2str(m.bottom),num2str(m.header),num2str(m.footer))
  end
  if self.page_setup then
    local p=self.page_setup; local attrs={'orientation="'..p.orientation..'"','paperSize="'..p.paper_size..'"'}
    if p.scale then attrs[#attrs+1]='scale="'..num2str(p.scale)..'"' end
    if p.fit_to_width ~= nil then attrs[#attrs+1]='fitToWidth="'..p.fit_to_width..'"' end
    if p.fit_to_height ~= nil then attrs[#attrs+1]='fitToHeight="'..p.fit_to_height..'"' end
    b[#b+1]='<pageSetup '..table.concat(attrs,' ')..'/>'
  end
  if self.header_footer then
    local h=self.header_footer; local attrs={}
    if h.different_first then attrs[#attrs+1]='differentFirst="1"' end
    if h.different_odd_even then attrs[#attrs+1]='differentOddEven="1"' end
    b[#b+1]='<headerFooter'..(#attrs>0 and (' '..table.concat(attrs,' ')) or '')..'>'
    local header=(h.header_left and '&L'..h.header_left or '')..(h.header_center and '&C'..h.header_center or '')..(h.header_right and '&R'..h.header_right or '')
    local footer=(h.footer_left and '&L'..h.footer_left or '')..(h.footer_center and '&C'..h.footer_center or '')..(h.footer_right and '&R'..h.footer_right or '')
    if header~='' then b[#b+1]='<oddHeader>'..esc_xml(header,"en-tête",0)..'</oddHeader>' end
    if footer~='' then b[#b+1]='<oddFooter>'..esc_xml(footer,"pied de page",0)..'</oddFooter>' end
    b[#b+1]='</headerFooter>'
  end

  if #self.row_page_breaks>0 then
    b[#b+1]='<rowBreaks count="'..#self.row_page_breaks..'" manualBreakCount="'..#self.row_page_breaks..'">'
    for _,id in ipairs(self.row_page_breaks) do b[#b+1]='<brk id="'..id..'" min="0" max="'..MAX_COL..'" man="1"/>' end
    b[#b+1]='</rowBreaks>'
  end
  if #self.column_page_breaks>0 then
    b[#b+1]='<colBreaks count="'..#self.column_page_breaks..'" manualBreakCount="'..#self.column_page_breaks..'">'
    for _,id in ipairs(self.column_page_breaks) do b[#b+1]='<brk id="'..id..'" min="0" max="'..MAX_ROW..'" man="1"/>' end
    b[#b+1]='</colBreaks>'
  end

  local comments, vml = comments_xml(self)
  if comments then
    local comments_id = "rId" .. (#rels + 1)
    rels[#rels + 1] = { id=comments_id, target="../comments/comment" .. self._sheet_index .. ".xml", kind="comments" }
    local vml_id = "rId" .. (#rels + 1)
    rels[#rels + 1] = { id=vml_id, target="../drawings/commentsDrawing" .. self._sheet_index .. ".vml", kind="vmlDrawing" }
    b[#b + 1] = '<legacyDrawing r:id="' .. vml_id .. '"/>'
  end
  if #self.images + #self.charts > 0 then
    local drawing_id="rId"..(#rels+1); rels[#rels+1]={id=drawing_id,target="../drawings/drawing"..self._sheet_index..".xml",kind="drawing"}
    b[#b+1]='<drawing r:id="'..drawing_id..'"/>'
  end
  if #self.tables > 0 then
    b[#b+1]='<tableParts count="'..#self.tables..'">'
    for _, table_def in ipairs(self.tables) do
      local id="rId"..(#rels+1); rels[#rels+1]={id=id,target="../tables/table"..table_def._table_id..".xml",kind="table"}
      b[#b+1]='<tablePart r:id="'..id..'"/>'
    end
    b[#b+1]='</tableParts>'
  end
  if #self.sparklines>0 then
    b[#b+1]='<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
    for _,sp in ipairs(self.sparklines) do
      local attrs={'type="'..SPARKLINE_TYPES[sp.type]..'"'}
      local flags={show_markers="markers",show_high="high",show_low="low",show_first="first",show_last="last",show_negative="negative",right_to_left="rightToLeft"}
      for key,attr in pairs(flags) do if sp[key] then attrs[#attrs+1]=attr..'="1"' end end
      b[#b+1]='<x14:sparklineGroup '..table.concat(attrs,' ')..'>'
      local colors={color="colorSeries",negative_color="colorNegative",high_color="colorHigh",low_color="colorLow",first_color="colorFirst",last_color="colorLast"}
      for key,tag in pairs(colors) do if sp[key] then b[#b+1]='<x14:'..tag..' rgb="'..sp[key]..'"/>' end end
      b[#b+1]='<x14:sparklines><x14:sparkline><xm:f>'..esc_xml(sp.source,"source de sparkline",0)..'</xm:f><xm:sqref>'..sp.target..'</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>'
    end
    b[#b+1]='</x14:sparklineGroups></ext></extLst>'
  end
  b[#b + 1] = '</worksheet>'
  return table.concat(b), worksheet_relationships_xml(rels), comments, vml
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

function Workbook:_dxf_id(style)
  local data = require_style(style, 3)
  if not self._style_registry then error("xlsx: registre de styles indisponible", 2) end
  return register_dxf(self._style_registry, data)
end

local function validate_defined_name(name, level)
  if type(name) ~= "string" or name == "" then error("xlsx: le nom défini doit être une string non vide", level or 3) end
  validate_xml_text(name, "nom défini", level or 3)
  if #name > 255 then error("xlsx: le nom défini dépasse 255 octets", level or 3) end
  if not name:match("^[A-Za-z_\\][A-Za-z0-9_.\\]*$") then error("xlsx: nom défini invalide", level or 3) end
  if name:match("^[Rr][Cc]$") or name:match("^[Rr]%d+$") or name:match("^[Cc]%d+$") or name:match("^[A-Za-z]+%d+$") then
    error("xlsx: le nom défini ressemble à une référence de cellule", level or 3)
  end
  return name
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
    _style_registry = nil, _active_sheet = 1, _defined_names = {}, _defined_name_map = {},
    _workbook_protection = nil, _properties = {},
  }, Workbook)
end

--- Définit les propriétés principales et étendues du document.
function Workbook:set_properties(opts)
  if opts == false then self._properties = {}; return self end
  opts = opts or {}; if type(opts) ~= "table" then error("xlsx: set_properties attend une table", 2) end
  local allowed = { title=true, subject=true, creator=true, description=true, keywords=true, category=true, company=true, manager=true }
  for k in pairs(opts) do if not allowed[k] then error("xlsx: propriété inconnue : " .. tostring(k), 2) end end
  local out = {}
  for _, key in ipairs({"title","subject","creator","description","category","company","manager"}) do
    out[key] = validate_optional_text(opts[key], "propriété "..key, 4096, 3)
  end
  if opts.keywords ~= nil then
    if type(opts.keywords) == "table" then
      dense_array(opts.keywords,"mots-clés",3); local values={}
      for i,v in ipairs(opts.keywords) do values[i]=validate_optional_text(v,"mot-clé",255,3) end
      out.keywords=table.concat(values, ", ")
    else out.keywords=validate_optional_text(opts.keywords,"mots-clés",4096,3) end
  end
  self._properties=out; return self
end
function Workbook:get_properties() return shallow_copy(self._properties) end

--- Protège la structure du classeur sans chiffrer son contenu.
function Workbook:protect(opts)
  if opts == false then self._workbook_protection=nil; return self end
  opts=opts or {}; if type(opts)~="table" then error("xlsx: protect attend une table",2) end
  local allowed={password=true,structure=true,windows=true}
  for k in pairs(opts) do if not allowed[k] then error("xlsx: option de protection du classeur inconnue : "..tostring(k),2) end end
  for _,key in ipairs({"structure","windows"}) do if opts[key]~=nil and type(opts[key])~="boolean" then error("xlsx: "..key.." doit être un booléen",2) end end
  self._workbook_protection={ password_hash=legacy_password_hash(opts.password,3), structure=opts.structure~=false, windows=opts.windows==true }
  return self
end
function Workbook:get_protection() return self._workbook_protection and shallow_copy(self._workbook_protection) or nil end

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
  sh._sheet_index = #self.sheets + 1
  self.sheets[#self.sheets + 1] = sh
  self._sheet_names[key] = true
  return sh
end

function Workbook:set_active_sheet(which)
  local index
  if math.type(which) == "integer" then index = which
  elseif type(which) == "string" then
    for i, sheet in ipairs(self.sheets) do if sheet.name == which then index = i; break end end
  else error("xlsx: set_active_sheet attend un nom ou un index 1-based", 2) end
  if not index or not self.sheets[index] then error("xlsx: feuille active introuvable", 2) end
  if self.sheets[index]._visibility ~= "visible" then error("xlsx: la feuille active doit être visible", 2) end
  self._active_sheet = index
  return self
end

function Workbook:get_active_sheet()
  return self.sheets[self._active_sheet]
end

local function resolve_local_sheet(workbook, value, level)
  if value == nil then return nil end
  if math.type(value) == "integer" then
    if value < 1 or not workbook.sheets[value] then error("xlsx: local_sheet invalide", level or 3) end
    return value
  end
  if type(value) == "string" then
    for i, sheet in ipairs(workbook.sheets) do if sheet.name == value then return i end end
    error("xlsx: local_sheet introuvable", level or 3)
  end
  error("xlsx: local_sheet doit être un nom ou un index 1-based", level or 3)
end

function Workbook:define_name(name, reference, opts)
  name = validate_defined_name(name, 3)
  if type(reference) ~= "string" or reference == "" then error("xlsx: reference doit être une string non vide", 2) end
  validate_xml_text(reference, "référence de nom défini", 3)
  if #reference > 8192 then error("xlsx: référence de nom défini trop longue", 2) end
  opts = opts or {}; if type(opts) ~= "table" then error("xlsx: opts doit être une table", 2) end
  for k in pairs(opts) do if k ~= "local_sheet" and k ~= "hidden" and k ~= "comment" then error("xlsx: option de nom défini inconnue : " .. tostring(k), 2) end end
  if opts.hidden ~= nil and type(opts.hidden) ~= "boolean" then error("xlsx: hidden doit être un booléen", 2) end
  local local_sheet = resolve_local_sheet(self, opts.local_sheet, 3)
  local comment = validate_optional_text(opts.comment, "commentaire du nom défini", 255, 3)
  local key = name:lower() .. "@" .. tostring(local_sheet or 0)
  if self._defined_name_map[key] then error("xlsx: nom défini dupliqué : " .. name, 2) end
  local item = { name=name, reference=reference, local_sheet=local_sheet, hidden=opts.hidden == true, comment=comment }
  self._defined_names[#self._defined_names + 1] = item; self._defined_name_map[key] = item
  return self
end

function Workbook:remove_defined_name(name, opts)
  name = validate_defined_name(name, 3); opts = opts or {}
  local local_sheet = resolve_local_sheet(self, opts.local_sheet, 3)
  local key = name:lower() .. "@" .. tostring(local_sheet or 0)
  local item = self._defined_name_map[key]
  if not item then return self end
  for i, value in ipairs(self._defined_names) do if value == item then table.remove(self._defined_names, i); break end end
  self._defined_name_map[key] = nil
  return self
end

function Workbook:get_defined_names()
  local out = {}
  for i, item in ipairs(self._defined_names) do out[i] = shallow_copy(item) end
  return out
end


local function table_xml(item)
  local b={'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="'..item._table_id..'" name="'..esc_xml(item.name,"nom de tableau",0)..'" displayName="'..esc_xml(item.name,"nom de tableau",0)..'" ref="'..item.ref..'" totalsRowShown="0">',
    '<autoFilter ref="'..item.ref..'"/>','<tableColumns count="'..#item.columns..'">'}
  for i,name in ipairs(item.columns) do b[#b+1]='<tableColumn id="'..i..'" name="'..esc_xml(name,"colonne de tableau",0)..'"/>' end
  b[#b+1]='</tableColumns><tableStyleInfo name="'..item.style..'" showFirstColumn="'..(item.show_first_column and '1' or '0')..'" showLastColumn="'..(item.show_last_column and '1' or '0')..'" showRowStripes="'..(item.show_row_stripes and '1' or '0')..'" showColumnStripes="'..(item.show_column_stripes and '1' or '0')..'"/></table>'
  return table.concat(b)
end

local function chart_title_xml(title)
  if not title then return '<c:autoTitleDeleted val="1"/>' end
  return '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="fr-FR"/><a:t>'..esc_xml(title,"titre de graphique",0)..'</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title>'
end
local function axis_title_xml(title)
  return title and chart_title_xml(title) or ''
end
local function series_shape_xml(series, line_only)
  if not series.color and not series.line_width then return '' end
  local b={'<c:spPr>'}
  if series.color and not line_only then b[#b+1]='<a:solidFill><a:srgbClr val="'..series.color:sub(-6)..'"/></a:solidFill>' end
  b[#b+1]='<a:ln'..(series.line_width and (' w="'..math.floor(series.line_width*12700+0.5)..'"') or '')..'>'
  if series.color then b[#b+1]='<a:solidFill><a:srgbClr val="'..series.color:sub(-6)..'"/></a:solidFill>' end
  b[#b+1]='<a:prstDash val="solid"/></a:ln></c:spPr>'
  return table.concat(b)
end
local function data_labels_xml(labels)
  if not labels then return '' end
  local b={'<c:dLbls>'}
  if labels.position then b[#b+1]='<c:dLblPos val="'..labels.position..'"/>' end
  b[#b+1]='<c:showLegendKey val="0"/><c:showVal val="'..(labels.show_value and '1' or '0')..'"/><c:showCatName val="'..(labels.show_category and '1' or '0')..'"/><c:showSerName val="'..(labels.show_series_name and '1' or '0')..'"/><c:showPercent val="'..(labels.show_percent and '1' or '0')..'"/>'
  b[#b+1]='</c:dLbls>'; return table.concat(b)
end
local function chart_axis_xml(tag,id,cross,pos,axis,category)
  local b={'<c:'..tag..'><c:axId val="'..id..'"/><c:scaling><c:orientation val="minMax"/>'}
  if axis.min~=nil then b[#b+1]='<c:min val="'..num2str(axis.min)..'"/>' end; if axis.max~=nil then b[#b+1]='<c:max val="'..num2str(axis.max)..'"/>' end
  b[#b+1]='</c:scaling><c:delete val="0"/><c:axPos val="'..pos..'"/>'
  if axis.major_gridlines then b[#b+1]='<c:majorGridlines/>' end
  if axis.title then b[#b+1]=axis_title_xml(axis.title) end
  b[#b+1]='<c:numFmt formatCode="'..esc_xml(axis.number_format or 'General','format axe',0)..'" sourceLinked="'..(axis.number_format and '0' or '1')..'"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="'..cross..'"/><c:crosses val="autoZero"/>'
  if category then b[#b+1]='<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>' else b[#b+1]='<c:crossBetween val="between"/>' end
  b[#b+1]='</c:'..tag..'>'; return table.concat(b)
end
local function chart_xml(chart, chart_id)
  local x_axis=100000+chart_id*2; local y_axis=x_axis+1
  local b={'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>','<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart>',chart_title_xml(chart.title),'<c:plotArea><c:layout/>'}
  local tag=({line='lineChart',column='barChart',bar='barChart',pie='pieChart',doughnut='doughnutChart',area='areaChart',scatter='scatterChart'})[chart.type]
  b[#b+1]='<c:'..tag..'>'
  if chart.type=='line' or chart.type=='area' then b[#b+1]='<c:grouping val="'..CHART_GROUPINGS[chart.grouping]..'"/><c:varyColors val="0"/>'
  elseif chart.type=='column' or chart.type=='bar' then b[#b+1]='<c:barDir val="'..(chart.type=='bar' and 'bar' or 'col')..'"/><c:grouping val="'..CHART_GROUPINGS[chart.grouping]..'"/><c:varyColors val="0"/>'
  elseif chart.type=='pie' or chart.type=='doughnut' then b[#b+1]='<c:varyColors val="1"/>' end
  for i,series in ipairs(chart.series) do
    b[#b+1]='<c:ser><c:idx val="'..(i-1)..'"/><c:order val="'..(i-1)..'"/>'
    if series.name_ref then b[#b+1]='<c:tx><c:strRef><c:f>'..esc_xml(series.name_ref,'référence de série',0)..'</c:f></c:strRef></c:tx>' elseif series.name then b[#b+1]='<c:tx><c:v>'..esc_xml(series.name,'nom de série',0)..'</c:v></c:tx>' end
    local line_only=chart.type=='line' or chart.type=='scatter'; b[#b+1]=series_shape_xml(series,line_only)
    if chart.type=='line' or chart.type=='scatter' then b[#b+1]='<c:marker><c:symbol val="'..(series.marker or 'none')..'"/></c:marker>' end
    if chart.type=='scatter' then
      b[#b+1]='<c:xVal><c:numRef><c:f>'..esc_xml(series.x_values,'valeurs X',0)..'</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>'..esc_xml(series.values,'valeurs Y',0)..'</c:f></c:numRef></c:yVal>'
    else
      b[#b+1]='<c:cat><c:strRef><c:f>'..esc_xml(chart.categories,'catégories',0)..'</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>'..esc_xml(series.values,'valeurs',0)..'</c:f></c:numRef></c:val>'
    end
    if chart.type=='line' or chart.type=='scatter' then b[#b+1]='<c:smooth val="'..(series.smooth and '1' or '0')..'"/>' end
    b[#b+1]='</c:ser>'
  end
  b[#b+1]=data_labels_xml(chart.data_labels)
  if chart.type=='doughnut' then b[#b+1]='<c:firstSliceAng val="0"/><c:holeSize val="'..math.floor(chart.hole_size+0.5)..'"/>'
  elseif chart.type=='pie' then b[#b+1]='<c:firstSliceAng val="0"/>'
  elseif chart.type=='column' or chart.type=='bar' then b[#b+1]='<c:gapWidth val="150"/>'; if chart.grouping~='standard' then b[#b+1]='<c:overlap val="100"/>' end end
  if chart.type~='pie' and chart.type~='doughnut' then b[#b+1]='<c:axId val="'..x_axis..'"/><c:axId val="'..y_axis..'"/>' end
  b[#b+1]='</c:'..tag..'>'
  if chart.type=='scatter' then
    b[#b+1]=chart_axis_xml('valAx',x_axis,y_axis,'b',chart.x_axis,false); b[#b+1]=chart_axis_xml('valAx',y_axis,x_axis,'l',chart.y_axis,false)
  elseif chart.type~='pie' and chart.type~='doughnut' then
    b[#b+1]=chart_axis_xml('catAx',x_axis,y_axis,'b',chart.x_axis,true); b[#b+1]=chart_axis_xml('valAx',y_axis,x_axis,'l',chart.y_axis,false)
  end
  b[#b+1]='</c:plotArea>'
  if chart.legend then b[#b+1]='<c:legend><c:legendPos val="'..LEGEND_POSITIONS[chart.legend_position]..'"/><c:layout/><c:overlay val="0"/></c:legend>' end
  b[#b+1]='<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart></c:chartSpace>'
  return table.concat(b)
end
local function drawing_xml(sheet)
  local b={'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>','<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'}
  local rels={}; local object_id=1
  for _,img in ipairs(sheet.images) do
    local rid='rId'..(#rels+1); rels[#rels+1]={id=rid,kind='image',target='../media/image'..img._media_id..'.'..img.format}
    local name=img.name or ('Image '..object_id); local descr=img.alt_text or ''
    b[#b+1]='<xdr:oneCellAnchor><xdr:from><xdr:col>'..img.col..'</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>'..img.row..'</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="'..math.floor(img.width*9525+0.5)..'" cy="'..math.floor(img.height*9525+0.5)..'"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="'..object_id..'" name="'..esc_xml(name,"nom d'image",0)..'" descr="'..esc_xml(descr,"texte alternatif",0)..'"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="'..rid..'"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="'..math.floor(img.width*9525+0.5)..'" cy="'..math.floor(img.height*9525+0.5)..'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>'
    object_id=object_id+1
  end
  for _,chart in ipairs(sheet.charts) do
    local rid='rId'..(#rels+1); rels[#rels+1]={id=rid,kind='chart',target='../charts/chart'..chart._chart_id..'.xml'}
    b[#b+1]='<xdr:oneCellAnchor><xdr:from><xdr:col>'..chart.col..'</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>'..chart.row..'</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="'..math.floor(chart.width*9525+0.5)..'" cy="'..math.floor(chart.height*9525+0.5)..'"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="'..object_id..'" name="Graphique '..object_id..'"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="'..math.floor(chart.width*9525+0.5)..'" cy="'..math.floor(chart.height*9525+0.5)..'"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="'..rid..'"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>'
    object_id=object_id+1
  end
  b[#b+1]='</xdr:wsDr>'
  return table.concat(b), worksheet_relationships_xml(rels)
end

-- construit la liste des parties ZIP {name=, data=}
function Workbook:_parts()
  local nsheets = #self.sheets
  if nsheets == 0 then error("xlsx: le classeur doit contenir au moins une feuille", 2) end
  local visible = 0
  for _, sheet in ipairs(self.sheets) do if sheet._visibility == "visible" then visible = visible + 1 end end
  if visible == 0 then error("xlsx: au moins une feuille doit rester visible", 2) end
  if not self.sheets[self._active_sheet] or self.sheets[self._active_sheet]._visibility ~= "visible" then
    error("xlsx: la feuille active doit être visible", 2)
  end
  self._style_registry = new_style_registry()
  local table_names, next_table_id, next_media_id, next_chart_id = {}, 1, 1, 1
  for _, sheet in ipairs(self.sheets) do
    for _, item in ipairs(sheet.tables) do
      local key=item.name:lower(); if table_names[key] then error("xlsx: nom de tableau dupliqué dans le classeur : "..item.name,2) end
      table_names[key]=true; item._table_id=next_table_id; next_table_id=next_table_id+1
    end
    for _, item in ipairs(sheet.images) do item._media_id=next_media_id; next_media_id=next_media_id+1 end
    for _, item in ipairs(sheet.charts) do item._chart_id=next_chart_id; next_chart_id=next_chart_id+1 end
  end
  local sst, sst_index = {}, {}
  local sheet_xmls, sheet_rels, sheet_comments, sheet_vml = {}, {}, {}, {}
  for i = 1, nsheets do
    sheet_xmls[i], sheet_rels[i], sheet_comments[i], sheet_vml[i] = self.sheets[i]:_xml(sst, sst_index)
  end

  local parts = {}
  local function add(name, data) parts[#parts + 1] = { name = name, data = data } end

  do
    local t = {}
    t[#t + 1] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    t[#t + 1] = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    t[#t + 1] = '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    t[#t + 1] = '<Default Extension="xml" ContentType="application/xml"/>'
    local has_comments = false
    for i = 1, nsheets do if sheet_comments[i] then has_comments = true; break end end
    if has_comments then t[#t + 1] = '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' end
    if next_media_id > 1 then
      t[#t+1]='<Default Extension="png" ContentType="image/png"/>'
      t[#t+1]='<Default Extension="jpeg" ContentType="image/jpeg"/>'
    end
    t[#t + 1] = '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    t[#t + 1] = '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    t[#t + 1] = '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    for i = 1, nsheets do
      t[#t + 1] = '<Override PartName="/xl/worksheets/sheet' .. i ..
        '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    end
    t[#t + 1] = '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    t[#t + 1] = '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    for i = 1, nsheets do if sheet_comments[i] then
      t[#t + 1] = '<Override PartName="/xl/comments/comment' .. i .. '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>'
    end end
    for i=1,nsheets do if #self.sheets[i].images + #self.sheets[i].charts > 0 then t[#t+1]='<Override PartName="/xl/drawings/drawing'..i..'.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' end end
    for _,sheet in ipairs(self.sheets) do for _,item in ipairs(sheet.charts) do t[#t+1]='<Override PartName="/xl/charts/chart'..item._chart_id..'.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' end end
    for _,sheet in ipairs(self.sheets) do for _,item in ipairs(sheet.tables) do t[#t+1]='<Override PartName="/xl/tables/table'..item._table_id..'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' end end
    t[#t + 1] = '</Types>'
    add("[Content_Types].xml", table.concat(t))
  end

  add("_rels/.rels", table.concat({
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>',
    '</Relationships>',
  }))

  do
    local t = {}
    t[#t + 1] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    t[#t + 1] = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    t[#t + 1] = self._date1904 and '<workbookPr date1904="1"/>' or '<workbookPr/>'
    if self._workbook_protection then
      local p=self._workbook_protection; local attrs={}
      if p.password_hash then attrs[#attrs+1]='workbookPassword="'..p.password_hash..'"' end
      if p.structure then attrs[#attrs+1]='lockStructure="1"' end
      if p.windows then attrs[#attrs+1]='lockWindows="1"' end
      t[#t+1]='<workbookProtection '..table.concat(attrs,' ')..'/>'
    end
    t[#t + 1] = '<bookViews><workbookView activeTab="' .. (self._active_sheet - 1) .. '"/></bookViews>'
    t[#t + 1] = '<sheets>'
    for i = 1, nsheets do
      local state = self.sheets[i]._visibility ~= "visible" and (' state="' .. (self.sheets[i]._visibility == "very_hidden" and "veryHidden" or "hidden") .. '"') or ""
      t[#t + 1] = '<sheet name="' .. esc_xml(self.sheets[i].name, "nom de feuille", 0) ..
        '" sheetId="' .. i .. '"' .. state .. ' r:id="rId' .. i .. '"/>'
    end
    t[#t + 1] = '</sheets>'
    local has_print_names=false; for _,sheet in ipairs(self.sheets) do if sheet._print_area or sheet._repeat_rows or sheet._repeat_cols then has_print_names=true; break end end
    if #self._defined_names > 0 or has_print_names then
      t[#t + 1] = '<definedNames>'
      for _, item in ipairs(self._defined_names) do
        local attrs = { 'name="' .. esc_xml(item.name, "nom défini", 0) .. '"' }
        if item.local_sheet then attrs[#attrs + 1] = 'localSheetId="' .. (item.local_sheet - 1) .. '"' end
        if item.hidden then attrs[#attrs + 1] = 'hidden="1"' end
        if item.comment then attrs[#attrs + 1] = 'comment="' .. esc_xml(item.comment, "commentaire", 0) .. '"' end
        t[#t + 1] = '<definedName ' .. table.concat(attrs, " ") .. '>' .. esc_xml(item.reference, "référence", 0) .. '</definedName>'
      end
      for i,sheet in ipairs(self.sheets) do
        local q=quote_sheet_name(sheet.name)
        if sheet._print_area then t[#t+1]='<definedName name="_xlnm.Print_Area" localSheetId="'..(i-1)..'">'..esc_xml(q..'!'..absolute_a1_range(sheet._print_area,"zone d'impression",0),"zone d'impression",0)..'</definedName>' end
        if sheet._repeat_rows or sheet._repeat_cols then
          local refs={}; if sheet._repeat_rows then refs[#refs+1]=q..'!'..sheet._repeat_rows end; if sheet._repeat_cols then refs[#refs+1]=q..'!'..sheet._repeat_cols end
          t[#t+1]='<definedName name="_xlnm.Print_Titles" localSheetId="'..(i-1)..'">'..esc_xml(table.concat(refs,','),"titres d'impression",0)..'</definedName>'
        end
      end
      t[#t + 1] = '</definedNames>'
    end
    t[#t + 1] = '</workbook>'
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
    if sheet_comments[i] then add("xl/comments/comment" .. i .. ".xml", sheet_comments[i]) end
    if sheet_vml[i] then add("xl/drawings/commentsDrawing" .. i .. ".vml", sheet_vml[i]) end
    local sheet=self.sheets[i]
    if #sheet.images + #sheet.charts > 0 then local dx,dr=drawing_xml(sheet); add("xl/drawings/drawing"..i..".xml",dx); add("xl/drawings/_rels/drawing"..i..".xml.rels",dr) end
    for _,img in ipairs(sheet.images) do add("xl/media/image"..img._media_id.."."..img.format,img.data) end
    for _,chart in ipairs(sheet.charts) do add("xl/charts/chart"..chart._chart_id..".xml",chart_xml(chart,chart._chart_id)) end
    for _,item in ipairs(sheet.tables) do add("xl/tables/table"..item._table_id..".xml",table_xml(item)) end
  end

  do
    local t = {}
    t[#t + 1] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    t[#t + 1] = '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' ..
      #sst .. '" uniqueCount="' .. #sst .. '">'
    for i = 1, #sst do
      local item=sst[i]
      if getmetatable(item)==RICH_TEXT_MT then
        local data=RICH_TEXT_DATA[item]; t[#t+1]='<si>'
        for _,run in ipairs(data.runs) do
          t[#t+1]='<r><rPr>'
          if run.bold then t[#t+1]='<b/>' end; if run.italic then t[#t+1]='<i/>' end
          if run.underline=='single' then t[#t+1]='<u/>' elseif run.underline=='double' then t[#t+1]='<u val="double"/>' end
          if run.strike then t[#t+1]='<strike/>' end
          if run.font_color then t[#t+1]='<color rgb="'..run.font_color..'"/>' end
          if run.font_size then t[#t+1]='<sz val="'..num2str(run.font_size)..'"/>' end
          if run.font_name then t[#t+1]='<rFont val="'..esc_xml(run.font_name,"police enrichie",0)..'"/>' end
          t[#t+1]='</rPr><t xml:space="preserve">'..esc(run.text)..'</t></r>'
        end
        t[#t+1]='</si>'
      else t[#t + 1] = '<si><t xml:space="preserve">' .. esc(item) .. '</t></si>' end
    end
    t[#t + 1] = '</sst>'
    add("xl/sharedStrings.xml", table.concat(t))
  end

  do
    local p=self._properties or {}; local t={'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>','<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">'}
    local tags={creator='dc:creator',title='dc:title',description='dc:description',subject='dc:subject',category='cp:category',keywords='cp:keywords'}
    for _,key in ipairs({'creator','title','description','subject','category','keywords'}) do if p[key] then t[#t+1]='<'..tags[key]..'>'..esc_xml(p[key],"propriété",0)..'</'..tags[key]..'>' end end
    t[#t+1]='</cp:coreProperties>'; add('docProps/core.xml',table.concat(t))
    local a={'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>','<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>lua-xlsx</Application><AppVersion>1.5</AppVersion>'}
    if p.company then a[#a+1]='<Company>'..esc_xml(p.company,"société",0)..'</Company>' end; if p.manager then a[#a+1]='<Manager>'..esc_xml(p.manager,"responsable",0)..'</Manager>' end
    a[#a+1]='</Properties>'; add('docProps/app.xml',table.concat(a))
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
-- filtres, hyperliens, validations, règles conditionnelles, commentaires,
-- propriétés de feuilles et noms définis. read() conserve son contrat historique
-- et renvoie la valeur mise en cache ; get_formula() expose l'expression.

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

local function parse_rich_text_container(xml)
  local buf, runs = {}, {}
  for body in xml:gmatch("<r[^>]*>(.-)</r>") do
    local text = unescape(body:match("<t[^>]*>(.-)</t>") or "")
    local props = body:match("<rPr>(.-)</rPr>") or ""
    local underline = props:match('<u[^>]-val="([^"]+)"')
      or (props:find("<u/>", 1, true) and "single" or nil)
    local color = props:match('<color[^>]-rgb="([%x]+)"')
      or props:match('<color[^>]-val="([%x]+)"')
    runs[#runs + 1] = {
      text = text,
      bold = props:find("<b", 1, true) ~= nil,
      italic = props:find("<i", 1, true) ~= nil,
      underline = underline,
      strike = props:find("<strike", 1, true) ~= nil,
      font_name = unescape(props:match('<rFont[^>]-val="([^"]+)"')
        or props:match('<name[^>]-val="([^"]+)"') or ""),
      font_size = tonumber(props:match('<sz[^>]-val="([^"]+)"')),
      font_color = color and normalize_color(color, "couleur enrichie", 0) or nil,
    }
    if runs[#runs].font_name == "" then runs[#runs].font_name = nil end
    buf[#buf + 1] = text
  end
  if #runs == 0 then
    for text in xml:gmatch("<t[^>]*>(.-)</t>") do buf[#buf + 1] = unescape(text) end
    return table.concat(buf), nil
  end
  local obj = setmetatable({}, RICH_TEXT_MT)
  RICH_TEXT_DATA[obj] = { runs = runs, text = table.concat(buf) }
  return table.concat(buf), obj
end

local function parse_shared_strings(xml)
  local res, rich = {}, {}
  local idx = 0
  xml = xml:gsub("<si%s*/>", "<si></si>")
  for si in xml:gmatch("<si[%s>].-</si>") do
    local text, obj = parse_rich_text_container(si)
    res[idx] = text
    if obj then rich[idx] = obj end
    idx = idx + 1
  end
  return res, rich
end
local function decode_cell(inner, ty, shared, rich_shared)
  if ty == "s" then
    local raw = inner:match("<v>(.-)</v>")
    if not raw then return nil end
    local index = tonumber(raw)
    if not index or math.type(index) ~= "integer" or index < 0 or shared[index] == nil then
      error("xlsx: index sharedStrings invalide : " .. tostring(raw), 0)
    end
    return shared[index], rich_shared and rich_shared[index] or nil
  elseif ty == "b" then
    local raw = inner:match("<v>(.-)</v>")
    if raw == "1" then return true end
    if raw == "0" then return false end
    error("xlsx: booléen de cellule invalide", 0)
  elseif ty == "inlineStr" then
    return parse_rich_text_container(inner)
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

local function parse_dxf_style(body)
  local fontbody = body:match("<font>(.-)</font>") or ""
  local fillbody = body:match("<fill>(.-)</fill>") or ""
  local borderbody = body:match("<border>(.-)</border>") or ""
  local color = fontbody:match('<color[^>]-rgb="([%x]+)"')
  local fill = fillbody:match('<fgColor[^>]-rgb="([%x]+)"')
  local underline = fontbody:match('<u[^>]-val="([^\"]+)"') or (fontbody:find("<u%s*/>") and "single" or nil)
  local border = {}
  for _, side in ipairs({ "left", "right", "top", "bottom" }) do border[side] = parse_border_side(borderbody, side) end
  local align = body:match("<alignment%s*(.-)/>") or body:match("<alignment%s*(.-)>")
  local fmt = body:match('<numFmt[^>]-formatCode="([^\"]*)"')
  local data = {
    bold=fontbody:find("<b[%s/>]", 1) ~= nil, italic=fontbody:find("<i[%s/>]", 1) ~= nil,
    underline=underline, strike=fontbody:find("<strike[%s/>]", 1) ~= nil,
    font_name=unescape(fontbody:match('<name[^>]-val="([^\"]*)"') or ""),
    font_size=tonumber(fontbody:match('<sz[^>]-val="([^\"]*)"')),
    font_color=color and normalize_color(color, "couleur dxf", 0) or nil,
    fill_color=fill and normalize_color(fill, "fond dxf", 0) or nil,
    horizontal=xml_attr(align, "horizontal"), vertical=xml_attr(align, "vertical"),
    wrap_text=xml_attr(align, "wrapText") == "1" or xml_attr(align, "wrapText") == "true",
    number_format=fmt and unescape(fmt) or nil, border=next(border) and border or nil,
  }
  if data.font_name == "" then data.font_name = nil end
  return style_object_from_data(data)
end

local function parse_styles(xml)
  local datestyle, style_objects, dxf_objects = {}, { [0] = false }, {}
  if not xml then return datestyle, style_objects, dxf_objects end
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
    local protection_attrs = body:match("<protection%s*(.-)/>") or body:match("<protection%s*(.-)>")
    local locked, hidden
    if protection_attrs then
      local locked_attr = xml_attr(protection_attrs, "locked")
      local hidden_attr = xml_attr(protection_attrs, "hidden")
      locked = not (locked_attr == "0" or locked_attr == "false")
      hidden = hidden_attr == "1" or hidden_attr == "true"
    end
    local data = {
      bold=font.bold == true, italic=font.italic == true, underline=font.underline,
      strike=font.strike == true, font_name=font.font_name, font_size=font.font_size,
      font_color=font.font_color, fill_color=fill,
      horizontal=xml_attr(alignment_attrs, "horizontal"),
      vertical=xml_attr(alignment_attrs, "vertical"),
      wrap_text=xml_attr(alignment_attrs, "wrapText") == "1" or xml_attr(alignment_attrs, "wrapText") == "true",
      number_format=fmtcode[numFmtId] or (numFmtId ~= 0 and BUILTIN_NUMFMT[numFmtId] or nil),
      locked=locked, hidden=hidden,
      border=copy_border(border),
    }
    local style_index = idx - 1
    if classify(numFmtId) then datestyle[style_index] = classify(numFmtId) end
    local is_default = style_index == 0 or (not data.bold and not data.italic and not data.underline and
      not data.strike and (not data.font_name or data.font_name == "Calibri") and
      (not data.font_size or data.font_size == 11) and not data.font_color and not data.fill_color and
      not data.horizontal and not data.vertical and not data.wrap_text and not data.number_format and data.locked==nil and data.hidden==nil and not data.border)
    style_objects[style_index] = is_default and false or style_object_from_data(data)
  end
  local dxfs = xml:match("<dxfs[^>]*>(.-)</dxfs>") or ""
  for i, element in ipairs(parse_repeated_elements(dxfs, "dxf")) do dxf_objects[i - 1] = parse_dxf_style(element.body) end
  return datestyle, style_objects, dxf_objects
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

local function parse_comments(xml)
  local map, list = {}, {}
  if not xml then return map, list end
  local authors = {}
  local authors_body = xml:match("<authors>(.-)</authors>") or ""
  for text in authors_body:gmatch("<author>(.-)</author>") do authors[#authors + 1] = unescape(text) end
  local comment_body = xml:match("<commentList>(.-)</commentList>") or ""
  for attrs, inner in comment_body:gmatch("<comment%s+(.-)>(.-)</comment>") do
    local ref = xml_attr(attrs, "ref")
    local r, c
    if ref then r, c = ref_to_rowcol(ref) end
    if r then
      local author = authors[(tonumber(xml_attr(attrs, "authorId")) or 0) + 1] or ""
      local text_body = inner:match("<text>(.-)</text>") or inner
      local pieces = {}; for text in text_body:gmatch("<t[^>]*>(.-)</t>") do pieces[#pieces + 1] = unescape(text) end
      local item = { row=r, col=c, ref=ref, author=author, text=table.concat(pieces) }
      local rr = map[r]; if not rr then rr = {}; map[r] = rr end; rr[c] = item
      list[#list + 1] = item
    end
  end
  return map, list
end

local function parse_sheet(xml, shared, rich_shared, datestyle, style_objects, dxf_objects, date1904, relationships, comment_xml)
  local cells, formulas, styles, rich_text = {}, {}, {}, {}
  local row_heights, column_widths, row_hidden, column_hidden = {}, {}, {}, {}
  local maxrow, maxcol = -1, -1

  for tag in xml:gmatch("<col%s.-/>") do
    local first, last = tonumber(tag:match('min="(%d+)"')), tonumber(tag:match('max="(%d+)"'))
    local width = tonumber(tag:match('width="([^"]+)"'))
    local hidden = tag:match('hidden="([^"]+)"')
    if first and last and first >= 1 and last <= MAX_COL + 1 and first <= last then
      for col = first - 1, last - 1 do
        if width then column_widths[col] = width end
        if hidden == "1" or hidden == "true" then column_hidden[col] = true end
      end
    end
  end

  for crow in xml:gmatch("<row[%s>].-</row>") do
    local rowtag = crow:match("^(<row[^>]*>)") or ""
    local rownum = tonumber(rowtag:match('r="(%d+)"'))
    local height = tonumber(rowtag:match('ht="([^"]+)"'))
    local hidden = rowtag:match('hidden="([^"]+)"')
    if rownum and rownum >= 1 and rownum <= MAX_ROW + 1 then
      if height then row_heights[rownum - 1] = height end
      if hidden == "1" or hidden == "true" then row_hidden[rownum - 1] = true end
    end
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
          local raw_value, rich_value = decode_cell(inner, ty, shared, rich_shared)
          local display_value = raw_value
          if rich_value then local rr=rich_text[r0]; if not rr then rr={}; rich_text[r0]=rr end; rr[c0]=rich_value end
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
  local data_validations = {}
  local dvs = xml:match("<dataValidations[^>]*>(.-)</dataValidations>") or ""
  for attrs, inner in dvs:gmatch("<dataValidation%s+(.-)>(.-)</dataValidation>") do
    local ref = xml_attr(attrs, "sqref")
    if ref and not ref:find(" ", 1, true) then
      local kind_xml = xml_attr(attrs, "type") or "none"
      local kind = ({textLength="text_length"})[kind_xml] or kind_xml
      local item = {
        ref=ref, type=kind, operator=VALIDATION_OPERATOR_FROM_XML[xml_attr(attrs, "operator")],
        allow_blank=xml_attr(attrs, "allowBlank") == "1" or xml_attr(attrs, "allowBlank") == "true",
        show_input_message=xml_attr(attrs, "showInputMessage") == "1" or xml_attr(attrs, "showInputMessage") == "true",
        show_error_message=xml_attr(attrs, "showErrorMessage") == "1" or xml_attr(attrs, "showErrorMessage") == "true",
        show_dropdown=not (xml_attr(attrs, "showDropDown") == "1" or xml_attr(attrs, "showDropDown") == "true"),
        prompt_title=(xml_attr(attrs, "promptTitle") and unescape(xml_attr(attrs, "promptTitle"))),
        prompt=(xml_attr(attrs, "prompt") and unescape(xml_attr(attrs, "prompt"))),
        error_title=(xml_attr(attrs, "errorTitle") and unescape(xml_attr(attrs, "errorTitle"))),
        error=(xml_attr(attrs, "error") and unescape(xml_attr(attrs, "error"))), error_style=xml_attr(attrs, "errorStyle"),
        formula1=unescape(inner:match("<formula1>(.-)</formula1>") or ""),
        formula2=unescape(inner:match("<formula2>(.-)</formula2>") or ""),
      }
      if item.formula1 == "" then item.formula1 = nil end; if item.formula2 == "" then item.formula2 = nil end
      if kind == "list" and item.formula1 and item.formula1:match('^".*"$') then
        item.values = {}; for value in item.formula1:sub(2,-2):gmatch("[^,]+") do item.values[#item.values + 1] = value end
      end
      data_validations[#data_validations + 1] = item
    end
  end
  local conditional_formats = {}
  local function parse_cfvo(body)
    local out={}
    for attrs in body:gmatch("<cfvo%s+(.-)/>") do
      local raw_type=xml_attr(attrs,"type")
      out[#out+1]={type=CFVO_FROM_XML[raw_type] or raw_type,value=tonumber(xml_attr(attrs,"val")) or xml_attr(attrs,"val")}
    end
    return out
  end
  for cfattrs, cfbody in xml:gmatch("<conditionalFormatting%s+(.-)>(.-)</conditionalFormatting>") do
    local ref = xml_attr(cfattrs, "sqref")
    for attrs, inner in cfbody:gmatch("<cfRule%s+(.-)>(.-)</cfRule>") do
      local xmltype = xml_attr(attrs, "type")
      local kind = CF_TYPE_FROM_XML[xmltype] or xmltype
      if xmltype=="top10" then kind=(xml_attr(attrs,"bottom")=="1" or xml_attr(attrs,"bottom")=="true") and "bottom" or "top" end
      if xmltype=="aboveAverage" then kind=(xml_attr(attrs,"aboveAverage")=="0" or xml_attr(attrs,"aboveAverage")=="false") and "below_average" or "above_average" end
      local formulas = {}; for formula in inner:gmatch("<formula>(.-)</formula>") do formulas[#formulas + 1] = unescape(formula) end
      local item={ref=ref,type=kind,operator=VALIDATION_OPERATOR_FROM_XML[xml_attr(attrs,"operator")],text=(xml_attr(attrs,"text") and unescape(xml_attr(attrs,"text"))),
        stop_if_true=xml_attr(attrs,"stopIfTrue")=="1" or xml_attr(attrs,"stopIfTrue")=="true",priority=tonumber(xml_attr(attrs,"priority")),
        style=dxf_objects[tonumber(xml_attr(attrs,"dxfId") or "-1")],formula1=formulas[1],formula2=formulas[2]}
      if kind=="color_scale" then
        local body=inner:match("<colorScale>(.-)</colorScale>") or ""; local v=parse_cfvo(body); item.min=v[1]; item.mid=#v==3 and v[2] or nil; item.max=v[#v]
        local colors={}; for c in body:gmatch('<color[^>]-rgb="([%x]+)"') do colors[#colors+1]=normalize_color(c,"couleur conditionnelle",0) end
        item.min_color=colors[1]; item.mid_color=#colors==3 and colors[2] or nil; item.max_color=colors[#colors]
      elseif kind=="data_bar" then
        local body=inner:match("<dataBar%s*(.-)>(.-)</dataBar>"); local dattrs,dbody=inner:match("<dataBar%s*(.-)>(.-)</dataBar>"); local v=parse_cfvo(dbody or "")
        item.start=v[1]; item.finish=v[2]; local c=(dbody or ""):match('<color[^>]-rgb="([%x]+)"'); item.color=c and normalize_color(c,"barre",0) or nil
        item.show_value=not (xml_attr(dattrs,"showValue")=="0" or xml_attr(dattrs,"showValue")=="false")
      elseif kind=="icon_set" then
        local iattrs,ibody=inner:match("<iconSet%s*(.-)>(.-)</iconSet>"); item.icons=ICON_SETS_FROM_XML[xml_attr(iattrs,"iconSet")] or xml_attr(iattrs,"iconSet")
        item.thresholds=parse_cfvo(ibody or ""); item.show_value=not (xml_attr(iattrs,"showValue")=="0" or xml_attr(iattrs,"showValue")=="false"); item.reverse=xml_attr(iattrs,"reverse")=="1" or xml_attr(iattrs,"reverse")=="true"
      elseif kind=="top" or kind=="bottom" then item.rank=tonumber(xml_attr(attrs,"rank")) or 10; item.percent=xml_attr(attrs,"percent")=="1" or xml_attr(attrs,"percent")=="true"; item.bottom=kind=="bottom"
      elseif kind=="above_average" or kind=="below_average" then item.above_average=kind=="above_average" end
      conditional_formats[#conditional_formats+1]=item
    end
  end
  table.sort(conditional_formats, function(a,b) return (a.priority or 0) < (b.priority or 0) end)
  local tab_color = xml:match('<tabColor[^>]-rgb="([%x]+)"')
  if tab_color then tab_color = normalize_color(tab_color, "couleur d'onglet", 0) end
  local page_margins, page_setup, print_options, header_footer
  local margin_attrs = xml:match("<pageMargins%s+(.-)/>")
  if margin_attrs then
    page_margins={}; for _,key in ipairs({"left","right","top","bottom","header","footer"}) do page_margins[key]=tonumber(xml_attr(margin_attrs,key)) end
  end
  local setup_attrs = xml:match("<pageSetup%s+(.-)/>")
  if setup_attrs then page_setup={ orientation=xml_attr(setup_attrs,"orientation"), paper_size=tonumber(xml_attr(setup_attrs,"paperSize")), scale=tonumber(xml_attr(setup_attrs,"scale")), fit_to_width=tonumber(xml_attr(setup_attrs,"fitToWidth")), fit_to_height=tonumber(xml_attr(setup_attrs,"fitToHeight")) } end
  local option_attrs = xml:match("<printOptions%s+(.-)/>")
  if option_attrs then print_options={ horizontal_centered=xml_attr(option_attrs,"horizontalCentered")=="1" or xml_attr(option_attrs,"horizontalCentered")=="true", vertical_centered=xml_attr(option_attrs,"verticalCentered")=="1" or xml_attr(option_attrs,"verticalCentered")=="true", grid_lines=xml_attr(option_attrs,"gridLines")=="1" or xml_attr(option_attrs,"gridLines")=="true", headings=xml_attr(option_attrs,"headings")=="1" or xml_attr(option_attrs,"headings")=="true" } end
  local hf_attrs, hf_body = xml:match("<headerFooter%s*(.-)>(.-)</headerFooter>")
  if hf_body then
    local function split_hf(text)
      text=unescape(text or ""); local out={}; local current,buf
      local i=1
      while i<=#text do
        local code=text:sub(i,i+1)
        if code=="&L" or code=="&C" or code=="&R" then if current then out[current]=table.concat(buf) end; current=code:sub(2); buf={}; i=i+2
        else if current then buf[#buf+1]=text:sub(i,i) end; i=i+1 end
      end
      if current then out[current]=table.concat(buf) end; return out
    end
    local h=split_hf(hf_body:match("<oddHeader>(.-)</oddHeader>")); local f=split_hf(hf_body:match("<oddFooter>(.-)</oddFooter>"))
    header_footer={ header_left=h.L,header_center=h.C,header_right=h.R,footer_left=f.L,footer_center=f.C,footer_right=f.R,
      different_first=xml_attr(hf_attrs,"differentFirst")=="1" or xml_attr(hf_attrs,"differentFirst")=="true",
      different_odd_even=xml_attr(hf_attrs,"differentOddEven")=="1" or xml_attr(hf_attrs,"differentOddEven")=="true" }
  end
  local sheet_protection
  local prot_attrs=xml:match("<sheetProtection%s+(.-)/>")
  if prot_attrs then
    sheet_protection={password_hash=xml_attr(prot_attrs,"password")}
    local names={select_locked_cells="selectLockedCells",select_unlocked_cells="selectUnlockedCells",format_cells="formatCells",format_columns="formatColumns",format_rows="formatRows",insert_columns="insertColumns",insert_rows="insertRows",insert_hyperlinks="insertHyperlinks",delete_columns="deleteColumns",delete_rows="deleteRows",sort="sort",auto_filter="autoFilter",pivot_tables="pivotTables",objects="objects",scenarios="scenarios"}
    for key,attr in pairs(names) do sheet_protection[key]=xml_attr(prot_attrs,attr)=="1" or xml_attr(prot_attrs,attr)=="true" end
  end
  local row_page_breaks,column_page_breaks={},{}
  local rb=xml:match("<rowBreaks[^>]*>(.-)</rowBreaks>") or ""; for attrs in rb:gmatch("<brk%s+(.-)/>") do local id=tonumber(xml_attr(attrs,"id")); if id then row_page_breaks[#row_page_breaks+1]=id end end
  local cb=xml:match("<colBreaks[^>]*>(.-)</colBreaks>") or ""; for attrs in cb:gmatch("<brk%s+(.-)/>") do local id=tonumber(xml_attr(attrs,"id")); if id then column_page_breaks[#column_page_breaks+1]=id end end
  local sparklines={}; local plain_ext=xml:gsub("x14:",""):gsub("xm:","")
  for attrs,body in plain_ext:gmatch("<sparklineGroup%s*(.-)>(.-)</sparklineGroup>") do
    local source=unescape(body:match("<f>(.-)</f>") or ""); local target=unescape(body:match("<sqref>(.-)</sqref>") or "")
    local st=xml_attr(attrs,"type") or "line"; if st=="stacked" then st="win_loss" end
    local item={source=source,target=target,type=st,show_markers=xml_attr(attrs,"markers")=="1",show_high=xml_attr(attrs,"high")=="1",show_low=xml_attr(attrs,"low")=="1",show_first=xml_attr(attrs,"first")=="1",show_last=xml_attr(attrs,"last")=="1",show_negative=xml_attr(attrs,"negative")=="1",right_to_left=xml_attr(attrs,"rightToLeft")=="1"}
    local colors={color="colorSeries",negative_color="colorNegative",high_color="colorHigh",low_color="colorLow",first_color="colorFirst",last_color="colorLast"}
    for key,tag in pairs(colors) do local c=body:match('<'..tag..'[^>]-rgb="([%x]+)"'); if c then item[key]=normalize_color(c,"sparkline",0) end end
    sparklines[#sparklines+1]=item
  end
  local comment_map, comment_list = parse_comments(comment_xml)
  return cells, formulas, styles, maxrow, maxcol, row_heights, column_widths,
    freeze_rows, freeze_cols, auto_filter, merged, hyperlink_ranges,
    row_hidden, column_hidden, data_validations, conditional_formats, tab_color, comment_map, comment_list,
    page_margins, page_setup, print_options, header_footer, rich_text, sheet_protection, row_page_breaks, column_page_breaks, sparklines
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


local function parse_table_part(xml)
  local tag=xml:match("<table%s+(.-)>") or xml:match("<table%s+(.-)/>") or ""
  local style_attrs=xml:match("<tableStyleInfo%s+(.-)/>") or ""
  local columns={}; for attrs in xml:gmatch("<tableColumn%s+(.-)/>") do columns[#columns+1]=xml_attr(attrs,"name") end
  return { name=xml_attr(tag,"displayName") or xml_attr(tag,"name"), ref=xml_attr(tag,"ref"), style=xml_attr(style_attrs,"name"),
    show_first_column=xml_attr(style_attrs,"showFirstColumn")=="1", show_last_column=xml_attr(style_attrs,"showLastColumn")=="1",
    show_row_stripes=xml_attr(style_attrs,"showRowStripes")=="1", show_column_stripes=xml_attr(style_attrs,"showColumnStripes")=="1", columns=columns }
end

local function parse_chart_part(xml)
  local plain=xml:gsub("c:",""):gsub("a:","")
  local kind="line"
  if plain:find("<pieChart",1,true) then kind="pie"
  elseif plain:find("<doughnutChart",1,true) then kind="doughnut"
  elseif plain:find("<areaChart",1,true) then kind="area"
  elseif plain:find("<scatterChart",1,true) then kind="scatter"
  elseif plain:find("<barChart",1,true) then kind=(plain:match('<barDir%s+val="([^"]+)"')=="bar") and "bar" or "column" end
  local title=plain:match("<title>.-<t>(.-)</t>.-</title>"); if title then title=unescape(title) end
  local chart_tag=({line='lineChart',column='barChart',bar='barChart',pie='pieChart',doughnut='doughnutChart',area='areaChart',scatter='scatterChart'})[kind]
  local chart_body=plain:match('<'..chart_tag..'>(.-)</'..chart_tag..'>') or plain
  local series={}; local categories
  for body in chart_body:gmatch("<ser>(.-)</ser>") do
    local name=body:match("<tx>.-<v>(.-)</v>.-</tx>"); local name_ref=body:match("<tx>.-<strRef>.-<f>(.-)</f>")
    local cat=body:match("<cat>.-<f>(.-)</f>"); local values=body:match("<val>.-<f>(.-)</f>") or body:match("<yVal>.-<f>(.-)</f>"); local xv=body:match("<xVal>.-<f>(.-)</f>")
    if not categories and cat then categories=unescape(cat) end
    local color=body:match('<srgbClr%s+val="([%x]+)"'); if color then color=normalize_color(color,'couleur de série',0) end
    local width=tonumber(body:match('<ln[^>]-w="(%d+)"'))
    series[#series+1]={ name=name and unescape(name) or nil, name_ref=name_ref and unescape(name_ref) or nil, values=values and unescape(values) or nil,
      x_values=xv and unescape(xv) or nil, color=color, marker=body:match('<symbol%s+val="([^"]+)"'),
      line_width=width and width/12700 or nil, smooth=body:match('<smooth%s+val="1"')~=nil }
  end
  local grouping=chart_body:match('<grouping%s+val="([^"]+)"') or 'standard'; if grouping=='percentStacked' then grouping='percent_stacked' end
  local legendpos=plain:match('<legendPos%s+val="([^"]+)"')
  local hole=tonumber(chart_body:match('<holeSize%s+val="([^"]+)"'))

  local label_positions={ctr='center',inEnd='inside_end',inBase='inside_base',outEnd='outside_end',bestFit='best_fit',l='left',r='right',t='top',b='bottom'}
  local data_labels
  local label_body=chart_body:match("<dLbls>(.-)</dLbls>")
  if label_body then
    local function flag(tag) return label_body:match('<'..tag..'%s+val="1"')~=nil or label_body:match('<'..tag..'%s+val="true"')~=nil end
    local pos=label_body:match('<dLblPos%s+val="([^"]+)"')
    data_labels={show_value=flag('showVal'),show_percent=flag('showPercent'),show_category=flag('showCatName'),show_series_name=flag('showSerName'),position=label_positions[pos]}
  end

  local function parse_axis(body)
    if not body then return {} end
    local axis_title=body:match("<title>.-<t>(.-)</t>.-</title>")
    local min=tonumber(body:match('<min%s+val="([^"]+)"'))
    local max=tonumber(body:match('<max%s+val="([^"]+)"'))
    local fmt=body:match('<numFmt%s+[^>]-formatCode="([^"]*)"')
    return { title=axis_title and unescape(axis_title) or nil, min=min, max=max,
      number_format=fmt and unescape(fmt) or nil, major_gridlines=body:find("<majorGridlines",1,true)~=nil }
  end
  local x_axis,y_axis={},{}
  local function capture_axes(tag)
    for body in plain:gmatch('<'..tag..'>(.-)</'..tag..'>') do
      local pos=body:match('<axPos%s+val="([^"]+)"')
      if pos=='b' or pos=='t' then x_axis=parse_axis(body)
      elseif pos=='l' or pos=='r' then y_axis=parse_axis(body) end
    end
  end
  capture_axes('catAx'); capture_axes('valAx')

  return { type=kind,title=title,categories=categories,series=series,legend=plain:find("<legend",1,true)~=nil,
    legend_position=LEGEND_POSITIONS_FROM_XML[legendpos] or 'right',grouping=grouping,hole_size=hole,
    data_labels=data_labels,x_axis=x_axis,y_axis=y_axis }
end

local function parse_drawing_part(xml, relxml, drawing_path, byname, data, limits)
  local relationships=parse_relationships(relxml); local plain=xml:gsub("xdr:",""):gsub("a:",""):gsub("c:","")
  local images,charts={},{}; local base=drawing_path:match("^(.*[/])") or ""
  local function process_anchor(anchor)
    local from=anchor:match("<from>(.-)</from>") or ""
    local col=tonumber(from:match("<col>(%d+)</col>")) or 0
    local row=tonumber(from:match("<row>(%d+)</row>")) or 0
    local cx,cy=anchor:match('<ext%s+[^>]-cx="(%d+)"[^>]-cy="(%d+)"')
    if not cx then cx,cy=anchor:match('<ext%s+[^>]-cy="(%d+)"[^>]-cx="(%d+)"'); cx,cy=cy,cx end
    local width,height=(tonumber(cx) or 0)/9525,(tonumber(cy) or 0)/9525
    if anchor:find("<pic>",1,true) then
      local rid=anchor:match('blip[^>]-r:embed="([^"]+)"'); local rel=rid and relationships[rid]
      if rel then
        local path=package_path(base,rel.target); local entry=path and byname[path]; local binary=entry and zip_extract(data,entry,limits) or nil
        local format=path and path:match("%.([^.]+)$") or nil; if format=="jpg" then format="jpeg" end
        local nattrs=anchor:match("<cNvPr%s+([^>]-)/?>") or ""
        images[#images+1]={row=row,col=col,width=width,height=height,format=format,name=xml_attr(nattrs,"name"),alt_text=xml_attr(nattrs,"descr"),data=binary,path=path}
      end
    elseif anchor:find("<graphicFrame",1,true) then
      local rid=anchor:match('chart[^>]-r:id="([^"]+)"'); local rel=rid and relationships[rid]
      if rel then
        local path=package_path(base,rel.target); local entry=path and byname[path]
        if entry then
          local chart=parse_chart_part(validate_xml_document(zip_extract(data,entry,limits),path))
          chart.row,chart.col,chart.width,chart.height,chart.path=row,col,width,height,path
          charts[#charts+1]=chart
        end
      end
    end
  end
  for anchor in plain:gmatch("<oneCellAnchor[^>]*>(.-)</oneCellAnchor>") do process_anchor(anchor) end
  for anchor in plain:gmatch("<twoCellAnchor[^>]*>(.-)</twoCellAnchor>") do process_anchor(anchor) end
  return images,charts
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
    local state = tag:match('state="([^"]+)"') or "visible"
    if state == "veryHidden" then state = "very_hidden" end
    sheets[#sheets + 1] = { name = name, rid = rid, visibility = state }
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
  local view = wbxml:match("<workbookView%s.-/>") or ""
  local active_sheet = (tonumber(view:match('activeTab="(%d+)"')) or 0) + 1
  if not sheets[active_sheet] then active_sheet = 1 end
  local defined_names = {}
  local body = wbxml:match("<definedNames>(.-)</definedNames>") or ""
  for attrs, reference in body:gmatch("<definedName%s+(.-)>(.-)</definedName>") do
    local name = xml_attr(attrs, "name"); local local_sheet=(tonumber(xml_attr(attrs,"localSheetId")) and tonumber(xml_attr(attrs,"localSheetId"))+1 or nil); reference=unescape(reference)
    if name == "_xlnm.Print_Area" and local_sheet and sheets[local_sheet] then
      local ref=reference:match("!(.+)$"); if ref then sheets[local_sheet].print_area=ref:gsub("%$","") end
    elseif name == "_xlnm.Print_Titles" and local_sheet and sheets[local_sheet] then
      for part in reference:gmatch("[^,]+") do local ref=part:match("!(.+)$"); if ref then if ref:match("^%$%d+:%$%d+$") then sheets[local_sheet].repeat_rows=ref:gsub("%$","") elseif ref:match("^%$[A-Z]+:%$[A-Z]+$") then sheets[local_sheet].repeat_cols=ref:gsub("%$","") end end end
    elseif name then defined_names[#defined_names + 1] = {
      name=unescape(name), reference=reference, local_sheet=local_sheet,
      hidden=xml_attr(attrs, "hidden") == "1" or xml_attr(attrs, "hidden") == "true",
      comment=(xml_attr(attrs, "comment") and unescape(xml_attr(attrs, "comment"))),
    } end
  end
  local workbook_protection
  local pattrs=wbxml:match("<workbookProtection%s+(.-)/>")
  if pattrs then workbook_protection={password_hash=xml_attr(pattrs,"workbookPassword"),structure=xml_attr(pattrs,"lockStructure")=="1" or xml_attr(pattrs,"lockStructure")=="true",windows=xml_attr(pattrs,"lockWindows")=="1" or xml_attr(pattrs,"lockWindows")=="true"} end
  return sheets, date1904, active_sheet, defined_names, workbook_protection
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

function ReadSheet:is_row_hidden(row)
  check_index(row, "row", 3); return self.row_hidden[row] == true
end

function ReadSheet:is_column_hidden(col)
  check_index(col, "col", 3); return self.column_hidden[col] == true
end

function ReadSheet:get_tab_color() return self.tab_color end
function ReadSheet:get_visibility() return self.visibility end

function ReadSheet:get_data_validations()
  local out = {}; for i, value in ipairs(self.data_validations) do out[i] = copy_validation(value) end; return out
end

function ReadSheet:get_conditional_formats()
  local out = {}; for i, value in ipairs(self.conditional_formats) do out[i] = copy_conditional_format(value) end; return out
end

function ReadSheet:get_comment(row, col)
  check_index(row, "row", 3); check_index(col, "col", 3)
  local rr = self.comment_map[row]; local value = rr and rr[col]
  return value and { author=value.author, text=value.text, ref=value.ref } or nil
end

function ReadSheet:get_comments()
  local out = {}; for i, value in ipairs(self.comment_list) do out[i] = { row=value.row, col=value.col, ref=value.ref, author=value.author, text=value.text } end; return out
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


function ReadSheet:get_images()
  local out={}; for i,v in ipairs(self.image_list) do out[i]=shallow_copy(v) end; return out
end
function ReadSheet:get_charts()
  local out={}; for i,v in ipairs(self.chart_list) do local c=shallow_copy(v); c.series={}; for j,ser in ipairs(v.series or {}) do c.series[j]=shallow_copy(ser) end; out[i]=c end; return out
end
function ReadSheet:get_tables()
  local out={}; for i,v in ipairs(self.table_list) do local t=shallow_copy(v); t.columns={table.unpack(v.columns or {})}; out[i]=t end; return out
end
function ReadSheet:get_page_margins() return self.page_margins and shallow_copy(self.page_margins) or nil end
function ReadSheet:get_page_setup()
  if not self.page_setup and not self.print_options then return nil end
  local out=shallow_copy(self.page_setup or {}); for k,v in pairs(self.print_options or {}) do out[k]=v end; return out
end
function ReadSheet:get_header_footer() return self.header_footer and shallow_copy(self.header_footer) or nil end
function ReadSheet:get_rich_text(row,col)
  check_index(row,"row",3); check_index(col,"col",3); local rr=self.rich_text[row]; local v=rr and rr[col]; return v
end
function ReadSheet:get_protection() return self.sheet_protection and shallow_copy(self.sheet_protection) or nil end
function ReadSheet:get_row_page_breaks() return {table.unpack(self.row_page_breaks or {})} end
function ReadSheet:get_column_page_breaks() return {table.unpack(self.column_page_breaks or {})} end
function ReadSheet:get_sparklines() local out={}; for i,v in ipairs(self.sparkline_list or {}) do out[i]=shallow_copy(v) end; return out end
function ReadSheet:get_print_area() return self.print_area end
function ReadSheet:get_print_titles() return self.repeat_rows, self.repeat_cols end

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

function ReadWB:get_active_sheet()
  return self._sheets[self._active_sheet] and self._sheets[self._active_sheet].name or nil
end

function ReadWB:get_defined_names()
  local out = {}; for i, item in ipairs(self._defined_names) do out[i] = shallow_copy(item) end; return out
end

function ReadWB:get_defined_name(name, local_sheet)
  if type(name) ~= "string" then error("xlsx: get_defined_name attend un nom", 2) end
  if type(local_sheet) == "string" then
    local resolved
    for i, sheet in ipairs(self._sheets) do if sheet.name == local_sheet then resolved = i; break end end
    if not resolved then return nil end
    local_sheet = resolved
  elseif local_sheet ~= nil and math.type(local_sheet) ~= "integer" then
    error("xlsx: local_sheet doit être un nom ou un index 1-based", 2)
  end
  for _, item in ipairs(self._defined_names) do
    if item.name:lower() == name:lower() and item.local_sheet == local_sheet then return shallow_copy(item) end
  end
  return nil
end

function ReadWB:get_properties() return shallow_copy(self._properties or {}) end
function ReadWB:get_protection() return self._workbook_protection and shallow_copy(self._workbook_protection) or nil end

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
    local comment_xml; local image_list,chart_list,table_list={},{},{}
    for _, rel in pairs(relationships) do
      if rel.rel_type and rel.rel_type:match("/comments$") then
        local path = package_path("xl/worksheets/", rel.target)
        local ce = path and self._byname[path]
        if ce then comment_xml = validate_xml_document(zip_extract(self._data, ce, self._limits), path) end
      elseif rel.rel_type and rel.rel_type:match("/drawing$") then
        local path=package_path("xl/worksheets/",rel.target); local de=path and self._byname[path]
        if de then local drawing_xml=validate_xml_document(zip_extract(self._data,de,self._limits),path); local rp=rels_path_for_part(path); local re=self._byname[rp]; local rxml=re and validate_xml_document(zip_extract(self._data,re,self._limits),rp) or nil; image_list,chart_list=parse_drawing_part(drawing_xml,rxml,path,self._byname,self._data,self._limits) end
      elseif rel.rel_type and rel.rel_type:match("/table$") then
        local path=package_path("xl/worksheets/",rel.target); local te=path and self._byname[path]
        if te then local t=parse_table_part(validate_xml_document(zip_extract(self._data,te,self._limits),path)); t.path=path; table_list[#table_list+1]=t end
      end
    end
    local values={ parse_sheet(xml, self._shared, self._rich_shared, self._datestyle, self._style_objects, self._dxf_objects, self._date1904, relationships, comment_xml) }
    values[29],values[30],values[31]=image_list,chart_list,table_list
    return values
  end)
  if not ok then return nil, tostring(result) end
  local values = result
  local rs = setmetatable({
    name=def.name, cells=values[1], formulas=values[2], styles=values[3],
    maxrow=values[4], maxcol=values[5], row_heights=values[6], column_widths=values[7],
    freeze_rows=values[8], freeze_cols=values[9], auto_filter=values[10],
    merged_ranges=values[11], hyperlink_ranges=values[12],
    row_hidden=values[13], column_hidden=values[14], data_validations=values[15],
    conditional_formats=values[16], tab_color=values[17], comment_map=values[18], comment_list=values[19],
    page_margins=values[20], page_setup=values[21], print_options=values[22], header_footer=values[23],
    rich_text=values[24], sheet_protection=values[25], row_page_breaks=values[26], column_page_breaks=values[27], sparkline_list=values[28],
    image_list=values[29], chart_list=values[30], table_list=values[31],
    print_area=def.print_area, repeat_rows=def.repeat_rows, repeat_cols=def.repeat_cols,
    visibility=def.visibility or "visible",
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

local function parse_document_properties(core, app)
  local out={}
  if core then
    local tags={creator="creator",title="title",description="description",subject="subject",category="category",keywords="keywords"}
    for key,tag in pairs(tags) do local value=core:match("<[^>]*"..tag.."[^>]*>(.-)</[^>]*"..tag..">"); if value then out[key]=unescape(value) end end
  end
  if app then local company=app:match("<Company>(.-)</Company>"); local manager=app:match("<Manager>(.-)</Manager>"); if company then out.company=unescape(company) end; if manager then out.manager=unescape(manager) end end
  return out
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
    local shared, rich_shared = {}, {}
    if ssxml then shared, rich_shared = parse_shared_strings(ssxml) end
    local sheets, date1904, active_sheet, defined_names, workbook_protection = parse_workbook(wbxml, relsxml)
    local datestyle, style_objects, dxf_objects = parse_styles(part("xl/styles.xml"))
    return setmetatable({
      _data=data, _byname=byname, _shared=shared, _rich_shared=rich_shared, _sheets=sheets,
      _cache={}, _datestyle=datestyle, _style_objects=style_objects, _dxf_objects=dxf_objects,
      _date1904=date1904, _limits=limits, _active_sheet=active_sheet, _defined_names=defined_names,
      _workbook_protection=workbook_protection, _properties=parse_document_properties(part("docProps/core.xml"),part("docProps/app.xml")),
    }, ReadWB)
  end)
  if not ok then return nil, tostring(result) end
  return result
end

return xlsx
