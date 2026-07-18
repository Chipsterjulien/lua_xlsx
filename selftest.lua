-- Copyright (C) 2026  Freyermuth Julien
--
-- This program is free software: you can redistribute it and/or modify
-- it under the terms of the GNU General Public License as published by
-- the Free Software Foundation, either version 3 of the License, or
-- (at your option) any later version.

--- selftest.lua — auto-test des modules xlsx + dataframe.
--- Cible principale : Babet (Lua 5.5), avec compatibilité Lua standard 5.3+.
---
--- Sous Babet, placer ce fichier sous le nom main.lua à côté de xlsx.lua et
--- dataframe.lua, puis lancer : babet <dossier>
--- Sous Lua standard : lua selftest.lua

local xlsx = require("xlsx")
local DF   = require("dataframe")

local fails = 0
local function check(cond, msg)
  if cond then
    print("  ok   " .. msg)
  else
    print("  FAIL " .. msg)
    fails = fails + 1
  end
end

local function expect_error(fn, msg)
  local ok = pcall(fn)
  check(not ok, msg)
end

local function write_file(path, data)
  local f, err = io.open(path, "wb")
  assert(f, err)
  assert(f:write(data))
  assert(f:close())
end

local function read_file(path)
  local f, err = io.open(path, "rb")
  assert(f, err)
  local data = assert(f:read("a"))
  assert(f:close())
  return data
end

print("Lua : " .. _VERSION)
print("Babet : " .. tostring(type(rawget(_G, "babet")) == "table"))
print("string.pack=" .. tostring(string.pack ~= nil)
  .. "  math.type=" .. tostring(math.type ~= nil)
  .. "  utf8.char=" .. tostring(utf8 and utf8.char ~= nil))

local TMP = "._xlsx_selftest.xlsx"
local TMP2 = "._xlsx_selftest_2.xlsx"
local BAD = "._xlsx_selftest_bad.xlsx"

-- ---- ÉCRITURE puis LECTURE -------------------------------------------------
print("\n[1] écriture + relecture xlsx")
do
  local wb = xlsx.new()
  local s = wb:add_sheet("Données")
  s:write(0, 0, "Nom"); s:write(0, 1, "Âge"); s:write(0, 2, "Actif"); s:write(0, 3, "Note")
  s:append_row({ "Alice <x>", 30, true,  12.5 })
  s:append_row({ "Bob & Co",  25, false, -0.5 })
  s:write(4, 1, 999)
  local s2 = wb:add_sheet("Autre")
  s2:write_rows({ { "a", "b" }, { 1, 2 } })
  assert(wb:save(TMP, { use_babet = false }))

  local r = assert(xlsx.open(TMP, { use_babet = false }))
  check(table.concat(r:sheet_names(), ",") == "Données,Autre", "feuilles dans l'ordre")
  check(r:date_system() == "1900", "système de dates 1900 détecté")
  local sh = assert(r:sheet("Données"))
  check(sh:read(0, 0) == "Nom",        "string + unicode entête")
  check(sh:read(0, 1) == "Âge",        "accent préservé")
  check(sh:read(1, 0) == "Alice <x>",  "entité XML déséchappée")
  check(sh:read(1, 1) == 30,           "entier")
  check(sh:read(1, 2) == true,         "booléen true")
  check(sh:read(2, 2) == false,        "booléen false (pas nil)")
  check(sh:read(1, 3) == 12.5,         "flottant")
  check(sh:read(2, 3) == -0.5,         "flottant négatif")
  check(sh:read(3, 0) == nil,          "ligne vide -> nil")
  check(sh:read(4, 1) == 999,          "cellule éparse")
  local mr, mc = sh:dims()
  check(mr == 4 and mc == 3,           "dimensions 0-based (4,3)")
end

-- ---- DATAFRAME -------------------------------------------------------------
print("\n[2] dataframe : pipeline et non-destruction")
do
  local r = assert(xlsx.open(TMP, { use_babet = false }))
  local d = DF.from_sheet(assert(r:sheet("Données")), { header = true })
  check(d:nrow() == 4, "from_sheet : 4 lignes (dont vides)")

  d = d:filter(function(row) return row["Nom"] ~= nil end)
  check(d:nrow() == 2, "filter : 2 lignes nommées")

  d = d:mutate("Double", function(row) return (row["Âge"] or 0) * 2 end)
  check(d:column("Double")[1] == 60, "mutate : Double=60")

  local g = d:groupby("Actif"):agg({
    moy_age = { "mean", "Âge" },
    n       = { "count" },
    note_min = { "min", "Note" },
  })
  check(g:nrow() == 2, "groupby Actif : 2 groupes")

  assert(xlsx.write_rows(TMP2, g:to_rows(), { sheet = "g", use_babet = false }))
  local back = DF.from_sheet(assert(assert(xlsx.open(TMP2, { use_babet = false })):sheet("g")), { header = true })
  check(back:ncol() == 4, "round-trip : 4 colonnes (Actif + 3 agg)")

  local source = DF.from_records({ { x = 2 }, { x = 1 } }, { "x" })
  local filtered = source:filter(function() return true end)
  filtered.rows[1].x = 99
  check(source.rows[1].x == 2, "filter copie les records")
  local sorted = source:sort("x")
  sorted.rows[1].x = 88
  check(source.rows[2].x == 1, "sort copie les records")
  local headed = source:head(1)
  headed.rows[1].x = 77
  check(source.rows[1].x == 2, "head copie les records")
  local iterrow = source:iter()()
  iterrow.x = 66
  check(source.rows[1].x == 2, "iter renvoie une copie")

  expect_error(function()
    DF.from_records({ { a = 1, b = 2 } }, { "a", "b" }):rename({ a = "x", b = "x" })
  end, "rename refuse les collisions")
  expect_error(function()
    DF.from_rows({ { "x", "x" }, { 1, 2 } }, { header = true })
  end, "header refuse les colonnes dupliquées")

  local collision = DF.from_records({
    { a = "x", b = "y\2string\1z" },
    { a = "x\2string\1y", b = "z" },
  }, { "a", "b" }):groupby("a", "b"):count()
  check(collision:nrow() == 2, "groupby sans collision sur chaînes binaires")
end

-- ---- DATES -----------------------------------------------------------------
print("\n[3] dates 1900 et 1904")
do
  local wb = xlsx.new()
  local s = wb:add_sheet("d")
  s:write(0, 0, xlsx.date(2024, 3, 15))
  s:write(1, 0, xlsx.datetime(2024, 3, 15, 13, 45, 30))
  s:write(2, 0, 7)
  assert(wb:save(TMP, { use_babet = false }))

  local r = assert(xlsx.open(TMP, { use_babet = false }))
  local sh = assert(r:sheet("d"))
  check(sh:read(0, 0) == "2024-03-15",          "date 1900 -> ISO")
  check(sh:read(1, 0) == "2024-03-15T13:45:30", "datetime 1900 -> ISO")
  check(sh:read(2, 0) == 7,                      "nombre normal préservé")

  local wb1904 = xlsx.new({ date_system = "1904" })
  wb1904:add_sheet("d"):write(0, 0, xlsx.datetime(2024, 3, 15, 13, 45, 30))
  assert(wb1904:save(TMP2, { use_babet = false }))
  local r1904 = assert(xlsx.open(TMP2, { use_babet = false }))
  check(r1904:date_system() == "1904", "système de dates 1904 détecté")
  check(assert(r1904:sheet("d")):read(0, 0) == "2024-03-15T13:45:30",
    "datetime 1904 correctement convertie")
end

-- ---- VALIDATIONS -----------------------------------------------------------
print("\n[4] validations XML, feuilles, dates et limites")
do
  local wb = xlsx.new()
  local long_unicode = string.rep("é", 20)
  check(pcall(function() wb:add_sheet(long_unicode) end), "31 caractères, pas 31 octets")
  expect_error(function() wb:add_sheet(long_unicode) end, "nom de feuille dupliqué refusé")
  expect_error(function()
    local x = xlsx.new(); x:add_sheet("Test"); x:add_sheet("test")
  end, "doublon ASCII insensible à la casse refusé")
  expect_error(function() xlsx.new():add_sheet("'interdit") end, "apostrophe initiale refusée")
  expect_error(function() xlsx.new():add_sheet("fin'") end, "apostrophe finale refusée")
  expect_error(function() xlsx.new():add_sheet("bad/name") end, "caractère de feuille interdit")
  expect_error(function() xlsx.new():add_sheet("x"):write(0, 0, "a\1b") end,
    "caractère XML 1.0 interdit refusé")
  expect_error(function() xlsx.date(2024, 13, 1) end, "mois invalide refusé")
  expect_error(function() xlsx.date(2023, 2, 29) end, "jour invalide refusé")
  expect_error(function() xlsx.datetime(2024, 1, 1, 24, 0, 0) end, "heure invalide refusée")
  expect_error(function() xlsx.new():add_sheet("x"):write(1048576, 0, 1) end,
    "ligne hors limites Excel refusée")

  local too_small, err = xlsx.open(TMP, { max_file_size = 16, use_babet = false })
  check(too_small == nil and type(err) == "string", "max_file_size appliqué avant parsing")
end

-- ---- INTÉGRITÉ ZIP ---------------------------------------------------------
print("\n[5] intégrité ZIP et contrat d'erreur")
do
  local wb = xlsx.new()
  wb:add_sheet("x"):write(0, 0, "Alice")
  assert(wb:save(TMP, { use_babet = false }))
  local bytes = read_file(TMP)
  local corrupt, n = bytes:gsub("Alice", "Alicf", 1)
  assert(n == 1)
  write_file(BAD, corrupt)
  local opened, err = xlsx.open(BAD, { use_babet = false })
  check(opened == nil and tostring(err):find("CRC%-32 invalide") ~= nil,
    "CRC corrompu renvoyé sous forme nil, err")
end

-- ---- INTÉGRATION BABET -----------------------------------------------------
print("\n[6] intégration Babet optionnelle")
do
  local original = rawget(_G, "babet")
  local calls = { write = 0, size = 0, test = 0 }
  _G.babet = {
    writeFileAtomic = function(path, data, opts)
      calls.write = calls.write + 1
      check(opts.overwrite == true and opts.durable == true, "options writeFileAtomic transmises")
      write_file(path, data)
      return true, nil
    end,
    fileSize = function(path)
      calls.size = calls.size + 1
      return #read_file(path), nil
    end,
    archive = {
      test = function(_path, opts)
        calls.test = calls.test + 1
        check(opts.max_entries ~= nil and opts.max_compression_ratio ~= nil,
          "limites transmises à archive.test")
        return { format = "zip" }, nil
      end,
    },
  }

  local wb = xlsx.new()
  wb:add_sheet("babet"):write(0, 0, "ok")
  local ok, save_err = wb:save(TMP2)
  check(ok == true and save_err == nil and calls.write == 1, "save utilise writeFileAtomic")
  local read, open_err = xlsx.open(TMP2)
  check(read ~= nil and open_err == nil and calls.size == 1 and calls.test == 1,
    "open utilise fileSize et archive.test")
  _G.babet = original
end


-- ---- PRÉSENTATION ET FORMULES ----------------------------------------------
print("\n[7] styles, dimensions, volets, filtres et formules")
do
  check(xlsx.VERSION == "1.1.0", "version du module 1.1.0")
  local header = xlsx.style({
    bold = true,
    fill_color = "D9EAF7",
    font_color = "112233",
    horizontal = "center",
    vertical = "center",
    wrap_text = true,
  })
  local money = xlsx.style({ number_format = "currency_eur" })
  local date_bold = xlsx.style({ bold = true })

  local wb = xlsx.new()
  local sh = wb:add_sheet("Présentation")
  sh:append_row({ "Nom", "Montant", "Date", "Total" }, header)
  sh:append_row({ "Alice", 12.5, xlsx.date(2024, 3, 15) })
  sh:write(1, 1, 12.5, money)
  sh:write(1, 2, xlsx.date(2024, 3, 15), date_bold)
  sh:append_row({ "Bob", 7.25, xlsx.date(2024, 3, 16) })
  sh:write(2, 1, 7.25, money)
  sh:write(2, 2, xlsx.date(2024, 3, 16), date_bold)
  sh:write(3, 3, xlsx.formula("=SUM(B2:B3)"), money)
  sh:set_style(4, 0, header)
  sh:set_column_width(0, 18)
  sh:set_column_width(1, 14)
  sh:set_row_height(0, 24)
  sh:freeze_panes(1, 1)
  sh:set_auto_filter("A1:D3")
  assert(wb:save(TMP2, { use_babet = false }))

  local bytes = read_file(TMP2)
  check(bytes:find('<pane xSplit="1" ySplit="1" topLeftCell="B2"', 1, true) ~= nil,
    "volets figés sérialisés")
  check(bytes:find('<autoFilter ref="A1:D3"/>', 1, true) ~= nil,
    "filtre automatique sérialisé")
  check(bytes:find('<col min="1" max="1" width="18" customWidth="1"/>', 1, true) ~= nil,
    "largeur de colonne sérialisée")
  check(bytes:find('<row r="1" ht="24" customHeight="1">', 1, true) ~= nil,
    "hauteur de ligne sérialisée")
  check(bytes:find('<f>SUM(B2:B3)</f>', 1, true) ~= nil,
    "formule sérialisée sans signe égal")
  check(bytes:find('<b/>', 1, true) ~= nil and bytes:find('fgColor rgb="FFD9EAF7"', 1, true) ~= nil,
    "police grasse et fond sérialisés")
  check(bytes:find('formatCode="#,##0.00 &quot;€&quot;"', 1, true) ~= nil,
    "format monétaire sérialisé")
  check(bytes:find('horizontal="center" vertical="center" wrapText="1"', 1, true) ~= nil,
    "alignement et retour à la ligne sérialisés")

  local reread = assert(xlsx.open(TMP2, { use_babet = false }))
  check(assert(reread:sheet("Présentation")):read(1, 2) == "2024-03-15",
    "date stylée relue comme date")

  expect_error(function() xlsx.style({ bold = "oui" }) end,
    "style refuse un booléen invalide")
  expect_error(function() xlsx.style({ fill_color = "rouge" }) end,
    "style refuse une couleur invalide")
  expect_error(function() xlsx.formula(42) end,
    "formula exige une string")
  expect_error(function() xlsx.formula("=") end,
    "formule vide refusée")
  expect_error(function() xlsx.new():add_sheet("x"):write(0, 0, 1, {}) end,
    "write refuse une table de style brute")
  expect_error(function() xlsx.new():add_sheet("x"):set_column_width(0, 256) end,
    "largeur de colonne hors limite refusée")
  expect_error(function() xlsx.new():add_sheet("x"):set_row_height(0, 410) end,
    "hauteur de ligne hors limite refusée")
  expect_error(function() xlsx.new():add_sheet("x"):freeze_panes(-1, 0) end,
    "nombre de lignes figées invalide refusé")
  expect_error(function() xlsx.new():add_sheet("x"):set_auto_filter("D3:A1") end,
    "plage de filtre inversée refusée")
end

os.remove(TMP)
os.remove(TMP2)
os.remove(BAD)

print("\n========================================")
if fails == 0 then
  print("SELFTEST : PASS  (" .. _VERSION .. ")")
  os.exit(0)
else
  print("SELFTEST : " .. fails .. " ÉCHEC(S)")
  os.exit(1)
end
