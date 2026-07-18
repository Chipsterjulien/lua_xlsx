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
  check(xlsx.VERSION == "1.5.0", "version du module 1.5.0")
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


-- ---- STRUCTURE, STYLES ENRICHIS, HYPERLIENS ET LECTURE ---------------------
print("\n[8] fusions, bordures, polices, hyperliens et lecture de présentation")
do
  local rich = xlsx.style({
    bold = true,
    italic = true,
    underline = "double",
    strike = true,
    font_name = "Liberation Sans",
    font_size = 14,
    font_color = "112233",
    fill_color = "DDEEFF",
    horizontal = "center",
    vertical = "center",
    border = {
      left = { style = "thin", color = "FF0000" },
      right = { style = "double" },
      top = { style = "dashed", color = "00AA00" },
      bottom = { style = "dotted" },
    },
  })
  check(xlsx.is_style(rich), "is_style reconnaît un style")
  local options = xlsx.style_options(rich)
  options.border.left.style = "thick"
  check(rich.border.left.style == "thin", "style_options renvoie une copie profonde")

  local wb = xlsx.new()
  local sh = wb:add_sheet("Présentation 1.2")
  sh:write(0, 0, "Titre", rich)
  sh:write(0, 1, "supprimé par la fusion")
  sh:merge_cells("A1:D1")
  sh:write(1, 0, xlsx.hyperlink("https://example.com/?a=1&b=2", "Site", {
    tooltip = "Ouvrir le site",
  }))
  sh:write(2, 0, "Aller à la cible")
  sh:set_hyperlink(2, 0, "'Cible'!A1", { internal = true })
  sh:write(3, 0, xlsx.formula("SUM(B2:B3)", 19.75), rich)
  sh:set_column_width(0, 22)
  sh:set_row_height(0, 30)
  sh:freeze_panes(1, 2)
  sh:set_auto_filter("A2:D4")
  wb:add_sheet("Cible"):write(0, 0, "ici")
  assert(wb:save(TMP2, { use_babet = false }))

  local read = assert(xlsx.open(TMP2, { use_babet = false }))
  local rs = assert(read:sheet("Présentation 1.2"))
  check(rs:read(0, 0) == "Titre" and rs:read(0, 1) == nil,
    "la fusion conserve seulement la cellule supérieure gauche")
  check(table.concat(rs:merged_cells(), ",") == "A1:D1", "fusions exposées en lecture")

  local formula = rs:get_formula(3, 0)
  check(xlsx.is_formula(formula) and formula.expression == "SUM(B2:B3)"
      and formula.cached_value == 19.75 and rs:read(3, 0) == 19.75,
    "formule et valeur mise en cache exposées séparément")

  local read_style = rs:get_style(0, 0)
  check(read_style and read_style.font_name == "Liberation Sans"
      and read_style.font_size == 14 and read_style.underline == "double"
      and read_style.strike == true, "police enrichie relue")
  check(read_style and read_style.border.left.style == "thin"
      and read_style.border.left.color == "FFFF0000"
      and read_style.border.right.style == "double", "bordures relues")

  local frozen_rows, frozen_cols = rs:get_frozen_panes()
  check(frozen_rows == 1 and frozen_cols == 2, "volets figés relus")
  check(rs:get_column_width(0) == 22, "largeur de colonne relue")
  check(rs:get_row_height(0) == 30, "hauteur de ligne relue")
  check(rs:get_auto_filter() == "A2:D4", "filtre automatique relu")

  local external = rs:get_hyperlink(1, 0)
  check(external and external.target == "https://example.com/?a=1&b=2"
      and external.internal == false and external.tooltip == "Ouvrir le site",
    "hyperlien externe relu")
  local internal = rs:get_hyperlink(2, 0)
  check(internal and internal.internal == true and internal.target == "'Cible'!A1",
    "hyperlien interne relu")

  local temp = xlsx.new():add_sheet("x")
  temp:merge_cells(0, 0, 0, 1)
  temp:unmerge_cells("A1:B1")
  check(pcall(function() temp:merge_cells("A1:B1") end), "fusion après unmerge possible")
  expect_error(function() temp:merge_cells("B1:C1") end, "chevauchement de fusions refusé")
  expect_error(function() xlsx.new():add_sheet("x"):merge_cells("A1:A1") end,
    "fusion d'une seule cellule refusée")
  expect_error(function() xlsx.style({ font_size = 0 }) end, "taille de police invalide refusée")
  expect_error(function()
    xlsx.style({ border = { left = { style = "zigzag" } } })
  end, "style de bordure inconnu refusé")
  expect_error(function() xlsx.hyperlink("", "vide") end, "cible d'hyperlien vide refusée")
  expect_error(function() xlsx.formula("A1", {}) end, "cache de formule invalide refusé")
end


-- ---- VALIDATIONS, MISE EN FORME CONDITIONNELLE ET MÉTADONNÉES 1.3 ---------
print("\n[9] validations, mise en forme conditionnelle, commentaires et propriétés")
do
  local wb = xlsx.new()
  local sh = wb:add_sheet("Saisie")
  sh:write_rows({
    { "Statut", "Score", "Note", "Technique" },
    { "Oui", 42, "À vérifier", "masqué" },
  })
  sh:add_data_validation("A2:A100", {
    type = "list", values = { "Oui", "Non", "En attente" },
    allow_blank = true, show_input_message = true, show_error_message = true,
    prompt_title = "Choix", prompt = "Sélectionner un statut",
    error_title = "Erreur", error = "Valeur invalide", error_style = "stop",
  })
  sh:add_data_validation("B2:B100", {
    type = "whole", operator = "between", minimum = 0, maximum = 100,
  })
  sh:add_data_validation("C2:C100", {
    type = "text_length", operator = "less_or_equal", value = 20,
  })
  local negative = xlsx.style({ font_color = "9C0006", fill_color = "FFC7CE" })
  local long_note = xlsx.style({ fill_color = "FFF2CC" })
  sh:add_conditional_format("B2:B100", {
    type = "cell", operator = "less_than", value = 0, style = negative,
  })
  sh:add_conditional_format("C2:C100", {
    type = "custom", formula = "LEN(C2)>20", style = long_note, stop_if_true = true,
  })
  sh:set_comment(1, 2, { author = "Julien", text = "Valeur à vérifier avant publication." })
  sh:set_row_hidden(4, true)
  sh:set_column_hidden(3, true)
  sh:set_tab_color("4472C4")

  local hidden = wb:add_sheet("Masquée")
  hidden:set_visibility("hidden")
  local very_hidden = wb:add_sheet("Très masquée")
  very_hidden:set_visibility("very_hidden")
  wb:set_active_sheet("Saisie")
  wb:define_name("TauxTVA", "'Saisie'!$B$2", { comment = "Taux global" })
  wb:define_name("ZoneStatut", "'Saisie'!$A$2:$A$100")
  wb:define_name("LocalScore", "'Saisie'!$B$2", { local_sheet = "Saisie", hidden = true })
  assert(wb:save(TMP2, { use_babet = false }))

  local read = assert(xlsx.open(TMP2, { use_babet = false }))
  check(read:get_active_sheet() == "Saisie", "feuille active relue")
  local names = read:get_defined_names()
  check(#names == 3 and names[1].name == "TauxTVA" and names[1].comment == "Taux global",
    "plages nommées relues")
  check(read:get_defined_name("LocalScore", 1).hidden == true, "nom local masqué relu")
  local rs = assert(read:sheet("Saisie"))
  check(rs:get_tab_color() == "FF4472C4", "couleur d'onglet relue")
  check(rs:is_row_hidden(4) and rs:is_column_hidden(3), "ligne et colonne masquées relues")
  check(assert(read:sheet("Masquée")):get_visibility() == "hidden"
      and assert(read:sheet("Très masquée")):get_visibility() == "very_hidden",
    "visibilités de feuilles relues")
  local comment = rs:get_comment(1, 2)
  check(comment and comment.author == "Julien" and comment.text:match("vérifier"), "commentaire relu")
  check(#rs:get_comments() == 1, "liste des commentaires exposée")
  local validations = rs:get_data_validations()
  check(#validations == 3 and validations[1].type == "list"
      and table.concat(validations[1].values, ",") == "Oui,Non,En attente",
    "validation de liste relue")
  check(validations[2].type == "whole" and validations[2].operator == "between"
      and validations[2].formula1 == "0" and validations[2].formula2 == "100",
    "validation numérique relue")
  local formats = rs:get_conditional_formats()
  check(#formats == 2 and formats[1].type == "cell" and formats[1].operator == "less_than"
      and formats[1].style.fill_color == "FFFFC7CE", "format conditionnel cell relu")
  check(formats[2].type == "custom" and formats[2].stop_if_true == true
      and formats[2].formula1 == "LEN(C2)>20", "format conditionnel custom relu")

  local temp = xlsx.new()
  local ts = temp:add_sheet("x")
  ts:add_data_validation("A1", { type = "list", values = { "Oui", "Non" } })
  ts:remove_data_validation("A1")
  check(#ts:get_data_validations() == 0, "suppression d'une validation")
  ts:add_conditional_format("A1", { type = "blanks", style = xlsx.style({ fill_color = "FFFF00" }) })
  ts:remove_conditional_format("A1")
  check(#ts:get_conditional_formats() == 0, "suppression d'un format conditionnel")
  ts:set_comment(0, 0, { author = "A", text = "B" }):remove_comment(0, 0)
  check(true, "suppression d'un commentaire")

  expect_error(function()
    xlsx.new():add_sheet("x"):add_data_validation("A1", { type = "list", values = { "a,b" } })
  end, "choix inline ambigu refusé")
  expect_error(function()
    xlsx.new():add_sheet("x"):add_data_validation("A1", { type = "whole", operator = "between", minimum = 0 })
  end, "validation between incomplète refusée")
  expect_error(function()
    xlsx.new():add_sheet("x"):add_conditional_format("A1", { type = "cell", operator = "less_than", value = 0, style = {} })
  end, "style conditionnel brut refusé")
  expect_error(function()
    xlsx.new():add_sheet("x"):add_conditional_format("A1", {
      type = "cell", operator = "less_than", value = 0,
      style = xlsx.style({ number_format = "0.00" }),
    })
  end, "format numérique conditionnel non pris en charge refusé")
  expect_error(function()
    xlsx.new():add_sheet("x"):set_comment(0, 0, { author = "", text = "x" })
  end, "auteur de commentaire vide refusé")
  expect_error(function() xlsx.new():define_name("A1", "Feuil1!A1") end,
    "nom défini ressemblant à une cellule refusé")
  do
    local bad = xlsx.new(); bad:add_sheet("x"):set_visibility("hidden")
    local ok, err = bad:save(TMP2, { use_babet = false })
    check(ok == nil and type(err) == "string", "classeur sans feuille visible refusé")
  end
end



-- ---- IMAGES, GRAPHIQUES, TABLEAUX ET IMPRESSION 1.4 ------------------------
print("\n[10] images, graphiques, tableaux et mise en page d'impression")
do
  local function from_hex(hex)
    return (hex:gsub("%x%x", function(pair) return string.char(tonumber(pair, 16)) end))
  end
  local png = from_hex("89504e470d0a1a0a0000000d49484452000000040000000308020000003b9639910000001049444154789c63fccf80004c0cb8380026760105e4f3b8d50000000049454e44ae426082")
  local wb = xlsx.new()
  local sh = wb:add_sheet("Rapport")
  sh:write_rows({
    { "Mois", "Ventes", "Objectif" },
    { "Janvier", 12, 10 },
    { "Février", 18, 15 },
    { "Mars", 16, 17 },
  })
  sh:add_table("A1:C4", { name = "VentesTable", style = "TableStyleMedium4", show_row_stripes = true })
  sh:add_image_data(png, "png", 5, 0, { width = 80, alt_text = "Carré rouge", name = "Logo" })
  sh:add_chart({
    type = "column", title = "Ventes mensuelles", categories = "A2:A4",
    series = {
      { name_ref = "B1", values = "B2:B4" },
      { name = "Objectif", values = "C2:C4" },
    },
    row = 0, col = 4, width = 520, height = 300,
  })
  sh:set_page_setup({ orientation = "landscape", paper_size = "a4", fit_to_width = 1, fit_to_height = 0, horizontal_centered = true, grid_lines = true })
  sh:set_page_margins({ left = 0.4, right = 0.4, top = 0.5, bottom = 0.5 })
  sh:set_header_footer({ header_left = "lua-xlsx", header_center = "Rapport", footer_right = "Page &P / &N" })
  sh:set_print_area("A1:H20")
  sh:set_print_titles({ rows = "1:1", columns = "A:A" })
  assert(wb:save(TMP2, { use_babet = false }))

  local read = assert(xlsx.open(TMP2, { use_babet = false }))
  local rs = assert(read:sheet("Rapport"))
  local images = rs:get_images()
  check(#images == 1 and images[1].format == "png" and images[1].row == 5 and images[1].col == 0
      and images[1].alt_text == "Carré rouge" and #images[1].data == #png, "image PNG relue")
  local charts = rs:get_charts()
  check(#charts == 1 and charts[1].type == "column" and charts[1].title == "Ventes mensuelles"
      and #charts[1].series == 2 and charts[1].series[1].name_ref:match("B%$1"), "graphique et séries relus")
  local tables = rs:get_tables()
  check(#tables == 1 and tables[1].name == "VentesTable" and tables[1].ref == "A1:C4"
      and tables[1].columns[2] == "Ventes" and tables[1].show_row_stripes == true, "tableau structuré relu")
  local setup = rs:get_page_setup()
  check(setup and setup.orientation == "landscape" and setup.paper_size == 9
      and setup.fit_to_width == 1 and setup.fit_to_height == 0 and setup.horizontal_centered == true,
    "mise en page relue")
  local margins = rs:get_page_margins()
  check(margins and margins.left == 0.4 and margins.top == 0.5, "marges relues")
  local hf = rs:get_header_footer()
  check(hf and hf.header_left == "lua-xlsx" and hf.header_center == "Rapport"
      and hf.footer_right == "Page &P / &N", "en-tête et pied de page relus")
  local rows, cols = rs:get_print_titles()
  check(rs:get_print_area() == "A1:H20" and rows == "1:1" and cols == "A:A",
    "zone et titres d'impression relus")

  expect_error(function() xlsx.new():add_sheet("x"):add_image_data("bad", "png", 0, 0) end,
    "image invalide refusée")
  expect_error(function()
    local b=xlsx.new(); local x=b:add_sheet("x"); x:write_rows({{"A","B"},{1,2}}); x:add_chart({type="radar",categories="A2:A2",series={{values="B2:B2"}}})
  end, "type de graphique non pris en charge refusé")
  expect_error(function()
    local b=xlsx.new(); local x=b:add_sheet("x"); x:write_rows({{"A","A"},{1,2}}); x:add_table("A1:B2",{name="T"})
  end, "en-têtes de tableau dupliqués refusés")
  expect_error(function() xlsx.new():add_sheet("x"):set_page_setup({ scale=100, fit_to_width=1 }) end,
    "scale et ajustement simultanés refusés")
  expect_error(function() xlsx.new():add_sheet("x"):set_print_titles({ rows="2:1" }) end,
    "titres d'impression inversés refusés")
end

-- ---- FINITION ET VISUALISATION 1.5 ----------------------------------------
print("\n[11] graphiques avancés, visualisation, protections et métadonnées")
do
  local alert = xlsx.style({ bold=true, font_color="9C0006", fill_color="FFC7CE" })
  local input = xlsx.style({ fill_color="FFF2CC", locked=false })
  local hidden_formula = xlsx.style({ hidden=true })

  local wb = xlsx.new()
  wb:set_properties({
    title="Rapport 1.5", subject="Validation des nouveautés", creator="lua-xlsx",
    description="Classeur de test", keywords={"lua", "xlsx", "test"},
    category="Tests", company="lua-xlsx", manager="Julien",
  })
  wb:protect({ password="secret", structure=true })

  local sh = wb:add_sheet("Synthèse")
  sh:write_rows({
    {"Mois", "Ventes", "Objectif", "Écart", "Tendance"},
    {"Janvier", 12, 10, 2}, {"Février", 18, 15, 3},
    {"Mars", 16, 17, -1}, {"Avril", 24, 20, 4},
  })
  sh:write(1, 1, 12, input)
  sh:write(1, 3, xlsx.formula("=B2-C2", 2), hidden_formula)
  sh:write_rich_text(6, 0, {
    { text="Attention : ", bold=true, font_color="FF0000" },
    { text="objectif non atteint", italic=true },
  })
  sh:protect({ password="feuille", select_unlocked_cells=true, auto_filter=true })
  sh:add_row_page_break(4)
  sh:add_column_page_break("D")
  sh:add_sparkline("E2", "B2:D2", { type="line", color="4472C4", show_markers=true, show_high=true, show_low=true })
  sh:add_sparkline("E3", "B3:D3", { type="column", color="70AD47", show_negative=true, negative_color="C00000" })
  sh:add_sparkline("E4", "B4:D4", { type="win_loss", color="5B9BD5" })

  sh:add_conditional_format("B2:B5", {
    type="color_scale", min_color="F8696B", mid_type="percentile", mid_value=50,
    mid_color="FFEB84", max_color="63BE7B",
  })
  sh:add_conditional_format("C2:C5", { type="data_bar", color="5B9BD5", show_value=true })
  sh:add_conditional_format("D2:D5", {
    type="icon_set", icons="3_traffic_lights", value_type="number", values={-1, 0, 3}, reverse=true,
  })
  sh:add_conditional_format("B2:B5", { type="top", rank=2, style=alert })
  sh:add_conditional_format("D2:D5", { type="below_average", style=alert })

  sh:add_chart({
    type="pie", title="Répartition", categories="A2:A5",
    series={{ name="Ventes", values="B2:B5", color="4472C4" }},
    legend_position="bottom", data_labels={show_percent=true, position="best_fit"}, row=0, col=6,
  })
  sh:add_chart({
    type="doughnut", title="Objectifs", categories="A2:A5",
    series={{ name="Objectif", values="C2:C5" }}, hole_size=65, row=18, col=6,
  })
  sh:add_chart({
    type="area", title="Évolution", categories="A2:A5", grouping="stacked",
    series={{name="Ventes",values="B2:B5",color="4472C4"},{name="Objectif",values="C2:C5",color="ED7D31"}},
    legend_position="top", x_axis={title="Mois"}, y_axis={title="Valeur",min=0,max=30,number_format="0",major_gridlines=true},
    row=36, col=6,
  })
  sh:add_chart({
    type="scatter", title="Ventes / objectif", x_values="C2:C5",
    series={{name="Ventes",values="B2:B5",color="70AD47",marker="circle",line_width=2,smooth=true}},
    legend=false, x_axis={title="Objectif",min=0,max=30}, y_axis={title="Ventes",min=0,max=30},
    row=54, col=6,
  })

  assert(wb:save(TMP2, { use_babet=false }))
  local read = assert(xlsx.open(TMP2, { use_babet=false }))
  local props = read:get_properties()
  check(props.title == "Rapport 1.5" and props.creator == "lua-xlsx"
      and props.keywords == "lua, xlsx, test" and props.company == "lua-xlsx",
    "propriétés du document relues")
  local wp = read:get_protection()
  check(wp and wp.structure == true and type(wp.password_hash) == "string",
    "protection du classeur relue")

  local rs = assert(read:sheet("Synthèse"))
  local sp = rs:get_protection()
  check(sp and sp.select_unlocked_cells == true and sp.auto_filter == true,
    "protection de la feuille relue")
  local input_opts = xlsx.style_options(assert(rs:get_style(1, 1)))
  local formula_opts = xlsx.style_options(assert(rs:get_style(1, 3)))
  check(input_opts.locked == false and formula_opts.hidden == true,
    "protection des cellules relue")

  local rich = rs:get_rich_text(6, 0)
  local runs = rich and xlsx.rich_text_runs(rich) or {}
  check(rs:read(6, 0) == "Attention : objectif non atteint" and xlsx.is_rich_text(rich)
      and #runs == 2 and runs[1].bold == true and runs[1].font_color == "FFFF0000"
      and runs[2].italic == true, "texte enrichi relu avec ses segments")

  local rb, cb = rs:get_row_page_breaks(), rs:get_column_page_breaks()
  check(#rb == 1 and rb[1] == 4 and #cb == 1 and cb[1] == 4,
    "sauts de page relus")
  local sparks = rs:get_sparklines()
  check(#sparks == 3 and sparks[1].type == "line" and sparks[1].target == "E2"
      and sparks[2].type == "column" and sparks[3].type == "win_loss",
    "sparklines relues")

  local cfs = rs:get_conditional_formats()
  local found = {}
  for _, cf in ipairs(cfs) do found[cf.type] = cf end
  check(found.color_scale and found.color_scale.mid and found.color_scale.mid.value == 50
      and found.data_bar and found.data_bar.color == "FF5B9BD5"
      and found.icon_set and found.icon_set.icons == "3_traffic_lights"
      and found.top and found.top.rank == 2 and found.below_average,
    "formats conditionnels avancés relus")

  local charts = rs:get_charts()
  local by_type = {}; for _, chart in ipairs(charts) do by_type[chart.type] = chart end
  check(#charts == 4 and by_type.pie and by_type.pie.legend_position == "bottom"
      and by_type.pie.data_labels and by_type.pie.data_labels.show_percent == true
      and by_type.doughnut and by_type.doughnut.hole_size == 65
      and by_type.area and by_type.area.grouping == "stacked"
      and by_type.area.x_axis.title == "Mois" and by_type.area.y_axis.max == 30
      and by_type.scatter and by_type.scatter.series[1].x_values ~= nil
      and by_type.scatter.series[1].line_width == 2 and by_type.scatter.series[1].smooth == true,
    "graphiques avancés et options relus")

  expect_error(function() xlsx.rich_text({{text=""}}) end,
    "segment enrichi vide refusé")
  expect_error(function() xlsx.rich_text({{text="x", shadow=true}}) end,
    "option de texte enrichi inconnue refusée")
  expect_error(function() xlsx.rich_text_runs("texte") end,
    "rich_text_runs refuse une valeur ordinaire")
  expect_error(function() xlsx.new():protect({password="1234567890123456"}) end,
    "mot de passe classique trop long refusé")
  expect_error(function() xlsx.new():add_sheet("x"):add_sparkline("A1", "B1:C1", {type="pie"}) end,
    "type de sparkline inconnu refusé")
  expect_error(function()
    xlsx.new():add_sheet("x"):add_conditional_format("A1:A3", {
      type="icon_set", icons="3_traffic_lights", values={0, 50},
    })
  end, "nombre de seuils d'icônes invalide refusé")
  expect_error(function()
    local b=xlsx.new(); local x=b:add_sheet("x"); x:write_rows({{"A","B","C"},{1,2,3}})
    x:add_chart({type="line",categories="A2:A2",series={{values="B2:B2"}},y_axis={min=10,max=5}})
  end, "bornes d'axe inversées refusées")
  expect_error(function()
    local b=xlsx.new(); local x=b:add_sheet("x"); x:write_rows({{"A","B","C"},{1,2,3}})
    x:add_chart({type="pie",categories="A2:A2",series={{values="B2:B2"},{values="C2:C2"}}})
  end, "graphique secteur à plusieurs séries refusé")
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
