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

--- selftest.lua — auto-test des modules xlsx + dataframe.
--- À lancer sous LuaPilot (Lua 5.5) pour valider la compatibilité sur la cible.
---
--- Usage (mode dossier) : placer xlsx.lua, dataframe.lua et ce fichier (renommé
--- main.lua) dans un dossier, puis :  ./test/luapilot <dossier>
--- Ou directement :  luapilot selftest.lua   (selon ton mode de lancement)
---
--- N'écrit qu'un fichier temporaire dans le dossier courant, supprimé à la fin.

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

print("Lua : " .. _VERSION)
print("string.pack=" .. tostring(string.pack ~= nil)
  .. "  math.type=" .. tostring(math.type ~= nil)
  .. "  utf8.char=" .. tostring(utf8 and utf8.char ~= nil))

local TMP = "._xlsx_selftest.xlsx"

-- ---- ÉCRITURE puis LECTURE -------------------------------------------------
print("\n[1] écriture + relecture xlsx")
do
  local wb = xlsx.new()
  local s = wb:add_sheet("Données")
  s:write(0, 0, "Nom"); s:write(0, 1, "Âge"); s:write(0, 2, "Actif"); s:write(0, 3, "Note")
  s:append_row({ "Alice <x>", 30, true,  12.5 })
  s:append_row({ "Bob & Co",  25, false, -0.5 })
  s:write(4, 1, 999)            -- cellule éparse
  local s2 = wb:add_sheet("Autre")
  s2:write_rows({ { "a", "b" }, { 1, 2 } })
  assert(wb:save(TMP))

  local r = assert(xlsx.open(TMP))
  check(table.concat(r:sheet_names(), ",") == "Données,Autre", "feuilles dans l'ordre")
  local sh = r:sheet("Données")
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
print("\n[2] dataframe : pipeline complet")
do
  local r = assert(xlsx.open(TMP))
  local d = DF.from_sheet(r:sheet("Données"), { header = true })
  check(d:nrow() == 4, "from_sheet : 4 lignes (dont vides)")

  -- ne garder que les lignes avec un nom
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

  -- round-trip dataframe -> xlsx -> dataframe
  assert(xlsx.write_rows(TMP, g:to_rows(), { sheet = "g" }))
  local back = DF.from_sheet(assert(xlsx.open(TMP)):sheet("g"), { header = true })
  check(back:ncol() == 4, "round-trip : 4 colonnes (Actif + 3 agg)")
end

os.remove(TMP)

-- ---- DATES -----------------------------------------------------------------
print("\n[3] dates (styles)")
do
  local wb = xlsx.new()
  local s = wb:add_sheet("d")
  s:write(0, 0, xlsx.date(2024, 3, 15))
  s:write(1, 0, xlsx.datetime(2024, 3, 15, 13, 45, 30))
  s:write(2, 0, 7) -- nombre normal, ne doit pas devenir une date
  assert(wb:save(TMP))

  local r = assert(xlsx.open(TMP))
  local sh = r:sheet("d")
  check(sh:read(0, 0) == "2024-03-15",          "date -> ISO")
  check(sh:read(1, 0) == "2024-03-15T13:45:30", "datetime -> ISO")
  check(sh:read(2, 0) == 7,                      "nombre normal préservé")
end

os.remove(TMP)

print("\n========================================")
if fails == 0 then
  print("SELFTEST : PASS  (" .. _VERSION .. ")")
  os.exit(0)
else
  print("SELFTEST : " .. fails .. " ÉCHEC(S)")
  os.exit(1)
end
