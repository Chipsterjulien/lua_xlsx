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

--- dataframe.lua — manipulation tabulaire type "pandas", en Lua pur.
--- Cible : LuaPilot (Lua 5.5). Aucune dépendance (ni C, ni SQLite).
--- S'interface avec le module xlsx : xlsx (lecture) -> dataframe -> xlsx (écriture).
---
--- Modèle : un DataFrame porte `columns` (liste ordonnée de noms) et `rows`
--- (tableau de records, chaque record étant une table nom -> valeur).
--- Les opérations sont NON destructives : elles renvoient un nouveau DataFrame.

local df = {}

local DataFrame = {}
DataFrame.__index = DataFrame

local function make(columns, rows)
  return setmetatable({ columns = columns, rows = rows }, DataFrame)
end

-- ----------------------------------------------------------------------------
-- Construction
-- ----------------------------------------------------------------------------

--- À partir d'une matrice (tableau de lignes). Chaque ligne est un tableau
--- 1-indexé (row[1] = 1re colonne ; .n optionnel, comme les lignes du lecteur
--- xlsx). opts.header = true : la 1re ligne donne les noms de colonnes.
--- opts.columns : noms explicites (sinon "col1".."colN").
function df.from_rows(matrix, opts)
  opts = opts or {}
  if type(matrix) ~= "table" then
    error("dataframe: from_rows attend une table de lignes", 2)
  end
  local start = 1
  local cols = opts.columns
  if opts.header then
    local hdr = matrix[1]
    if not hdr then error("dataframe: header demandé mais matrice vide", 2) end
    cols = {}
    local w = hdr.n or #hdr
    for i = 1, w do
      local h = hdr[i]
      cols[i] = (h == nil) and ("col" .. i) or tostring(h)
    end
    start = 2
  end
  if not cols then
    local first = matrix[start] or {}
    local w = first.n or #first
    cols = {}
    for i = 1, w do cols[i] = "col" .. i end
  end
  local ncol = #cols
  local rows = {}
  for i = start, #matrix do
    local src = matrix[i]
    local r = {}
    for c = 1, ncol do r[cols[c]] = src[c] end
    rows[#rows + 1] = r
  end
  return make(cols, rows)
end

--- À partir d'une liste de records (tables nom -> valeur).
--- columns : ordre des colonnes (recommandé ; sinon déduit puis trié).
function df.from_records(records, columns)
  if type(records) ~= "table" then
    error("dataframe: from_records attend une table de records", 2)
  end
  local cols = columns
  if not cols then
    cols = {}
    local seen = {}
    for _, rec in ipairs(records) do
      for k in pairs(rec) do
        if not seen[k] then seen[k] = true; cols[#cols + 1] = k end
      end
    end
    table.sort(cols)
  end
  local rows = {}
  for i = 1, #records do
    local rec = records[i]
    local r = {}
    for _, c in ipairs(cols) do r[c] = rec[c] end
    rows[i] = r
  end
  return make(cols, rows)
end

--- Pratique : depuis une feuille du lecteur xlsx (utilise sheet:rows()).
function df.from_sheet(sheet, opts)
  local matrix = {}
  for row in sheet:rows() do matrix[#matrix + 1] = row end
  return df.from_rows(matrix, opts)
end

-- ----------------------------------------------------------------------------
-- Introspection
-- ----------------------------------------------------------------------------
function DataFrame:nrow() return #self.rows end
function DataFrame:ncol() return #self.columns end

function DataFrame:colnames()
  local t = {}
  for i, c in ipairs(self.columns) do t[i] = c end
  return t
end

function DataFrame:_has(name)
  if not self._set then
    local s = {}
    for _, c in ipairs(self.columns) do s[c] = true end
    self._set = s
  end
  return self._set[name] == true
end

--- Itérateur sur les records (chaque valeur produite est une table nom->valeur).
function DataFrame:iter()
  local i, rows = 0, self.rows
  return function()
    i = i + 1
    return rows[i]
  end
end

--- Valeurs d'une colonne (tableau).
function DataFrame:column(name)
  if not self:_has(name) then
    error("dataframe: colonne inconnue : " .. tostring(name), 2)
  end
  local t = {}
  for i, r in ipairs(self.rows) do t[i] = r[name] end
  return t
end

-- ----------------------------------------------------------------------------
-- Transformations (renvoient un nouveau DataFrame)
-- ----------------------------------------------------------------------------

--- Conserve les lignes pour lesquelles pred(row) est vrai.
function DataFrame:filter(pred)
  if type(pred) ~= "function" then
    error("dataframe: filter attend une fonction", 2)
  end
  local out = {}
  for _, r in ipairs(self.rows) do
    if pred(r) then out[#out + 1] = r end
  end
  return make(self:colnames(), out)
end

--- Projette un sous-ensemble de colonnes (dans l'ordre donné).
function DataFrame:select(...)
  local cols = { ... }
  if #cols == 1 and type(cols[1]) == "table" then cols = cols[1] end
  for _, c in ipairs(cols) do
    if not self:_has(c) then
      error("dataframe: colonne inconnue : " .. tostring(c), 2)
    end
  end
  local out = {}
  for i, r in ipairs(self.rows) do
    local nr = {}
    for _, c in ipairs(cols) do nr[c] = r[c] end
    out[i] = nr
  end
  return make(cols, out)
end

--- Renomme des colonnes selon map = { ancien = nouveau }.
function DataFrame:rename(map)
  if type(map) ~= "table" then error("dataframe: rename attend une table", 2) end
  local newcols = {}
  for i, c in ipairs(self.columns) do newcols[i] = map[c] or c end
  local out = {}
  for i, r in ipairs(self.rows) do
    local nr = {}
    for _, c in ipairs(self.columns) do nr[map[c] or c] = r[c] end
    out[i] = nr
  end
  return make(newcols, out)
end

--- Ajoute (ou remplace) une colonne calculée : value = fn(row).
function DataFrame:mutate(name, fn)
  if type(name) ~= "string" then
    error("dataframe: mutate attend un nom de colonne (string)", 2)
  end
  if type(fn) ~= "function" then
    error("dataframe: mutate attend une fonction", 2)
  end
  local cols = self:colnames()
  if not self:_has(name) then cols[#cols + 1] = name end
  local out = {}
  for i, r in ipairs(self.rows) do
    local nr = {}
    for _, c in ipairs(self.columns) do nr[c] = r[c] end
    nr[name] = fn(r)
    out[i] = nr
  end
  return make(cols, out)
end

local function cmp_vals(va, vb)
  if va == vb then return 0 end
  local ta, tb = type(va), type(vb)
  if ta == tb and (ta == "number" or ta == "string") then
    if va < vb then return -1 else return 1 end
  end
  local sa, sb = tostring(va), tostring(vb)
  if sa == sb then return 0 elseif sa < sb then return -1 else return 1 end
end

--- Trie par colonne. opts.desc = true pour décroissant. Tri STABLE ;
--- les valeurs nil vont toujours en dernier.
function DataFrame:sort(col, opts)
  opts = opts or {}
  if not self:_has(col) then
    error("dataframe: colonne inconnue : " .. tostring(col), 2)
  end
  local desc = opts.desc and true or false
  local rows = self.rows
  local idx = {}
  for i = 1, #rows do idx[i] = i end
  table.sort(idx, function(a, b)
    local va, vb = rows[a][col], rows[b][col]
    local an, bn = (va == nil), (vb == nil)
    if an or bn then
      if an and bn then return a < b end
      return bn -- a avant b ssi b est nil (nils en dernier)
    end
    local c = cmp_vals(va, vb)
    if desc then c = -c end
    if c == 0 then return a < b end -- stabilité
    return c < 0
  end)
  local out = {}
  for i = 1, #idx do out[i] = rows[idx[i]] end
  return make(self:colnames(), out)
end

local function slice(self, from, to)
  local out = {}
  for i = from, to do out[#out + 1] = self.rows[i] end
  return make(self:colnames(), out)
end

function DataFrame:head(n)
  n = n or 5
  return slice(self, 1, math.min(n, #self.rows))
end

function DataFrame:tail(n)
  n = n or 5
  return slice(self, math.max(1, #self.rows - n + 1), #self.rows)
end

-- ----------------------------------------------------------------------------
-- GroupBy + agrégations
-- ----------------------------------------------------------------------------
local AGG = {
  count = function(rows, col)
    if not col then return #rows end
    local n = 0
    for _, r in ipairs(rows) do if r[col] ~= nil then n = n + 1 end end
    return n
  end,
  sum = function(rows, col)
    local s = 0
    for _, r in ipairs(rows) do
      local v = r[col]; if type(v) == "number" then s = s + v end
    end
    return s
  end,
  mean = function(rows, col)
    local s, n = 0, 0
    for _, r in ipairs(rows) do
      local v = r[col]; if type(v) == "number" then s = s + v; n = n + 1 end
    end
    if n == 0 then return nil end
    return s / n
  end,
  min = function(rows, col)
    local m
    for _, r in ipairs(rows) do
      local v = r[col]
      if v ~= nil and (m == nil or cmp_vals(v, m) < 0) then m = v end
    end
    return m
  end,
  max = function(rows, col)
    local m
    for _, r in ipairs(rows) do
      local v = r[col]
      if v ~= nil and (m == nil or cmp_vals(v, m) > 0) then m = v end
    end
    return m
  end,
  first = function(rows, col)
    for _, r in ipairs(rows) do if r[col] ~= nil then return r[col] end end
    return nil
  end,
  last = function(rows, col)
    local last
    for _, r in ipairs(rows) do if r[col] ~= nil then last = r[col] end end
    return last
  end,
}
AGG.avg = AGG.mean

local GroupBy = {}
GroupBy.__index = GroupBy

--- Groupe par une ou plusieurs colonnes.
function DataFrame:groupby(...)
  local keys = { ... }
  if #keys == 1 and type(keys[1]) == "table" then keys = keys[1] end
  if #keys == 0 then error("dataframe: groupby requiert au moins une colonne", 2) end
  for _, k in ipairs(keys) do
    if not self:_has(k) then
      error("dataframe: colonne de groupe inconnue : " .. tostring(k), 2)
    end
  end
  return setmetatable({ _df = self, _keys = keys }, GroupBy)
end

--- Agrège. spec = { nom_sortie = { fonction, colonne_source }, ... }
--- fonctions : sum, mean/avg, min, max, count, first, last.
--- "count" sans colonne = nombre de lignes du groupe.
--- L'ordre des colonnes de sortie : clés de groupe puis agrégats triés par nom.
function GroupBy:agg(spec)
  if type(spec) ~= "table" then error("dataframe: agg attend une table de specs", 2) end
  local d = self._df
  local keys = self._keys

  local aggnames = {}
  for name in pairs(spec) do aggnames[#aggnames + 1] = name end
  table.sort(aggnames)

  for _, name in ipairs(aggnames) do
    local e = spec[name]
    if type(e) ~= "table" or type(e[1]) ~= "string" or not AGG[e[1]] then
      error("dataframe: spec d'agrégation invalide pour '" .. tostring(name) .. "'", 2)
    end
    if e[1] ~= "count" and not e[2] then
      error("dataframe: l'agrégat '" .. e[1] .. "' requiert une colonne source", 2)
    end
    if e[2] and not d:_has(e[2]) then
      error("dataframe: colonne source inconnue : " .. tostring(e[2]), 2)
    end
  end

  local function keystr(row)
    local parts = {}
    for i, k in ipairs(keys) do
      local v = row[k]
      parts[i] = type(v) .. "\1" .. tostring(v)
    end
    return table.concat(parts, "\2")
  end

  local order, groups = {}, {}
  for _, row in ipairs(d.rows) do
    local ks = keystr(row)
    local g = groups[ks]
    if not g then
      g = { keyvals = {}, rows = {} }
      for _, k in ipairs(keys) do g.keyvals[k] = row[k] end
      groups[ks] = g
      order[#order + 1] = ks
    end
    g.rows[#g.rows + 1] = row
  end

  local outcols = {}
  for _, k in ipairs(keys) do outcols[#outcols + 1] = k end
  for _, name in ipairs(aggnames) do outcols[#outcols + 1] = name end

  local outrows = {}
  for _, ks in ipairs(order) do
    local g = groups[ks]
    local r = {}
    for _, k in ipairs(keys) do r[k] = g.keyvals[k] end
    for _, name in ipairs(aggnames) do
      local e = spec[name]
      r[name] = AGG[e[1]](g.rows, e[2])
    end
    outrows[#outrows + 1] = r
  end
  return make(outcols, outrows)
end

--- Raccourci : nombre de lignes par groupe (colonne "n").
function GroupBy:count()
  return self:agg({ n = { "count" } })
end

-- ----------------------------------------------------------------------------
-- Sorties
-- ----------------------------------------------------------------------------
function DataFrame:to_records()
  local out = {}
  for i, r in ipairs(self.rows) do
    local nr = {}
    for _, c in ipairs(self.columns) do nr[c] = r[c] end
    out[i] = nr
  end
  return out
end

--- Matrice (tableau de lignes 1-indexées) directement utilisable par
--- xlsx.write_rows. opts.header = false pour omettre la ligne d'en-tête.
function DataFrame:to_rows(opts)
  opts = opts or {}
  local out = {}
  if opts.header ~= false then
    local hdr = {}
    for i, c in ipairs(self.columns) do hdr[i] = c end
    out[#out + 1] = hdr
  end
  for _, r in ipairs(self.rows) do
    local arr = {}
    for i, c in ipairs(self.columns) do arr[i] = r[c] end
    out[#out + 1] = arr
  end
  return out
end

-- ----------------------------------------------------------------------------
-- Affichage (debug)
-- ----------------------------------------------------------------------------
local function cell_str(v)
  if v == nil then return "" end
  local t = type(v)
  if t == "number" then
    if math.type(v) == "integer" then return tostring(v) end
    return string.format("%.6g", v)
  elseif t == "boolean" then
    return v and "true" or "false"
  end
  return tostring(v)
end

--- Construit une représentation texte (tableau ASCII) des n premières lignes.
function DataFrame:tostring(n)
  n = n or 20
  local cols = self.columns
  local widths = {}
  for i, c in ipairs(cols) do widths[i] = #c end
  local shown = math.min(n, #self.rows)
  for i = 1, shown do
    local r = self.rows[i]
    for j, c in ipairs(cols) do
      local w = #cell_str(r[c])
      if w > widths[j] then widths[j] = w end
    end
  end
  local function line(vals)
    local parts = {}
    for j = 1, #cols do
      local s = vals[j]
      parts[j] = s .. string.rep(" ", widths[j] - #s)
    end
    return table.concat(parts, " | ")
  end
  local out = {}
  local hdr = {}
  for j, c in ipairs(cols) do hdr[j] = c end
  out[#out + 1] = line(hdr)
  local sep = {}
  for j = 1, #cols do sep[j] = string.rep("-", widths[j]) end
  out[#out + 1] = line(sep)
  for i = 1, shown do
    local r = self.rows[i]
    local vals = {}
    for j, c in ipairs(cols) do vals[j] = cell_str(r[c]) end
    out[#out + 1] = line(vals)
  end
  if #self.rows > shown then
    out[#out + 1] = "... (" .. (#self.rows - shown) .. " lignes de plus)"
  end
  return table.concat(out, "\n")
end

--- Affiche les n premières lignes ; renvoie self pour chaînage.
function DataFrame:show(n)
  print(self:tostring(n))
  return self
end

return df
