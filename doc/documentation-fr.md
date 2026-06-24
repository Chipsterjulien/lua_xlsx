# Documentation (Français)

Référence complète de l'API des modules `xlsx` et `dataframe`.

- [Conventions](#conventions)
- [Module `xlsx` — écriture](#module-xlsx--écriture)
- [Module `xlsx` — dates](#module-xlsx--dates)
- [Module `xlsx` — lecture](#module-xlsx--lecture)
- [Module `dataframe`](#module-dataframe)
- [Exemple complet](#exemple-complet)
- [Contrat d'erreur](#contrat-derreur)
- [Limites & internes](#limites--internes)

---

## Conventions

- **Les indices sont 0-indexés** pour `sheet:write(row, col, …)` et
  `sheet:read(row, col)` (la cellule `A1` est `(0, 0)`), afin de garder l'écriture
  et la lecture symétriques.
- L'**itérateur de lignes** `sheet:rows()` produit des tableaux **1-indexés**
  (`row[1]` = colonne A), car c'est ainsi qu'on boucle naturellement en Lua ;
  chaque `row` porte un champ `.n` indiquant le nombre de colonnes.
- Les cellules vides valent `nil`.
- À la lecture, les cellules date sont renvoyées **en chaînes ISO 8601**.
- Toutes les opérations DataFrame sont **non destructives** : elles renvoient un
  nouveau DataFrame.

---

## Module `xlsx` — écriture

### `xlsx.new() -> workbook`

Crée un classeur vide.

### `workbook:add_sheet([nom]) -> sheet`

Ajoute une feuille et la renvoie. `nom` vaut `"SheetN"` par défaut. Un nom doit
faire 1 à 31 caractères et ne pas contenir `: \ / ? * [ ]`.

### `sheet:write(row, col, value) -> sheet`

Écrit une cellule. `row` et `col` sont des entiers >= 0 (0-indexés). `value` peut
être une `string`, un `number`, un `boolean`, une **valeur date** (voir plus bas)
ou `nil` (vide). `NaN`/`Inf` sont refusés. Renvoie la feuille (chaînable).

### `sheet:append_row(values) -> sheet`

Ajoute une ligne complète après la dernière. `values` est un tableau 1-indexé ;
les `nil` de fin sont ignorés. Renvoie la feuille.

### `sheet:write_rows(matrix) -> sheet`

Ajoute chaque ligne de `matrix` (tableau de tableaux). Renvoie la feuille.

### `workbook:save(path) -> true | nil, err`

Écrit le fichier `.xlsx`. Renvoie `true`, ou `(nil, message)` en cas d'échec I/O.
Si le classeur n'a aucune feuille, une `"Sheet1"` vide est créée.

### `xlsx.write_rows(path, matrix [, opts]) -> true | nil, err`

Pratique : écrit une matrice unique dans `path`. `opts.sheet` fixe le nom de la
feuille (défaut `"Sheet1"`).

```lua
local xlsx = require("xlsx")
local wb = xlsx.new()
local sh = wb:add_sheet("Personnes")
sh:write(0, 0, "Nom"); sh:write(0, 1, "Âge")
sh:append_row({ "Alice", 30 })
sh:append_row({ "Bob", 25 })
assert(wb:save("personnes.xlsx"))
```

---

## Module `xlsx` — dates

Les dates sont écrites comme des nombres portant un style de format date. Utilise
les constructeurs ci-dessous comme valeur de cellule.

### `xlsx.date(année, mois, jour) -> valeur date`

Une date affichée `yyyy-mm-dd`. Les arguments doivent être des entiers.

### `xlsx.datetime(année, mois, jour [, heure, min, sec]) -> valeur date`

Une date-heure affichée `yyyy-mm-dd hh:mm:ss`. Les parties horaires valent 0 par
défaut.

```lua
sh:write(0, 0, xlsx.date(2024, 3, 15))
sh:write(1, 0, xlsx.datetime(2024, 3, 15, 13, 45, 30))
```

À la **lecture**, une cellule numérique dont le style est un format de date est
renvoyée en chaîne ISO 8601 : `"2024-03-15"` ou `"2024-03-15T13:45:30"`.

### Utilitaires

- `xlsx.serial_to_iso(serial, withtime) -> string` — convertit un numéro de série
  Excel en ISO 8601 (`withtime` vrai ajoute la partie horaire).
- `xlsx.date_to_serial(année, mois, jour [, heure, min, sec]) -> number` —
  l'inverse.

Les dates utilisent le système 1900 d'Excel avec l'epoch effectif 1899-12-30
(identique à openpyxl) ; la conversion est purement entière, indépendante de
`os.time`/`os.date`, donc sans souci de fuseau horaire ni de plage.

---

## Module `xlsx` — lecture

### `xlsx.open(path) -> workbook | nil, err`

Ouvre un fichier `.xlsx` en lecture. Renvoie un classeur, ou `(nil, message)` si
le fichier ne peut être ouvert ou n'est pas un `.xlsx` valide.

### `workbook:sheet_names() -> { string, … }`

Liste ordonnée des noms de feuilles.

### `workbook:sheet(which) -> sheet | nil [, err]`

Renvoie une feuille par **nom** (string) ou **index 1-based** (number), ou `nil`
si introuvable. Les feuilles sont analysées paresseusement puis mises en cache.

### `sheet:read(row, col) -> value | nil`

Lit une cellule (0-indexé). Renvoie la valeur (`string`, `number`, `boolean`, ou
une chaîne date ISO 8601), ou `nil` si vide. `false` est préservé (il ne devient
pas `nil`).

### `sheet:dims() -> maxrow, maxcol`

Renvoie les indices 0-indexés maximum de ligne et colonne, ou `-1, -1` pour une
feuille vide.

### `sheet:rows() -> itérateur`

Renvoie un itérateur produisant une `row` par ligne de tableur, de la ligne 0 à
`maxrow` (lignes vides intermédiaires incluses). Chaque `row` est un tableau
**1-indexé** (`row[1]` = colonne A) avec les cellules vides à `nil`, et `row.n`
contient le nombre de colonnes.

```lua
local wb = xlsx.open("personnes.xlsx")
local sh = wb:sheet("Personnes")
print(sh:read(1, 0))                 -- "Alice"
for row in sh:rows() do
  for c = 1, row.n do io.write(tostring(row[c]), "\t") end
  io.write("\n")
end
```

---

## Module `dataframe`

Un DataFrame porte une liste ordonnée de noms de colonnes et une liste de records
(chacun une table nom → valeur).

### Construction

#### `df.from_rows(matrix [, opts]) -> DataFrame`

Construit depuis une matrice (tableau de lignes 1-indexées, avec `.n` optionnel).
`opts.header = true` prend la 1re ligne comme noms de colonnes ; `opts.columns`
fournit des noms explicites ; sinon les colonnes sont nommées `col1`…`colN`.

#### `df.from_records(records [, columns]) -> DataFrame`

Construit depuis une liste de records (tables nom → valeur). Passer `columns` est
recommandé pour fixer l'ordre ; sinon les colonnes sont déduites puis triées.

#### `df.from_sheet(sheet [, opts]) -> DataFrame`

Pratique : consomme une feuille du lecteur `xlsx` via `sheet:rows()`. `opts`
identiques à `from_rows` (typiquement `{ header = true }`).

### Introspection

- `df:nrow() -> integer` — nombre de lignes.
- `df:ncol() -> integer` — nombre de colonnes.
- `df:colnames() -> { string, … }` — noms de colonnes ordonnés.
- `df:column(nom) -> { value, … }` — valeurs d'une colonne.
- `df:iter() -> itérateur` — itère sur les records (tables nom → valeur).

### Transformations (chacune renvoie un nouveau DataFrame)

- `df:filter(pred)` — garde les lignes où `pred(row)` est vrai.
- `df:select(...)` — projette un sous-ensemble de colonnes, dans l'ordre donné.
  Accepte plusieurs noms ou un seul tableau.
- `df:rename(map)` — renomme selon `{ ancien = nouveau }`.
- `df:mutate(nom, fn)` — ajoute ou remplace une colonne calculée : `value = fn(row)`.
- `df:sort(col [, opts])` — tri stable par `col` ; `opts.desc = true` pour
  décroissant. Les `nil` vont toujours en dernier.
- `df:head([n])` / `df:tail([n])` — `n` premières/dernières lignes (défaut 5).

### Groupement & agrégation

#### `df:groupby(...) -> GroupBy`

Groupe par une ou plusieurs colonnes (plusieurs noms ou un seul tableau).

#### `groupby:agg(spec) -> DataFrame`

`spec` vaut `{ nom_sortie = { fonction, colonne_source }, … }`. Fonctions : `sum`,
`mean` (alias `avg`), `min`, `max`, `count`, `first`, `last`. `"count"` sans
colonne compte les lignes du groupe. Les colonnes de sortie sont les clés de
groupe d'abord, puis les agrégats **triés par nom** (déterministe). L'ordre des
groupes suit la première apparition.

#### `groupby:count() -> DataFrame`

Raccourci pour le nombre de lignes par groupe, dans une colonne nommée `n`.

```lua
local g = d:groupby("ville", "produit"):agg({
  total = { "sum", "montant" },
  moy   = { "mean", "montant" },
  n     = { "count" },
})
```

### Sorties

- `df:to_records() -> { record, … }` — tableau de tables nom → valeur.
- `df:to_rows([opts]) -> matrix` — tableau de lignes 1-indexées, directement
  utilisable par `xlsx.write_rows`. Une ligne d'en-tête est incluse sauf si
  `opts.header == false`.
- `df:tostring([n]) -> string` — un tableau ASCII des `n` premières lignes
  (défaut 20).
- `df:show([n]) -> df` — affiche `tostring(n)` ; renvoie le DataFrame (chaînable).

---

## Exemple complet

```lua
local xlsx = require("xlsx")
local DF   = require("dataframe")

-- lire un vrai .xlsx (compressé en DEFLATE par Excel/LibreOffice/pandas)
local wb = assert(xlsx.open("ventes.xlsx"))
local d  = DF.from_sheet(wb:sheet("ventes"), { header = true })

-- transformer
local resume = d
  :mutate("total", function(r) return r.qte * r.prix end)
  :filter(function(r) return r.total >= 30 end)
  :groupby("ville")
  :agg({ total = { "sum", "total" }, n = { "count" } })
  :sort("total", { desc = true })

resume:show()

-- réécrire le résultat dans un nouveau classeur
assert(xlsx.write_rows("resume.xlsx", resume:to_rows(), { sheet = "resume" }))
```

---

## Contrat d'erreur

- **Mauvais types d'argument** (ex. un indice de ligne non entier, une table là
  où une string est attendue) lèvent via `error()` — ce sont des bugs côté
  appelant.
- **Échecs runtime** en lecture/écriture renvoient `(nil, "xlsx: …")` :
  fichier impossible à ouvrir, ZIP/`.xlsx` invalide, partie requise manquante,
  méthode de compression non supportée, flux DEFLATE tronqué ou corrompu.
- Les cas « l'opération a réussi mais la donnée est inhabituelle » ne sont pas des
  erreurs : une feuille vide se relit comme une feuille avec `dims() == -1, -1`.

---

## Limites & internes

**L'écriture** produit un ZIP non compressé (STORED). Cela ne nécessite aucun
compresseur et est lu correctement par Excel, LibreOffice et openpyxl. Le CRC-32
de chaque entrée est calculé en Lua pur.

**La lecture** analyse le répertoire central du ZIP et décompresse les entrées
avec un **décodeur DEFLATE complet en Lua pur** (RFC 1951 : blocs stored, Huffman
fixe et dynamique, et back-references LZ77). Le XML est analysé par des motifs
légers spécifiques au xlsx, pas par un parseur XML général.

**Hors périmètre** (additif plus tard) : styles visuels (gras, couleurs,
bordures), formules (la valeur cachée est lue, pas la formule), cellules
fusionnées, lecteur en streaming pour les très gros fichiers, et pont SQLite pour
les requêtes lourdes.

**Mémoire :** le décompresseur accumule les octets décompressés en mémoire avant
de les transformer en chaîne. C'est parfait pour des milliers à dizaines de
milliers de lignes ; pour des centaines de Mo décompressés, une fenêtre glissante
serait l'optimisation naturelle.

**Compatibilité :** les modules n'utilisent que la bibliothèque standard de Lua et
exigent Lua 5.3+ (opérateurs bit-à-bit, `string.pack`/`unpack`, `math.type`,
`utf8.char`, division entière `//`). Ils n'ont **aucune dépendance envers
LuaPilot** ; LuaPilot n'est qu'un hôte Lua 5.5 pratique.
