# lua-xlsx

Bibliothèque Lua pour **lire et écrire des fichiers XLSX**, accompagnée d'une
petite couche **DataFrame**. Le cœur reste en Lua pur et fonctionne avec Lua
5.3+, tandis que [Babet](https://github.com/Chipsterjulien/babet) et Lua 5.5
constituent la cible principale et recommandée.

- Documentation française : [`doc/documentation-fr.md`](doc/documentation-fr.md)
- English documentation: [`doc/documentation-en.md`](doc/documentation-en.md)
- PDF : exécuter [`doc/build_doc.sh`](doc/build_doc.sh)

## Points forts

- Écriture de classeurs XLSX : chaînes, nombres, booléens, dates, dates-heures,
  formules, plusieurs feuilles et cellules éparses.
- Styles : gras, italique, soulignement, texte barré, police, taille,
  couleurs, fond, alignement, retour à la ligne, bordures et formats numériques.
- Structure et mise en page : cellules fusionnées, largeurs de colonnes,
  hauteurs de lignes, volets figés, filtres automatiques, lignes et colonnes
  masquées, couleur d'onglet et visibilité des feuilles.
- Validations de données : listes, nombres, dates, heures, longueurs de texte
  et formules personnalisées.
- Mise en forme conditionnelle : règles classiques, échelles de couleurs,
  barres de données, jeux d’icônes, classements et comparaison à la moyenne.
- Commentaires de cellules, hyperliens externes et internes, feuille active
  et plages nommées globales ou locales.
- Rapports : images PNG/JPEG, graphiques en courbes, colonnes, barres, secteurs,
  anneaux, aires et nuages de points, tableaux structurés, sparklines, zones
  d’impression, titres répétés, sauts de page, marges, en-têtes et pieds de page.
- Finition : texte enrichi dans les cellules, protection des cellules, feuilles
  et structure du classeur, ainsi que propriétés principales du document.
- Lecture des formules, styles, dimensions, fusions, volets, filtres, liens,
  validations, règles conditionnelles, commentaires, propriétés de feuilles,
  texte enrichi, protections, sparklines, images, graphiques, tableaux,
  propriétés du document et paramètres d’impression.
- Lecture des ZIP XLSX STORED ou DEFLATE avec un décompresseur DEFLATE Lua.
- Contrôle du CRC-32, des tailles, des offsets, des doublons, du chiffrement,
  des chevauchements et des rapports de compression.
- Systèmes de dates Excel **1900 et 1904** en lecture et en écriture.
- Validation UTF-8 et XML 1.0 des textes, noms de feuilles Unicode et limites
  Excel officielles.
- DataFrame non destructif : `filter`, `select`, `rename`, `mutate`, `sort`,
  `head`, `tail`, `groupby` et agrégations.
- Intégration automatique à Babet lorsqu'il est présent :
  `babet.writeFileAtomic`, `babet.archive.test`, `babet.fileSize` et
  `babet.crc32`.
- Aucun module C obligatoire et aucune commande `unzip` externe.

## Prérequis

- Cible recommandée : **Babet 2.9.0+**, donc Lua 5.5.
- Compatibilité conservée : Lua 5.3 et 5.4 avec `string.pack`, les opérateurs
  bit-à-bit, `math.type`, `utf8` et la bibliothèque `io`.
- Lua 5.1 et 5.2 ne sont pas pris en charge.

## Installation

Copier `xlsx.lua` et `dataframe.lua` à côté du script ou dans `package.path` :

```lua
local xlsx = require("xlsx")
local DF = require("dataframe")
```

Sous Babet en mode dossier, placer les modules à côté de `main.lua` suffit.

## Écriture rapide

```lua
local xlsx = require("xlsx")

local wb = xlsx.new()
local sh = wb:add_sheet("Rapport")

local header = xlsx.style({
  bold = true,
  font_name = "Liberation Sans",
  font_size = 14,
  font_color = "FFFFFF",
  fill_color = "4472C4",
  horizontal = "center",
  border = {
    bottom = { style = "double", color = "1F4E78" },
  },
})
local money = xlsx.style({ number_format = "currency_eur" })

sh:write(0, 0, "Rapport", header)
sh:merge_cells("A1:D1")
sh:append_row({ "Nom", "Montant", "Date", "Lien" }, header)
sh:append_row({ "Alice", 12.5, xlsx.date(2026, 7, 18) })
sh:write(2, 1, 12.5, money)
sh:write(2, 3, xlsx.hyperlink("https://example.com", "Ouvrir"))
sh:write(3, 1, xlsx.formula("SUM(B3:B3)", 12.5), money)
sh:set_column_width(0, 18)
sh:set_row_height(0, 28)
sh:freeze_panes(2, 1)
sh:set_auto_filter("A2:D3")
sh:add_data_validation("A3:A100", {
  type = "list",
  values = { "Oui", "Non", "En attente" },
})
sh:add_conditional_format("B3:B100", {
  type = "cell",
  operator = "less_than",
  value = 0,
  style = xlsx.style({ font_color = "9C0006", fill_color = "FFC7CE" }),
})
sh:set_comment(2, 0, { author = "Julien", text = "Valeur contrôlée." })
sh:set_column_hidden(4, true)
sh:set_tab_color("4472C4")
sh:write_rich_text(4, 0, {
  { text = "Attention : ", bold = true, font_color = "FF0000" },
  { text = "valeur à vérifier", italic = true },
})
sh:add_conditional_format("B3:B100", {
  type = "color_scale",
  min_color = "F8696B", mid_type = "percentile", mid_value = 50,
  mid_color = "FFEB84", max_color = "63BE7B",
})
sh:add_sparkline("E3", "B3:D3", {
  type = "line", show_high = true, show_low = true,
})
sh:protect({ password = "lecture", select_unlocked_cells = true })
sh:add_row_page_break(40)
sh:add_table("A2:D3", { name = "RapportTable", style = "TableStyleMedium2" })
sh:add_chart({
  type = "doughnut", title = "Montants", categories = "A3:A3",
  series = { { name_ref = "B2", values = "B3:B3", color = "4472C4" } },
  hole_size = 60, legend_position = "bottom",
  data_labels = { show_percent = true },
  row = 1, col = 5,
})
sh:set_page_setup({
  orientation = "landscape", paper_size = "a4",
  fit_to_width = 1, fit_to_height = 0,
})
sh:set_header_footer({ header_center = "Rapport", footer_right = "Page &P / &N" })
sh:set_print_area("A1:J30")
sh:set_print_titles({ rows = "1:2", columns = "A:A" })

wb:define_name("ZoneMontants", "'Rapport'!$B$3:$B$100")
wb:set_properties({
  title = "Rapport", creator = "lua-xlsx", company = "Mon entreprise",
})
wb:protect({ password = "structure", structure = true })
wb:set_active_sheet("Rapport")

local ok, err = wb:save("rapport.xlsx")
assert(ok, err)
```

Sous Babet, `save()` utilise automatiquement une publication atomique et
durable. Sous Lua standard, un temporaire est écrit puis renommé dans le même
dossier.

## Lecture rapide

```lua
local xlsx = require("xlsx")

local wb, err = xlsx.open("rapport.xlsx")
assert(wb, err)

print("Système de dates :", wb:date_system())
local sh = assert(wb:sheet("Rapport"))

for row in sh:rows() do
  for col = 1, row.n do
    io.write(tostring(row[col]), "\t")
  end
  io.write("\n")
end

local formula = sh:get_formula(3, 1)
if formula then
  print(formula.expression, formula.cached_value)
end

local style = sh:get_style(0, 0)
local frozen_rows, frozen_cols = sh:get_frozen_panes()
print(style and style.font_name, frozen_rows, frozen_cols)
print(table.concat(sh:merged_cells(), ", "))
print(sh:get_tab_color(), sh:get_visibility())
print(#sh:get_data_validations(), #sh:get_conditional_formats())
local comment = sh:get_comment(2, 0)
print(comment and comment.author, comment and comment.text)
print(wb:get_active_sheet(), #wb:get_defined_names())
print(#sh:get_images(), #sh:get_charts(), #sh:get_tables())
print(#sh:get_sparklines(), #sh:get_row_page_breaks())
local rich = sh:get_rich_text(4, 0)
if rich then
  for _, run in ipairs(xlsx.rich_text_runs(rich)) do
    print(run.text, run.bold, run.italic)
  end
end
print(wb:get_properties().title, wb:get_protection() ~= nil)
local page = sh:get_page_setup()
local repeat_rows, repeat_cols = sh:get_print_titles()
print(page and page.orientation, sh:get_print_area(), repeat_rows, repeat_cols)

```

Les dates sont renvoyées sous forme ISO 8601, par exemple `2026-07-18` ou
`2026-07-18T13:45:30`. `read()` conserve la valeur mise en cache d'une formule ;
`get_formula()` expose séparément son expression.

## DataFrame

```lua
local xlsx = require("xlsx")
local DF = require("dataframe")

local wb = assert(xlsx.open("ventes.xlsx"))
local data = DF.from_sheet(assert(wb:sheet("ventes")), { header = true })

local résumé = data
  :mutate("total", function(row)
    return row.quantité * row.prix
  end)
  :filter(function(row)
    return row.total >= 30
  end)
  :groupby("ville")
  :agg({
    total = { "sum", "total" },
    moyenne = { "mean", "total" },
    nombre = { "count" },
  })
  :sort("total", { desc = true })

assert(xlsx.write_rows("résumé.xlsx", résumé:to_rows(), {
  sheet = "Résumé",
}))
```

Les transformations copient les records qu'elles retournent. Modifier une
ligne du résultat ne modifie donc pas le DataFrame source.

## Limites de sécurité configurables

```lua
local wb, err = xlsx.open("import.xlsx", {
  max_file_size = 256 * 1024 * 1024,
  max_entries = 10000,
  max_entry_size = 64 * 1024 * 1024,
  max_total_size = 512 * 1024 * 1024,
  max_path_length = 4096,
  max_total_name_bytes = 16 * 1024 * 1024,
  max_compression_ratio = 200,
})
assert(wb, err)
```

Avec Babet, ces limites sont également transmises à `babet.archive.test()`
avant le parseur Lua.

## Tests

```sh
./run_tests.sh
```

Le script :

1. cherche Babet dans `BABET_BIN`, puis dans `bin/babet`, puis dans le `PATH` ;
2. teste aussi un interpréteur Lua standard trouvé sur le système ;
3. crée un environnement virtuel Python temporaire ;
4. y installe `openpyxl>=3.1,<4` et `Pillow>=10,<12` avec `pip` ;
5. exécute l'aller-retour `openpyxl` séparément avec Babet et Lua standard lorsqu'ils sont disponibles ;
6. supprime le venv, le cache `pip` et tous les fichiers temporaires à la fin, y compris après un échec.

Le test d'interopérabilité n'est plus ignoré lorsque `openpyxl` n'est pas installé
globalement. Il faut seulement disposer de `python3`, du module `venv` et d'un
accès au dépôt Python configuré pour `pip`. Sous Debian ou Ubuntu, le paquet
`python3-venv` peut être nécessaire.

Pour utiliser le binaire local recommandé :

```sh
cp /chemin/vers/babet bin/babet
chmod +x bin/babet
./run_tests.sh
```

Variables utiles :

```sh
BABET_BIN=/chemin/vers/babet ./run_tests.sh
LUA_BIN=/chemin/vers/lua5.5 ./run_tests.sh
OPENPYXL_SPEC='openpyxl==3.1.5' PILLOW_SPEC='pillow==11.3.0' ./run_tests.sh
```

## Documentation PDF

```sh
cd doc
./build_doc.sh
```

Le script produit `documentation-fr.pdf` et `documentation-en.pdf` à partir des
fichiers Markdown. Il nécessite Pandoc et XeLaTeX ou LuaLaTeX.

## Limites actuelles

- Pas de tableaux croisés dynamiques, macros XLSM, graphiques combinés ou 3D,
  axe secondaire, images SVG ou mise en page avancée par zones multiples.
- Les sparklines sont écrites dans l’extension OOXML standard. `openpyxl 3.1.5`
  peut lire le classeur, mais ne conserve pas cette extension s’il le réenregistre.
- La protection XLSX classique limite les mots de passe à 15 octets et ne
  constitue pas un chiffrement du contenu.
- Les commentaires sont lus et écrits comme du texte simple : les fragments
  riches, dimensions et positions personnalisées ne sont pas conservés.
- Les formules sont lues et écrites, mais lua-xlsx ne les calcule pas. Une
  formule partagée sans expression résolue peut être inspectée mais pas
  réécrite directement.
- La lecture des styles expose le sous-ensemble pris en charge par lua-xlsx ;
  les thèmes, couleurs indexées, diagonales et variantes avancées ne sont pas
  reproduits.
- La lecture reste en mémoire : le ZIP, les XML utiles et les résultats
  décompressés sont chargés en RAM dans les limites configurées.
- Le lecteur Lua ne prend pas encore en charge ZIP64. Babet peut valider un ZIP64,
  mais le parseur interne le refuse ensuite explicitement.
- Le parseur XML est spécialisé XLSX et ne remplace pas un parseur XML général.

## Licence

GNU General Public License v3.0 — voir [`LICENSE`](LICENSE).
