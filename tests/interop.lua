-- Interop test executed by run_tests.sh.
local xlsx = require("xlsx")

local input = assert(os.getenv("LUA_XLSX_INPUT"), "LUA_XLSX_INPUT manquant")
local output = assert(os.getenv("LUA_XLSX_OUTPUT"), "LUA_XLSX_OUTPUT manquant")

local wb, err = xlsx.open(input)
assert(wb, err)
assert(wb:date_system() == "1904")
local sh = assert(wb:sheet("Entrée"))
assert(sh:read(1, 0) == "Zoé")
assert(sh:read(1, 1) == false)
assert(sh:read(1, 2) == "2025-01-02T03:04:05")

local formula = assert(sh:get_formula(1, 3))
assert(xlsx.is_formula(formula))
assert(formula.expression == "SUM(1,2)")
local style = assert(sh:get_style(1, 1))
assert(style.bold == true and style.italic == true)
assert(style.underline == "single" and style.strike == true)
assert(style.font_name == "DejaVu Sans" and style.font_size == 13)
assert(style.font_color == "FF123456" and style.fill_color == "FFABCDEF")
assert(style.border.left.style == "thin")
assert(style.border.right.style == "double")
assert(sh:get_column_width(0) == 20)
assert(sh:get_row_height(0) == 26)
local frozen_rows, frozen_cols = sh:get_frozen_panes()
assert(frozen_rows == 1 and frozen_cols == 2)
assert(sh:get_auto_filter() == "A1:D3")
assert(sh:merged_cells()[1] == "A4:C4")
local input_link = assert(sh:get_hyperlink(1, 0))
assert(input_link.target == "https://example.org/source?a=1&b=2")
assert(wb:get_active_sheet() == "Entrée")
assert(sh:get_tab_color() == "005B9BD5" or sh:get_tab_color() == "FF5B9BD5")
assert(sh:is_row_hidden(4) and sh:is_column_hidden(4))
local input_comment = assert(sh:get_comment(1, 2))
assert(input_comment.author == "Interop" and input_comment.text == "Commentaire openpyxl")
local input_dv = sh:get_data_validations()
assert(#input_dv == 2 and input_dv[1].type == "list" and input_dv[2].type == "whole")
local input_cf = sh:get_conditional_formats()
assert(#input_cf == 1 and input_cf[1].type == "cell" and input_cf[1].operator == "less_than")
assert(input_cf[1].style.fill_color == "FFFFC7CE" or input_cf[1].style.fill_color == "00FFC7CE")
assert(assert(wb:sheet("Masquée")):get_visibility() == "hidden")
assert(assert(wb:sheet("Très masquée")):get_visibility() == "very_hidden")
assert(wb:get_defined_name("ZoneNoms") and wb:get_defined_name("ValeurLocale", 1))

local out = xlsx.new({ date_system = "1904" })
local report = out:add_sheet("Résumé")
local header = xlsx.style({
  bold = true,
  italic = true,
  underline = "double",
  strike = true,
  font_name = "Liberation Sans",
  font_size = 14,
  fill_color = "D9EAF7",
  font_color = "112233",
  horizontal = "center",
  vertical = "center",
  wrap_text = true,
  border = {
    left = { style = "thin", color = "FF0000" },
    right = { style = "double" },
    top = { style = "dashed", color = "00AA00" },
    bottom = { style = "dotted" },
  },
})
local money = xlsx.style({ number_format = "currency_eur" })
local date_bold = xlsx.style({ bold = true })

report:write(0, 0, "Rapport 1.3", header)
report:merge_cells("A1:D1")
report:append_row({ "Nom", "Valeur", "Date", "Lien" }, header)
report:append_row({ "Élise & <test>", 12.5, xlsx.datetime(2024, 3, 15, 13, 45, 30) })
report:write(2, 1, 12.5, money)
report:write(2, 2, xlsx.datetime(2024, 3, 15, 13, 45, 30), date_bold)
report:write(2, 3, xlsx.hyperlink("https://example.com/?a=1&b=2", "Site", {
  tooltip = "Ouvrir le site",
}))
report:append_row({ "Zoé", 7.25, xlsx.datetime(2024, 3, 16, 8, 0, 0), "Cible" })
report:write(3, 1, 7.25, money)
report:write(3, 2, xlsx.datetime(2024, 3, 16, 8, 0, 0), date_bold)
report:set_hyperlink(3, 3, "'Cible'!A1", { internal = true })
report:write(4, 3, xlsx.formula("SUM(B3:B4)", 19.75), money)
report:set_column_width(0, 18)
report:set_column_width(1, 14)
report:set_row_height(0, 28)
report:freeze_panes(2, 1)
report:set_auto_filter("A2:D4")
report:set_tab_color("4472C4")
report:set_row_hidden(5, true)
report:set_column_hidden(4, true)
report:set_comment(3, 0, { author = "Julien", text = "Valeur contrôlée par lua-xlsx" })
report:add_data_validation("A3:A20", {
  type = "list", values = { "Oui", "Non", "En attente" },
  allow_blank = true, show_error_message = true,
})
report:add_data_validation("B3:B20", {
  type = "whole", operator = "between", minimum = 0, maximum = 100,
})
report:add_conditional_format("B3:B20", {
  type = "cell", operator = "less_than", value = 0,
  style = xlsx.style({ font_color = "9C0006", fill_color = "FFC7CE" }),
})
out:add_sheet("Cible"):write(0, 0, "destination")
out:add_sheet("Masquée"):set_visibility("hidden")
out:add_sheet("Très masquée"):set_visibility("very_hidden")
out:set_active_sheet("Résumé")
out:define_name("ZoneNoms", "'Résumé'!$A$3:$A$20")
out:define_name("ValeurLocale", "'Résumé'!$B$3", { local_sheet = "Résumé", hidden = true })

assert(out:save(output))
print("INTEROP LUA : PASS")
