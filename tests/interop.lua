-- Interop test executed by run_tests.sh.
local xlsx = require("xlsx")

local input = assert(os.getenv("LUA_XLSX_INPUT"), "LUA_XLSX_INPUT manquant")
local output = assert(os.getenv("LUA_XLSX_OUTPUT"), "LUA_XLSX_OUTPUT manquant")

local wb, err = xlsx.open(input)
assert(wb, err)
assert(wb:date_system() == "1904")
local props = wb:get_properties()
assert(props.title == "Fixture 1.5" and props.creator == "openpyxl")
assert(props.keywords == "lua, xlsx, interop")
local workbook_protection = assert(wb:get_protection())
assert(workbook_protection.structure == true and workbook_protection.password_hash ~= nil)

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
local style_opts = xlsx.style_options(style)
assert(style_opts.locked == false)
local hidden_style = xlsx.style_options(assert(sh:get_style(1, 3)))
assert(hidden_style.hidden == true)
local sheet_protection = assert(sh:get_protection())
assert(sheet_protection.password_hash ~= nil)

local rich = assert(sh:get_rich_text(5, 0))
local rich_runs = xlsx.rich_text_runs(rich)
assert(sh:read(5, 0) == "Attention : texte enrichi openpyxl")
assert(#rich_runs == 2 and rich_runs[1].bold == true and rich_runs[1].font_color == "FFFF0000")
assert(rich_runs[2].italic == true)

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
local cf_types = {}
for _, item in ipairs(input_cf) do cf_types[item.type] = item end
assert(cf_types.cell and cf_types.cell.operator == "less_than")
assert(cf_types.cell.style.fill_color == "FFFFC7CE" or cf_types.cell.style.fill_color == "00FFC7CE")
assert(cf_types.color_scale and cf_types.color_scale.mid.value == 50)
assert(cf_types.data_bar and cf_types.data_bar.color:sub(-6) == "5B9BD5")
assert(cf_types.icon_set and cf_types.icon_set.icons == "3_traffic_lights")

assert(assert(wb:sheet("Masquée")):get_visibility() == "hidden")
assert(assert(wb:sheet("Très masquée")):get_visibility() == "very_hidden")
assert(wb:get_defined_name("ZoneNoms") and wb:get_defined_name("ValeurLocale", 1))
assert(#sh:get_images() == 1 and sh:get_images()[1].format == "png")
local input_charts = sh:get_charts()
local input_chart_types = {}
for _, chart in ipairs(input_charts) do input_chart_types[chart.type] = true end
assert(#input_charts == 5 and input_chart_types.line and input_chart_types.pie
  and input_chart_types.doughnut and input_chart_types.area and input_chart_types.scatter)
assert(#sh:get_tables() == 1 and sh:get_tables()[1].name == "InputTable")
local input_setup = assert(sh:get_page_setup())
assert(input_setup.orientation == "landscape" and input_setup.fit_to_width == 1)
assert(sh:get_page_margins().left == 0.4)
assert(sh:get_header_footer().header_center == "OpenPyXL")
local input_repeat_rows, input_repeat_cols = sh:get_print_titles()
assert(sh:get_print_area() == "A1:D10" and input_repeat_rows == "1:1" and input_repeat_cols == "A:A")
local input_rows, input_cols = sh:get_row_page_breaks(), sh:get_column_page_breaks()
assert(#input_rows == 1 and input_rows[1] == 4 and #input_cols == 1 and input_cols[1] == 3)

local out = xlsx.new({ date_system = "1904" })
out:set_properties({
  title = "Rapport 1.5", subject = "Interopérabilité", creator = "lua-xlsx",
  description = "Classeur 1.5 généré par Lua", keywords = { "lua", "xlsx", "interop" },
  category = "Tests", company = "lua-xlsx", manager = "Julien",
})
out:protect({ password = "secret", structure = true })

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
local money = xlsx.style({ number_format = "currency_eur", locked = false })
local date_bold = xlsx.style({ bold = true })
local hidden_formula = xlsx.style({ hidden = true })

report:write(0, 0, "Rapport 1.5", header)
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
report:write(4, 3, xlsx.formula("SUM(B3:B4)", 19.75), hidden_formula)
report:write_rich_text(6, 0, {
  { text = "Attention : ", bold = true, font_color = "FF0000" },
  { text = "valeur à contrôler", italic = true },
})
report:set_column_width(0, 18)
report:set_column_width(1, 14)
report:set_row_height(0, 28)
report:freeze_panes(2, 1)
report:set_auto_filter("A2:D4")
report:set_tab_color("4472C4")
report:set_row_hidden(5, true)
report:set_column_hidden(4, true)
report:set_comment(3, 0, { author = "Julien", text = "Valeur contrôlée par lua-xlsx" })
report:protect({ password = "feuille", select_unlocked_cells = true, auto_filter = true })
report:add_row_page_break(4)
report:add_column_page_break("D")

report:add_data_validation("A3:A20", {
  type = "list", values = { "Oui", "Non", "En attente" },
  allow_blank = true, show_error_message = true,
})
report:add_data_validation("B3:B20", {
  type = "whole", operator = "between", minimum = 0, maximum = 100,
})
local alert = xlsx.style({ font_color = "9C0006", fill_color = "FFC7CE" })
report:add_conditional_format("B3:B20", {
  type = "cell", operator = "less_than", value = 0, style = alert,
})
report:add_conditional_format("B3:B20", {
  type = "color_scale", min_color = "F8696B", mid_type = "percentile", mid_value = 50,
  mid_color = "FFEB84", max_color = "63BE7B",
})
report:add_conditional_format("C3:C20", { type = "data_bar", color = "5B9BD5" })
report:add_conditional_format("D3:D20", {
  type = "icon_set", icons = "3_traffic_lights", value_type = "number", values = { -1, 0, 3 },
})
report:add_conditional_format("B3:B20", { type = "top", rank = 2, style = alert })
report:add_conditional_format("C3:C20", { type = "above_average", style = alert })

local png = ("89504e470d0a1a0a0000000d49484452000000040000000308020000003b9639910000001049444154789c63fccf80004c0cb8380026760105e4f3b8d50000000049454e44ae426082"):gsub("%x%x", function(pair) return string.char(tonumber(pair,16)) end)
report:add_image_data(png, "png", 7, 0, { width = 96, alt_text = "Logo lua-xlsx", name = "Logo" })
report:add_table("A2:D4", { name = "ResumeTable", style = "TableStyleMedium2" })
report:add_chart({
  type = "line", title = "Valeurs", categories = "A3:A4",
  series = { { name_ref = "B2", values = "B3:B4", color = "4472C4", marker = "circle" } },
  legend_position = "bottom", row = 1, col = 5,
})
report:add_chart({
  type = "pie", title = "Répartition", categories = "A3:A4",
  series = { { name = "Valeur", values = "B3:B4" } },
  legend_position = "bottom", data_labels = { show_percent = true }, row = 18, col = 5,
})
report:add_chart({
  type = "doughnut", title = "Anneau", categories = "A3:A4",
  series = { { name = "Valeur", values = "B3:B4" } }, hole_size = 65, row = 35, col = 5,
})
report:add_chart({
  type = "area", title = "Aire", categories = "A3:A4", grouping = "stacked",
  series = { { name = "Valeur", values = "B3:B4", color = "70AD47" } }, row = 52, col = 5,
})
report:add_chart({
  type = "scatter", title = "Nuage", x_values = "B3:B4",
  series = { { name = "Mesures", values = "C3:C4", marker = "square", smooth = true } },
  legend = false, row = 69, col = 5,
})
report:set_page_setup({ orientation = "landscape", paper_size = "a4", fit_to_width = 1, fit_to_height = 0, horizontal_centered = true })
report:set_page_margins({ left = 0.4, right = 0.4, top = 0.5, bottom = 0.5 })
report:set_header_footer({ header_center = "lua-xlsx 1.5", footer_right = "Page &P / &N" })
report:set_print_area("A1:J30")
report:set_print_titles({ rows = "1:2", columns = "A:A" })
out:add_sheet("Cible"):write(0, 0, "destination")
out:add_sheet("Masquée"):set_visibility("hidden")
out:add_sheet("Très masquée"):set_visibility("very_hidden")
out:set_active_sheet("Résumé")
out:define_name("ZoneNoms", "'Résumé'!$A$3:$A$20")
out:define_name("ValeurLocale", "'Résumé'!$B$3", { local_sheet = "Résumé", hidden = true })

assert(out:save(output))
print("INTEROP LUA : PASS")
