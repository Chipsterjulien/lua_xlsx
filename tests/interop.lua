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

local out = xlsx.new({ date_system = "1904" })
local report = out:add_sheet("Résumé")
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

report:append_row({ "Nom", "Valeur", "Date", "Total" }, header)
report:append_row({ "Élise & <test>", 12.5, xlsx.datetime(2024, 3, 15, 13, 45, 30) })
report:write(1, 1, 12.5, money)
report:write(1, 2, xlsx.datetime(2024, 3, 15, 13, 45, 30), date_bold)
report:append_row({ "Zoé", 7.25, xlsx.datetime(2024, 3, 16, 8, 0, 0) })
report:write(2, 1, 7.25, money)
report:write(2, 2, xlsx.datetime(2024, 3, 16, 8, 0, 0), date_bold)
report:write(3, 3, xlsx.formula("SUM(B2:B3)"), money)
report:set_column_width(0, 18)
report:set_column_width(1, 14)
report:set_row_height(0, 24)
report:freeze_panes(1, 1)
report:set_auto_filter("A1:D3")

assert(out:save(output))
print("INTEROP LUA : PASS")
