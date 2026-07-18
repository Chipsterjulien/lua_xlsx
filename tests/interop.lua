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
report:append_row({ "Nom", "Valeur", "Date" })
report:append_row({ "Élise & <test>", 12.5, xlsx.datetime(2024, 3, 15, 13, 45, 30) })
assert(out:save(output))
print("INTEROP LUA : PASS")
