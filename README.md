# lua-xlsx-df

Pure-Lua **XLSX read/write** and a small **DataFrame** layer (filter / groupby /
aggregate), with **zero external dependencies** — no C extension, no `unzip`, no
SQLite. Designed for and tested on [LuaPilot](https://github.com/Chipsterjulien/luapilot_standalone)
(Lua 5.5), and compatible with any standard Lua 5.3+.

> 🇫🇷 Documentation française : [`doc/documentation-fr.md`](doc/documentation-fr.md)
> · 🇬🇧 English documentation: [`doc/documentation-en.md`](doc/documentation-en.md)

![Lua 5.3+](https://img.shields.io/badge/Lua-5.3%2B-blue)
![License GPLv3](https://img.shields.io/badge/license-GPLv3-blue)

## Features

- **Write** `.xlsx` files: strings, integers, floats, booleans, dates and
  date-times, multiple worksheets, sparse cells. Output is a STORED (uncompressed)
  ZIP that Excel, LibreOffice and openpyxl read without complaint.
- **Read** real-world `.xlsx` files: a complete **DEFLATE inflater written in pure
  Lua** (so compressed files from Excel/LibreOffice/pandas work), shared strings,
  inline strings, multiple sheets, sparse cells, and automatic conversion of date
  cells to ISO 8601 strings.
- **DataFrame** layer: `filter`, `select`, `mutate`, `sort`, `groupby` / `agg`
  (sum, mean, min, max, count, first, last), and exports back to a matrix you can
  hand straight to the writer.
- **No dependencies.** Only the Lua standard library (`string`, `table`, `math`,
  `utf8`, `io`). Nothing is required from LuaPilot itself.

## Requirements

- Lua **5.3, 5.4 or 5.5** (uses bitwise operators, `string.pack`/`unpack`,
  `math.type`, `utf8.char`, integer `//`). Lua 5.1/5.2 are **not** supported.
- A Lua build with the `io` library available (for reading/writing files on disk).

## Installation

Copy `xlsx.lua` and `dataframe.lua` next to your script (or anywhere on your
`package.path`). Under LuaPilot in folder mode, dropping them beside `main.lua`
is enough — `require` finds them automatically.

```lua
local xlsx = require("xlsx")
local DF   = require("dataframe")
```

## Quick start

### Write

```lua
local xlsx = require("xlsx")

local wb = xlsx.new()
local sh = wb:add_sheet("Report")
sh:write(0, 0, "Name")                 -- (row, col), 0-indexed
sh:append_row({ "Alice", 30, true })
sh:write(2, 0, xlsx.date(2024, 3, 15)) -- a real date cell
wb:save("report.xlsx")

-- one-liner for a whole matrix:
xlsx.write_rows("quick.xlsx", { {"x","y"}, {1,2}, {3,4} })
```

### Read

```lua
local xlsx = require("xlsx")

local wb = xlsx.open("report.xlsx")
for _, name in ipairs(wb:sheet_names()) do
  local sh = wb:sheet(name)
  print(sh:read(0, 0))                  -- direct access, 0-indexed
  for row in sh:rows() do               -- row is 1-indexed (row[1] = column A)
    for c = 1, row.n do io.write(tostring(row[c]), "\t") end
    io.write("\n")
  end
end
```

### DataFrame

```lua
local xlsx = require("xlsx")
local DF   = require("dataframe")

local d = DF.from_sheet(xlsx.open("sales.xlsx"):sheet("sales"), { header = true })

local summary = d
  :mutate("total", function(r) return r.qty * r.price end)
  :filter(function(r) return r.total >= 30 end)
  :groupby("city")
  :agg({ total = {"sum","total"}, n = {"count"}, max_price = {"max","price"} })
  :sort("total", { desc = true })

summary:show()
xlsx.write_rows("summary.xlsx", summary:to_rows())
```

## Tests

`selftest.lua` is a self-contained harness: it writes a workbook, reads it back,
runs the DataFrame pipeline and checks dates, all via assertions. Run it on your
target to confirm everything works:

```sh
# LuaPilot folder mode (rename selftest.lua to main.lua in a folder
# that also contains xlsx.lua and dataframe.lua):
./bin/luapilot .

# or with any standard Lua 5.3+:
lua selftest.lua
```

Expected last line: `SELFTEST : PASS  (Lua 5.5)`.

## Documentation

The complete API reference lives in [`doc/`](doc/):

- English — [`doc/documentation-en.md`](doc/documentation-en.md)
- Français — [`doc/documentation-fr.md`](doc/documentation-fr.md)

## Limitations

Intentionally out of scope (all additive later if needed):

- Visual styles (bold, colours, borders) — only date number formats are emitted.
- Formulas — a formula's *cached value* is read, but the formula itself is not.
- Merged cells.
- Writing produces uncompressed (STORED) archives; reading handles both STORED and
  DEFLATE.
- Large files: the reader buffers decompressed output in memory; fine for
  thousands–tens of thousands of rows, heavy for hundreds of MB decompressed.

Dates use Excel's 1900 date system (epoch 1899-12-30, like openpyxl). The only
consequence — shared by nearly every library — is Excel's historical 1900 leap-year
quirk: dates before 1900-03-01 are off by one day.

## License

GNU General Public License v3.0 — see [`LICENSE`](LICENSE).
