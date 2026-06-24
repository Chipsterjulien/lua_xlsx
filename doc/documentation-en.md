# Documentation (English)

Complete API reference for the `xlsx` and `dataframe` modules.

- [Conventions](#conventions)
- [Module `xlsx` — writing](#module-xlsx--writing)
- [Module `xlsx` — dates](#module-xlsx--dates)
- [Module `xlsx` — reading](#module-xlsx--reading)
- [Module `dataframe`](#module-dataframe)
- [End-to-end example](#end-to-end-example)
- [Error contract](#error-contract)
- [Limitations & internals](#limitations--internals)

---

## Conventions

- **Indices are 0-based** for `sheet:write(row, col, …)` and `sheet:read(row, col)`
  (cell `A1` is `(0, 0)`), to keep the writer and reader symmetric.
- The **row iterator** `sheet:rows()` yields **1-based** arrays (`row[1]` is column
  A) because that is how Lua loops read naturally; each yielded `row` carries a
  `.n` field with the number of columns.
- Empty cells are `nil`.
- Date cells are returned **as ISO 8601 strings** when reading.
- All DataFrame operations are **non-destructive**: they return a new DataFrame.

---

## Module `xlsx` — writing

### `xlsx.new() -> workbook`

Creates an empty workbook.

### `workbook:add_sheet([name]) -> sheet`

Adds a worksheet and returns it. `name` defaults to `"SheetN"`. A name must be
1–31 characters and must not contain `: \ / ? * [ ]`.

### `sheet:write(row, col, value) -> sheet`

Writes a single cell. `row` and `col` are non-negative integers (0-based).
`value` may be a `string`, `number`, `boolean`, a **date value** (see below), or
`nil` (empty). `NaN`/`Inf` are rejected. Returns the sheet (chainable).

### `sheet:append_row(values) -> sheet`

Appends a full row after the current last row. `values` is a 1-based array;
trailing `nil`s are ignored. Returns the sheet.

### `sheet:write_rows(matrix) -> sheet`

Appends every row of `matrix` (an array of row arrays). Returns the sheet.

### `workbook:save(path) -> true | nil, err`

Writes the `.xlsx` file. Returns `true`, or `(nil, errmsg)` on I/O failure.
If the workbook has no sheet, an empty `"Sheet1"` is created.

### `xlsx.write_rows(path, matrix [, opts]) -> true | nil, err`

Convenience: writes a single matrix to `path`. `opts.sheet` sets the worksheet
name (default `"Sheet1"`).

```lua
local xlsx = require("xlsx")
local wb = xlsx.new()
local sh = wb:add_sheet("People")
sh:write(0, 0, "Name"); sh:write(0, 1, "Age")
sh:append_row({ "Alice", 30 })
sh:append_row({ "Bob", 25 })
assert(wb:save("people.xlsx"))
```

---

## Module `xlsx` — dates

Dates are written as numbers carrying a date number-format style. Use the
constructors below as a cell value.

### `xlsx.date(year, month, day) -> date value`

A date displayed as `yyyy-mm-dd`. Arguments must be integers.

### `xlsx.datetime(year, month, day [, hour, min, sec]) -> date value`

A date-time displayed as `yyyy-mm-dd hh:mm:ss`. Time parts default to 0.

```lua
sh:write(0, 0, xlsx.date(2024, 3, 15))
sh:write(1, 0, xlsx.datetime(2024, 3, 15, 13, 45, 30))
```

When **reading**, a numeric cell whose style is a date format is returned as an
ISO 8601 string: `"2024-03-15"` or `"2024-03-15T13:45:30"`.

### Helpers

- `xlsx.serial_to_iso(serial, withtime) -> string` — converts an Excel serial
  number to ISO 8601 (`withtime` truthy adds the time part).
- `xlsx.date_to_serial(year, month, day [, hour, min, sec]) -> number` — the
  inverse.

Dates use the Excel 1900 system with effective epoch 1899-12-30 (identical to
openpyxl); conversion is pure integer arithmetic, independent of `os.time`/
`os.date`, so there are no timezone or range issues.

---

## Module `xlsx` — reading

### `xlsx.open(path) -> workbook | nil, err`

Opens an `.xlsx` file for reading. Returns a read workbook, or `(nil, errmsg)`
if the file cannot be opened or is not a valid `.xlsx`.

### `workbook:sheet_names() -> { string, … }`

Ordered list of worksheet names.

### `workbook:sheet(which) -> sheet | nil [, err]`

Returns a worksheet by **name** (string) or **1-based index** (number), or `nil`
if not found. Sheets are parsed lazily and cached.

### `sheet:read(row, col) -> value | nil`

Reads one cell (0-based). Returns the value (`string`, `number`, `boolean`, or
an ISO 8601 date string), or `nil` if empty. `false` is preserved (it does not
collapse to `nil`).

### `sheet:dims() -> maxrow, maxcol`

Returns the 0-based maximum row and column indices, or `-1, -1` for an empty
sheet.

### `sheet:rows() -> iterator`

Returns an iterator yielding one `row` per spreadsheet row, from row 0 to
`maxrow` (including empty rows in the middle). Each `row` is a **1-based** array
(`row[1]` = column A) with empty cells as `nil`, and `row.n` holds the column
count.

```lua
local wb = xlsx.open("people.xlsx")
local sh = wb:sheet("People")
print(sh:read(1, 0))                 -- "Alice"
for row in sh:rows() do
  for c = 1, row.n do io.write(tostring(row[c]), "\t") end
  io.write("\n")
end
```

---

## Module `dataframe`

A DataFrame holds an ordered list of column names and a list of row records
(each a table mapping name → value).

### Construction

#### `df.from_rows(matrix [, opts]) -> DataFrame`

Builds from a matrix (array of 1-based row arrays, optionally carrying `.n`).
`opts.header = true` uses the first row as column names; `opts.columns` supplies
explicit names; otherwise columns are named `col1`…`colN`.

#### `df.from_records(records [, columns]) -> DataFrame`

Builds from an array of records (name → value tables). Passing `columns` is
recommended to fix the order; otherwise columns are inferred and sorted.

#### `df.from_sheet(sheet [, opts]) -> DataFrame`

Convenience: consumes an `xlsx` reader sheet via `sheet:rows()`. `opts` are the
same as `from_rows` (typically `{ header = true }`).

### Introspection

- `df:nrow() -> integer` — number of rows.
- `df:ncol() -> integer` — number of columns.
- `df:colnames() -> { string, … }` — ordered column names.
- `df:column(name) -> { value, … }` — values of one column.
- `df:iter() -> iterator` — iterates over records (name → value tables).

### Transformations (each returns a new DataFrame)

- `df:filter(pred)` — keeps rows where `pred(row)` is truthy.
- `df:select(...)` — projects a subset of columns, in the given order. Accepts
  multiple names or a single array.
- `df:rename(map)` — renames columns according to `{ old = new }`.
- `df:mutate(name, fn)` — adds or replaces a computed column: `value = fn(row)`.
- `df:sort(col [, opts])` — stable sort by `col`; `opts.desc = true` for
  descending. `nil` values always sort last.
- `df:head([n])` / `df:tail([n])` — first/last `n` rows (default 5).

### Grouping & aggregation

#### `df:groupby(...) -> GroupBy`

Groups by one or more columns (multiple names or a single array).

#### `groupby:agg(spec) -> DataFrame`

`spec` is `{ output_name = { func, source_col }, … }`. Functions: `sum`,
`mean` (alias `avg`), `min`, `max`, `count`, `first`, `last`. `"count"` without a
column counts rows in the group. Output columns are the group keys first, then
the aggregates **sorted by name** (deterministic). Group order follows first
appearance.

#### `groupby:count() -> DataFrame`

Shortcut for the number of rows per group, in a column named `n`.

```lua
local g = d:groupby("city", "product"):agg({
  total = { "sum", "amount" },
  avg   = { "mean", "amount" },
  n     = { "count" },
})
```

### Output

- `df:to_records() -> { record, … }` — array of name → value tables.
- `df:to_rows([opts]) -> matrix` — array of 1-based row arrays, directly usable by
  `xlsx.write_rows`. A header row is included unless `opts.header == false`.
- `df:tostring([n]) -> string` — an ASCII table of the first `n` rows (default 20).
- `df:show([n]) -> df` — prints `tostring(n)`; returns the DataFrame (chainable).

---

## End-to-end example

```lua
local xlsx = require("xlsx")
local DF   = require("dataframe")

-- read a real .xlsx (DEFLATE-compressed by Excel/LibreOffice/pandas)
local wb = assert(xlsx.open("sales.xlsx"))
local d  = DF.from_sheet(wb:sheet("sales"), { header = true })

-- transform
local summary = d
  :mutate("total", function(r) return r.qty * r.price end)
  :filter(function(r) return r.total >= 30 end)
  :groupby("city")
  :agg({ total = { "sum", "total" }, n = { "count" } })
  :sort("total", { desc = true })

summary:show()

-- write the result back to a new workbook
assert(xlsx.write_rows("summary.xlsx", summary:to_rows(), { sheet = "summary" }))
```

---

## Error contract

- **Bad argument types** (e.g. a non-integer row index, a table where a string is
  expected) raise an error via `error()` — these are caller bugs.
- **Runtime failures** during reading/writing return `(nil, "xlsx: …")`:
  file cannot be opened, not a valid ZIP/`.xlsx`, missing required part,
  unsupported compression method, truncated/corrupt DEFLATE stream.
- HTTP-style "the operation worked but the data is unusual" cases are not errors:
  an empty sheet reads back as a sheet with `dims() == -1, -1`.

---

## Limitations & internals

**Writing** produces an uncompressed (STORED) ZIP. This needs no compressor and
is read correctly by Excel, LibreOffice and openpyxl. Each entry's CRC-32 is
computed in pure Lua.

**Reading** parses the ZIP central directory and inflates entries with a complete
**pure-Lua DEFLATE decoder** (RFC 1951: stored, fixed and dynamic Huffman blocks,
and LZ77 back-references). XML is parsed with lightweight, xlsx-specific pattern
scanning rather than a general XML parser.

**Out of scope** (additive later): visual styles (bold, colours, borders),
formulas (the cached value is read, not the formula), merged cells, a streaming
reader for very large files, and a SQLite bridge for heavy queries.

**Memory:** the inflater accumulates decompressed bytes in memory before turning
them into a string. This is fine for thousands–tens of thousands of rows; for
hundreds of MB of decompressed data a sliding-window flush would be the natural
optimization.

**Compatibility:** the modules use only the Lua standard library and require Lua
5.3+ (bitwise operators, `string.pack`/`unpack`, `math.type`, `utf8.char`,
integer `//`). They have **no dependency on LuaPilot**; LuaPilot is simply a
convenient Lua 5.5 host.
