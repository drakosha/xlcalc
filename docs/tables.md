# Tables

A table is Excel's list object: a rectangle of a sheet with a header row of column names, rows of data and, if you want one, a totals row. Formulas refer to its parts by name, `SUM(Sales[Amount])` rather than `SUM(C2:C40)`, and such a reference follows the table as rows and columns are inserted and deleted in it. A book's tables are computed, restructured, read from xlsx files and written back as Excel does it.

```js
import { compute } from 'xlcalc'

const cells = {
  A1: { value: 'Region' }, B1: { value: 'Qty' }, C1: { value: 'Price' }, D1: { value: 'Total' },
  A2: { value: 'North' }, B2: { value: 2 }, C2: { value: 10 }, D2: { formula: '[@Qty]*[@Price]' },
  A3: { value: 'South' }, B3: { value: 3 }, C3: { value: 20 }, D3: { formula: '[@Qty]*[@Price]' },
  F1: { formula: 'SUM(Sales[Total])' },
  F2: { formula: 'SUMIF(Sales[Region],"North",Sales[Qty])' },
}
const columns = [{ name: 'Region' }, { name: 'Qty' }, { name: 'Price' }, { name: 'Total', formula: '[@Qty]*[@Price]' }]
const r = compute({ sheets: [{ name: 'S', cells, tables: [{ name: 'Sales', range: 'A1:D3', columns }] }] })
console.log(r.sheets[0].cells.F1.value, r.sheets[0].cells.F2.value) // → 80 2
console.log(r.sheets[0].cells.D2.formula) // → Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]
```

The rules below were measured in Excel 16.113 for Mac (en-US and ru-RU).

## A table in a book

A sheet lists its tables in `tables` (`TableInput`):

| field | |
|---|---|
| `name` | the table's name |
| `range` | its rectangle, the header and totals rows included (`B2:E10`) |
| `headers` | the first row is a header row; true by default |
| `totals` | the last row is a totals row; false by default |
| `columns` | one entry for each column of the range, left to right: its `name`, the `formula` of a calculated column, its `totals` |

A table's name follows Excel's rules for names: a letter of any script, `_` or `\` first, then letters, digits, `_` and `.`, not a cell address (`T2` is the cell T2) and not `R`, `C` or an R1C1 reference. Tables and defined names share one namespace: no table has the name of another table or of a defined name, in any case. A table has one data row at least.

With a header row, its cells hold the columns' names as text: `B1` holds `"Qty"` when the second column is named `Qty`. Column names are unique in a table, in any case, and may hold any characters.

A calculated column's `formula` is the formula Excel gives a row inserted into the table. The cells of the column hold their own formulas, which compute as any other. A column's `totals` records what its cell of the totals row is, as an xlsx file records it: one of Excel's functions (`sum`, `average`, `count`, `countNums`, `max`, `min`, `stdDev`, `var`), `custom` for a formula of its own, or a label. The cell holds the formula or the text, and that is what computes.

What Excel does not allow is an `InputError` with the place and the reason: a table over another, a merged area or an array formula of the old kind over several cells inside a table, a name taken, a header cell that does not hold its column's name.

## Structured references

| reference | takes |
|---|---|
| `Sales` or `Sales[]` | the data rows |
| `Sales[Qty]` | the data of a column |
| `Sales[[Qty]:[Price]]` | the data of the columns from Qty to Price |
| `Sales[#All]` | the whole table, header and totals rows included |
| `Sales[#Data]` | the data rows |
| `Sales[#Headers]` | the header row |
| `Sales[#Totals]` | the totals row |
| `Sales[[#Headers],[#Data]]`, `Sales[[#Data],[#Totals]]` | both parts |
| `Sales[[#Headers],[Qty]]` | the header cell of a column; any part goes with a column or a span of columns |
| `Sales[@Qty]`, `Sales[[#This Row],[Qty]]` | the column's cell in the formula's own row |
| `Sales[@]`, `Sales[#This Row]` | the formula's row of the table |
| `[@Qty]`, `[Qty]` | in a cell of the table, the same in its own table |

A reference gives a range like any other: `SUM` sums it, `ROWS` counts its rows, an operator works on each of its cells, and a formula whose result is a column of a table spills it (`=Sales[Qty]`).

A part the table lacks gives `#REF!`: `Sales[#Totals]` without a totals row, `Sales[#Headers]` without a header row. `[@Qty]` in a row outside the data rows, above or below the table or in its totals row, gives `#VALUE!`. From another sheet the formula's row is its row number: `Sales[@Qty]` in `Other!B3` takes the table's cell in row 3.

`Sales[[#Data],[#Totals],[Qty]]` without a totals row is the data alone, and `Sales[[#Headers],[#Data]]` without a header row too.

A name can hold a structured reference (`Sales[Qty]`). `Sales[[#This Row],[Price]]` in a name takes the row of the cell that uses the name.

### How they are written

Names of tables, columns and specifiers are read in any case and stored as the table spells them. Spaces may stand inside the outer brackets and after the commas between items (`Sales[ [Qty]:[Price] ]`, `Sales[[#Data], [Qty]]`), not around `:`. A column written alone is all the text up to `]`, so `Sales[ Qty ]` is a column named ` Qty `.

A column name with `[`, `]`, `#` or `'` in it takes `'` before each: the column `x#y` is `Sales[x'#y]`, `[br]` is `Sales['[br']]`, `it's` is `Sales[it''s]`. A name that starts with `@` takes one before it too, `Sales['@at]`, since `@` there means the formula's row. Spaces, digits, `$` and letters of any script need nothing: `Sales[Total $]`, `Sales[2024]`.

Excel refuses some forms, and so does the engine, with the check code `syntax`: two specifiers it does not combine (`[#Headers],[#Totals]`, `#All` or `#This Row` with another), two columns that are not a span (`[[Qty],[Price]]`), an unknown specifier. An unknown table or column, or `[@Qty]` outside a table, is the check code `unknown_name`, `#NAME?`.

A sheet before a table's name is dropped, as Excel drops it: `S!Sales[Qty]` is `Sales[Qty]`.

### Stored text and what Excel shows

A formula stores its structured references as an xlsx file does, whatever form was typed: the table named, the formula's row as `[#This Row]`, a table alone as `Sales[]`. `[@Qty]*[@Price]` in a cell of Sales is stored `Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]`. This is the canonical text of a result's `formula` and of `validateFormula` given the tables.

`toLocal` shows a formula as Excel shows it: `Sales[@Qty]`, `Sales[@[Unit Price]]`, `Sales[@]`, `Sales` for the table alone. Given the table of the formula's cell, `toLocal(formula, locale, { table: 'Sales' })` drops the name where Excel drops it in a cell of that table: before one column or a span of columns of the data or of the formula's row (`[@Qty]`, `SUM([Qty])`). A reference with another specifier keeps it (`Sales[#Headers]`, `Sales[[#Totals],[Qty]]`).

```js
import { toLocal, enUS, ruRU } from 'xlcalc'

const f = 'Sales[[#This Row],[Qty]]*SUM(Sales[[#Totals],[Price]])'
console.log(toLocal(f, enUS).formula) // → Sales[@Qty]*SUM(Sales[[#Totals],[Price]])
console.log(toLocal(f, enUS, { table: 'Sales' }).formula) // → [@Qty]*SUM(Sales[[#Totals],[Price]])
console.log(toLocal(f, ruRU, { table: 'Sales' }).formula) // → [@Qty]*СУММ(Sales[[#Итоги];[Price]])
```

### In ru-RU

Excel with the Russian interface writes the specifiers in Russian and separates the items with `;`:

| en-US | ru-RU |
|---|---|
| `#All` | `#Все` |
| `#Data` | `#Данные` |
| `#Headers` | `#Заголовки` |
| `#Totals` | `#Итоги` |
| `#This Row` | `#Эта строка` |

`Sales[[#Заголовки];[Qty]]`, `ЧСТРОК(Sales[#Все])`. The formula's row is shown as `@` in either language. `toInvariant` reads the Russian specifiers in any case, and the English ones too. A locale gives its words in `tableItems`; without them a locale writes the English ones.

## Arrays and tables

No array spills inside a table, nor into one: a formula in a table whose result is several values gives `#SPILL!`, and so does a formula beside a table whose array would land on it. A formula in a table may give an array of one value.

## Structure operations

Structured references name tables and columns, so inserted and deleted rows and columns do not change their text; the table changes, and they follow it.

### Rows

| rows inserted | the table |
|---|---|
| inside the data, at its first data row, or at the totals row | grows; each new row gets the formulas of the calculated columns |
| at or above the header row | moves down |
| right below the table | stays as it is |

A reference written by address into a table's data grows at its edges as the table does: `SUM(B2:B5)` over the data of a column becomes `SUM(B2:B6)` when a row is inserted at the first data row (row 2) or at the totals row (row 6).

| rows deleted | the table |
|---|---|
| some data rows | shrinks |
| every data row | keeps its first data row, cleared; the rows after it are deleted |
| the totals row | loses it |
| the whole table | is deleted; references to it become `#REF!` |

Excel refuses some, and `restructure` returns `{ ok: false, code: 'table' }`: rows inserted at a row inside two tables at once (side by side), rows deleted from two tables that keep some of their rows, a header row deleted without the rest of its table.

### Columns

A column inserted inside a table, after its first column, adds a column named `Column1`, or the next free number (`Column2`), and its header cell takes the name. A span such as `Sales[[Region]:[Price]]` takes the new column in when it lies between the ends. A column inserted at the first column moves the table, and one right of it changes nothing.

A deleted column takes the references to it: `Sales[Price]`, `Sales[@Price]` and a span ending on it become `#REF!`, and so does `[@Price]` in a calculated column (`[@Qty]*#REF!`). A span over the deleted column keeps the rest. All the columns of a table deleted delete the table.

Columns inserted into or deleted from two tables one above the other at once are refused with `table`.

In a ru-RU book a new column is named `Столбец1`.

### Header cells

Typing over a header cell renames its column, and `Session.setCells` does this: every reference to the column follows, calculated columns too.

```js
import { openSession } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { value: 'Qty' }, A2: { value: 5 }, C1: { formula: 'SUM(T[Qty])' } }, tables: [{ name: 'T', range: 'A1:A2', columns: [{ name: 'Qty' }] }] }] }
const s = openSession(book)
s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 'Quantity' } }])
console.log(s.book().sheets[0].cells.C1.formula, s.book().sheets[0].tables[0].columns[0].name) // → SUM(T[Quantity]) Quantity
```

A number or a boolean typed into a header cell becomes its text (`2024`). A cleared header cell names the column `Column1`, the first free number. A name another column already has gives a number to the column further right: typing `Price` over `Qty` when the next column is `Price` leaves `Price` and `Price2`. A formula in a header cell is an `InputError`.

### Copies and moves

A formula copied or moved out of its table gets the table's name in its references: `[@Qty]*[@Price]` pasted outside is `Sales[@Qty]*Sales[@Price]`. A copy of a whole table is a new table, named after it with a number (`Sales3` from a book of one table, as Excel numbers it), whose formulas refer to the new table. A whole table cut and pasted moves, to another sheet too.

A paste over header cells renames their columns to what lands there: the text of a value, or of a formula's value (a copy of `=SUM(Sales[Qty])` over a header gives the column the name `11`). Header cells cut away leave their columns named `Column1` and on.

### Tables themselves

| operation | |
|---|---|
| `addTable` | a table over cells of a sheet; with a header row, its header cells take the columns' names |
| `renameTable` | every reference follows; a name Excel would not take, or one a table or a defined name has, is refused with `table_name` |
| `deleteTable` | the table and its cells go; references to it become `#REF!` |
| `convertTable` | the table becomes plain cells; references become references to its cells with `$` and the sheet |

`convertTable` writes what Excel's Convert to Range writes: `Sales[Qty]` becomes `'S'!$B$2:$B$5`, `Sales[@Qty]` in row 4 becomes `'S'!$B4` (the row relative), `Sales[#All]` the whole range. A part the table lacks becomes the whole table (`Sales[[#Totals],[Qty]]` without a totals row is `'S'!$A$1:$D$5`), and in a name the formula's row of a column becomes the whole row of the table, as Excel converts them.

Deleting a sheet takes the references to its tables: they become `#REF!`. Renaming it changes nothing, as tables are named without their sheet.

## xlsx

`readXlsx` reads a sheet's table parts into its `tables`: the name formulas use (`displayName`), the range, the header and totals rows, the columns with their calculated column formulas and totals. Formulas come in the file's spelling, which is the stored one but for a column name starting with `@`, which a file writes without its escape. A header cell that does not hold its column's name as text is given it, and a table the model cannot hold is left out; both are reported with the code `table`.

`writeXlsx` writes each table as a table part, as Excel does. A formula of a totals row is written as a plain formula, the only kind Excel opens there, and recorded in the table part as its column's custom totals formula. A table's name alone in a defined name is written `Sales[]`.

## How it was checked

`npx tsx bench/coverage/run.ts tables <set>` writes each case's book with `writeXlsx`, opens it in Excel, takes the case's steps (cells typed, structure and table operations) through AppleScript, and reads the values of the case's cells, their formulas as Excel shows them, in Russian too for a ru-RU set, and the tables of every sheet. `test/tables-excel.test.ts` replays the steps in a session and compares the steps refused, the tables, the values and the formulas, read as stored text and as `toLocal` shows them.

| set | cases | what |
|---|---|---|
| `refs` | 8 | every part of a table, the formula's row from everywhere, references without a table name inside one, no header row, a totals row, escapes, names, arrays |
| `refs-held` | 5 | predicted before Excel computed them: a table away from A1, the totals row, more escapes, names, arrays beside tables |
| `ops` | 51 | rows and columns inserted and deleted around tables, renames, conversions, deletions, sheets, header cells, copies and moves |
| `ops-held` | 18 | predicted before Excel took them: inserts beside another table, a table emptied, pastes over header rows, renames Excel refuses |
| `locale` (ru-RU) | 4 | the Russian specifiers shown and typed, a totals row, the names of new columns |
| `more` | 5 | whole tables copied and moved, cells typed beside and into a table |

The predicted sets came out as predicted on 81 of 81 cells (`refs-held`), 424 of 433 (`ops-held`: two pastes over header cells, which Excel renames, and a conversion seen through a name were found there and fixed) and 69 of 69 (`locale`).

The xlsx round trip (`run.ts xlsx roundtrip`) sends the 29 distinct books of the sets `refs`, `refs-held`, `ops` and `ops-held`: Excel opens each file `writeXlsx` writes and saves it, the files it saved read back to the books written, and the values Excel stored in them equal the engine's.

## Known differences

- Excel's AutoCorrect extends a table when a value is typed right below it or right of it, and fills a calculated column when a formula is typed into an empty column of a table. `setCells` leaves the table and the other cells as they are; a host extends a table with `restructure` (`insertRows` inside it fills the calculated columns).
- A table renamed to a name Excel does not take (`AB12`) or to a defined name: Excel's object model makes the name one it takes (`AB12_`, `Rate_1`), as its name box refuses it; `renameTable` refuses with `table_name`.
- Excel's Fill Right and fill handle move the column of a structured reference (`SUM(Sales[Qty])` filled right is `SUM(Sales[Price])`); `copyRange`, as copy and paste, keeps it.
- An array formula of the old kind over several cells cannot be in a table, as in Excel; one over a single cell computes `[#This Row]` by its own row.
