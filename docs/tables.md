# Tables

A table is Excel's list object: a rectangle of a sheet with a header row of column names, rows of data and, optionally, a totals row. Formulas refer to its parts by name, `SUM(Sales[Amount])` rather than `SUM(C2:C40)`, and such a reference follows the table as rows and columns are inserted and deleted in it. A workbook's tables are calculated, restructured, read from xlsx files and written back as Excel does it.

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

The rules below were verified in Excel 16.113 for Mac (en-US and ru-RU).

## In short

```js
import { compute } from 'xlcalc'

const cells = {
  A1: { value: 'Item' }, B1: { value: 'Qty' }, C1: { value: 'Price' }, D1: { value: 'Total' },
  A2: { value: 'tea' }, B2: { value: 2 }, C2: { value: 10 }, D2: { formula: '[@Qty]*[@Price]' },
  A3: { value: 'coffee' }, B3: { value: 3 }, C3: { value: 20 }, D3: { formula: '[@Qty]*[@Price]' },
  F1: { formula: 'SUM(Orders[Total])' },
  F2: { formula: 'ROWS(Orders[#All])' },
  F3: { formula: 'Orders[[#Headers],[Price]]' },
}
const columns = [{ name: 'Item' }, { name: 'Qty' }, { name: 'Price' }, { name: 'Total', formula: '[@Qty]*[@Price]' }]
const r = compute({ sheets: [{ name: 'S', cells, tables: [{ name: 'Orders', range: 'A1:D3', columns }] }] })
const out = r.sheets[0].cells
console.log(out.F1.value, out.F2.value, out.F3.value) // → 80 3 Price
console.log(out.D2.formula) // → Orders[[#This Row],[Qty]]*Orders[[#This Row],[Price]]
```

A formula stores its references as an xlsx file does (`Orders[[#This Row],[Qty]]` for `[@Qty]`), and `toLocal` displays them as Excel does, in Russian too (`#Данные`). Structure operations grow, shrink and move tables as Excel does, editing a header cell renames its column, and `editStructure` adds, renames, deletes and converts tables. [docs/tables.md](tables.md) has the rules.

A formula typed in a cell of a table refers to it without its name, `[@Qty]`. Without the workbook, the table of the cell is an option: `toLocal(formula, locale, { table })` drops the name as Excel displays such a formula, `toInvariant(text, locale, { table })` adds it back, and `validateFormula` and `validateLocalFormula` with `ctx.table` check the formula as if it were in the table and return the text the workbook stores. `escapeColumn` writes a column's name for use in a reference, with `'` before the characters that have a special meaning there (`[`, `]`, `#`, `'`, a leading `@`).

```js
import { escapeColumn, toInvariant, validateFormula } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

console.log(toInvariant('=[@Qty]*1,5', ruRU, { table: 'Orders' }).formula) // → Orders[@Qty]*1.5
console.log(validateFormula('[@Qty]*[@Price]', { table: 'Orders' }).canonical) // → Orders[[#This Row],[Qty]]*Orders[[#This Row],[Price]]
console.log(`Orders[${escapeColumn('Price #2')}]`) // → Orders[Price '#2]
```

## A table in a workbook

A sheet lists its tables in `tables` (`TableInput`):

| field | |
|---|---|
| `name` | the table's name |
| `range` | its rectangle, including the header and totals rows (`B2:E10`) |
| `headers` | the first row is a header row; true by default |
| `totals` | the last row is a totals row; false by default |
| `columns` | one entry for each column of the range, left to right: its `name`, the `formula` of a calculated column, its `totals` |

A table's name follows Excel's naming rules, the same as for a defined name ([Defined names](names.md) has them, and how they were verified): it starts with a letter of any script, `_` or `\`, continues with letters, digits, `_`, `.`, `\` and `?`, and is not a cell address (`T2` is the cell T2), not `R`, `C` or an R1C1 reference, not TRUE or FALSE, and in a ru-RU workbook not `ИСТИНА` or `ЛОЖЬ`. Tables and defined names share one namespace: no table can have the name of another table or of a defined name, in any case, Cyrillic included (a table `Продажи` and a name `ПРОДАЖИ` cannot coexist). A table has at least one data row.

With a header row, its cells contain the column names as text: `B1` contains `"Qty"` when the second column is named `Qty`. Column names are unique within a table, case-insensitively, and can contain any characters.

A calculated column's `formula` is the formula Excel gives a row inserted into the table. The cells of the column contain their own formulas, which calculate like any other. A column's `totals` records what its cell in the totals row is, as an xlsx file records it: one of Excel's functions (`sum`, `average`, `count`, `countNums`, `max`, `min`, `stdDev`, `var`), `custom` for a formula of its own, or a label. The cell contains the formula or the text, and that is what is calculated.

What Excel does not allow throws an `InputError` with the location and the reason: a table overlapping another, a merged area or a legacy array formula over several cells inside a table, a name already in use, or a header cell that does not contain its column's name.

## Structured references

| reference | refers to |
|---|---|
| `Sales` or `Sales[]` | the data rows |
| `Sales[Qty]` | the data of a column |
| `Sales[[Qty]:[Price]]` | the data of the columns from Qty to Price |
| `Sales[#All]` | the entire table, including the header and totals rows |
| `Sales[#Data]` | the data rows |
| `Sales[#Headers]` | the header row |
| `Sales[#Totals]` | the totals row |
| `Sales[[#Headers],[#Data]]`, `Sales[[#Data],[#Totals]]` | both parts |
| `Sales[[#Headers],[Qty]]` | the header cell of a column; any part can be combined with a column or a span of columns |
| `Sales[@Qty]`, `Sales[[#This Row],[Qty]]` | the column's cell in the formula's own row |
| `Sales[@]`, `Sales[#This Row]` | the table's row in the formula's row |
| `[@Qty]`, `[Qty]` | in a cell of the table, the same in its own table |

A reference returns a range like any other: `SUM` sums it, `ROWS` counts its rows, an operator works on each of its cells, and a formula whose result is a column of a table spills it (`=Sales[Qty]`).

A part the table does not have returns `#REF!`: `Sales[#Totals]` without a totals row, `Sales[#Headers]` without a header row. `[@Qty]` in a row outside the data rows, above or below the table or in its totals row, returns `#VALUE!`. From another sheet, the formula's row is its row number: `Sales[@Qty]` in `Other!B3` refers to the table's cell in row 3.

`Sales[[#Data],[#Totals],[Qty]]` without a totals row refers to the data alone, and so does `Sales[[#Headers],[#Data]]` without a header row.

A name can contain a structured reference (`Sales[Qty]`). `Sales[[#This Row],[Price]]` in a name refers to the row of the cell that uses the name.

### How they are written

Names of tables, columns and specifiers are parsed case-insensitively and stored as the table spells them. Spaces are allowed inside the outer brackets and after the commas between items (`Sales[ [Qty]:[Price] ]`, `Sales[[#Data], [Qty]]`), but not around `:`. A column written alone is all the text up to `]`, so `Sales[ Qty ]` is a column named ` Qty `.

A column name containing `[`, `]`, `#` or `'` gets `'` before each of them: the column `x#y` is `Sales[x'#y]`, `[br]` is `Sales['[br']]`, and `it's` is `Sales[it''s]`. A name that starts with `@` gets one before it too, `Sales['@at]`, since `@` there means the formula's row. Spaces, digits, `$` and letters of any script need nothing: `Sales[Total $]`, `Sales[2024]`. `escapeColumn(name)` (from `xlcalc` and `xlcalc/text`) escapes a name this way.

Excel rejects some forms, and so does the engine, with the error code `syntax`: two specifiers it does not combine (`[#Headers],[#Totals]`, or `#All` or `#This Row` with another), two columns that are not a span (`[[Qty],[Price]]`), and an unknown specifier. An unknown table or column, or `[@Qty]` outside a table, has the error code `unknown_name` and evaluates to `#NAME?`.

A sheet name before a table's name is dropped, as Excel drops it: `S!Sales[Qty]` becomes `Sales[Qty]`.

### Stored text and what Excel displays

A formula stores its structured references as an xlsx file does, whatever form was typed: with the table named, the formula's row as `[#This Row]`, and a table alone as `Sales[]`. `[@Qty]*[@Price]` in a cell of Sales is stored as `Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]`. This is the canonical text of a result's `formula` and of `validateFormula` given the tables.

`toLocal` displays a formula as Excel displays it: `Sales[@Qty]`, `Sales[@[Unit Price]]`, `Sales[@]`, and `Sales` for the table alone. Given the table of the formula's cell, `toLocal(formula, locale, { table: 'Sales' })` drops the name where Excel drops it in a cell of that table: before a single column or a span of columns of the data or of the formula's row (`[@Qty]`, `SUM([Qty])`). A reference with another specifier keeps it (`Sales[#Headers]`, `Sales[[#Totals],[Qty]]`).

```js
import { toLocal, enUS } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

const f = 'Sales[[#This Row],[Qty]]*SUM(Sales[[#Totals],[Price]])'
console.log(toLocal(f, enUS).formula) // → Sales[@Qty]*SUM(Sales[[#Totals],[Price]])
console.log(toLocal(f, enUS, { table: 'Sales' }).formula) // → [@Qty]*SUM(Sales[[#Totals],[Price]])
console.log(toLocal(f, ruRU, { table: 'Sales' }).formula) // → [@Qty]*СУММ(Sales[[#Итоги];[Price]])
```

The reverse direction needs the table too. `toInvariant(text, locale, { table: 'Sales' })` adds the name before each structured reference that lacks one, so the formula parses outside the table (`[@Qty]` becomes `Sales[@Qty]`), and `validateFormula` and `validateLocalFormula` with `ctx.table` check a formula as if it were in a cell of the table: a reference without a table name refers to that table, and the canonical text is the one the workbook stores. With `ctx.tables` the columns are checked as well; without `ctx.table`, such a reference is rejected when `ctx.tables` is given, as outside a table.

```js
import { toInvariant, validateLocalFormula } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

console.log(toInvariant('=[@Qty]*СУММ(Sales[[#Итоги];[Price]])', ruRU, { table: 'Sales' }).formula) // → Sales[@Qty]*SUM(Sales[[#Totals],[Price]])
console.log(validateLocalFormula('=[@Qty]*1,5', ruRU, { table: 'Sales', tables: [{ name: 'Sales', columns: ['Qty', 'Price'] }] }).canonical) // → Sales[[#This Row],[Qty]]*1.5
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

`Sales[[#Заголовки];[Qty]]`, `ЧСТРОК(Sales[#Все])`. The formula's row is displayed as `@` in either language. A column whose name contains a comma (`1,5`, a number typed into a header cell in ru-RU) is written as is when alone, `Sales[1,5]`; after `@`, en-US puts it in brackets, as it does with any name containing a character that has a meaning in a formula (`[@[1,5]]`), while ru-RU, whose list separator is `;`, does not (`[@1,5]`). `toInvariant` parses the Russian specifiers in any case, and the English ones too. A locale provides its words in `tableItems`; without them, a locale uses the English ones.

## Arrays and tables

No array spills inside a table, nor into one: a formula in a table whose result is several values returns `#SPILL!`, and so does a formula next to a table whose array would land on it. A formula in a table can return an array with a single value.

## Structure operations

Structured references name tables and columns, so inserting and deleting rows and columns does not change their text; the table changes, and the references follow it.

### Rows

| rows inserted | the table |
|---|---|
| inside the data, at its first data row, or at the totals row | grows; each new row gets the formulas of the calculated columns |
| at or above the header row | moves down |
| immediately below the table | stays unchanged |

A reference written by address into a table's data grows at its edges as the table does: `SUM(B2:B5)` over the data of a column becomes `SUM(B2:B6)` when a row is inserted at the first data row (row 2) or at the totals row (row 6).

| rows deleted | the table |
|---|---|
| some data rows | shrinks |
| every data row | keeps its first data row, cleared; the rows after it are deleted |
| the totals row | loses it |
| the entire table | is deleted; references to it become `#REF!` |

Excel rejects some of these, and `editStructure` returns `{ ok: false, code: 'table' }`: inserting rows at a row that belongs to two tables at once (side by side); deleting rows that intersect two tables unless both are deleted entirely (deleting one table along with a single row of another is rejected too); and deleting a header row without the rest of its table.

### Columns

A column inserted inside a table, after its first column, adds a column named `Column1`, or with the next free number (`Column2`), and its header cell gets the name. A span such as `Sales[[Region]:[Price]]` includes the new column when it falls between the ends. A column inserted at the first column moves the table, and one inserted to its right changes nothing.

A deleted column removes the references to it: `Sales[Price]`, `Sales[@Price]` and a span ending on it become `#REF!`, and so does `[@Price]` in a calculated column (`[@Qty]*#REF!`). A span that covers the deleted column keeps the rest. Deleting all the columns of a table deletes the table.

Inserting columns into two tables stacked one above the other at once is rejected with `table`, and so is deleting columns that intersect two tables, unless both are deleted entirely.

In a ru-RU workbook, a new column is named `Столбец1`.

### Header cells

Typing over a header cell renames its column, and `Session.setCells` does this: every reference to the column follows, including in calculated columns.

```js
import { createSession } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { value: 'Qty' }, A2: { value: 5 }, C1: { formula: 'SUM(T[Qty])' } }, tables: [{ name: 'T', range: 'A1:A2', columns: [{ name: 'Qty' }] }] }] }
const s = createSession(book)
s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 'Quantity' } }])
console.log(s.getBook().sheets[0].cells.C1.formula, s.getBook().sheets[0].tables[0].columns[0].name) // → SUM(T[Quantity]) Quantity
```

A number or a boolean typed into a header cell becomes its text (`2024`, or `1,5` in ru-RU), as the cell's own number format displays it, if it has one, whatever the column's width (`1,50` with `0.00`, or `31.01.2025` with the short date in ru-RU); where that format displays it as General does, the text is what `&` produces (`0,000000001`, not `1E-09`). A cleared header cell names the column `Column1`, or with the first free number. A name another column already has gets a number on the column further to the right: typing `Price` over `Qty` when the next column is `Price` leaves `Price` and `Price2`. A formula in a header cell throws an `InputError`.

### Copies and moves

A formula copied or moved out of its table gets the table's name in its references: `[@Qty]*[@Price]` pasted outside becomes `Sales[@Qty]*Sales[@Price]`. A copy of an entire table is a new table, named after the original with a number (`Sales3` in a workbook with one table, as Excel numbers it, or the next number a table or a defined name does not have), whose formulas refer to the new table. A numbered name Excel would not accept, such as `Tab3`, a cell address, rejects the copy with `table_name`. An entire table cut and pasted moves, to another sheet too. The formula of a calculated column, which a row inserted into the table gets, is copied and moved with the table, like the formulas of its cells: `A2*2` in a table pasted ten rows down becomes `A12*2`, and a reference to a cell left behind gets its sheet name when the table moves to another sheet.

Pasting over header cells renames their columns to what lands there: the text of a value, or of a formula's value (pasting a copy of `=SUM(Sales[Qty])` over a header names the column `11`). Header cells that are cut away leave their columns named `Column1` and so on. A destination that is not a multiple of the copied cells gets one copy at its top-left cell, as with any copy ([Structure operations](structure.md)), and its other cells keep their formulas, including one in another table's calculated column.

### Tables themselves

| operation | |
|---|---|
| `addTable` | creates a table over cells of a sheet; with a header row, its header cells get the column names |
| `renameTable` | every reference follows; a name Excel would not accept, or one a table or a defined name already has, is rejected with `table_name` |
| `deleteTable` | removes the table and its cells; references to it become `#REF!` |
| `convertTable` | turns the table into plain cells; references become references to its cells with `$` and the sheet name |

`convertTable` writes what Excel's Convert to Range writes: `Sales[Qty]` becomes `'S'!$B$2:$B$5`, `Sales[@Qty]` in row 4 becomes `'S'!$B4` (with the row relative), and `Sales[#All]` the entire range. A part the table does not have becomes the entire table (`Sales[[#Totals],[Qty]]` without a totals row becomes `'S'!$A$1:$D$5`), and in a name, the formula's row of a column becomes the entire row of the table, as Excel converts them.

Deleting a sheet removes the references to its tables: they become `#REF!`. Renaming it changes nothing, as tables are referenced without their sheet. A reference by address in the formula of a calculated column follows a renamed, moved or deleted sheet as in a cell's formula: `'S'!A2` becomes `'Data'!A2` when S is renamed to Data, and `#REF!A2` when S is deleted.

## xlsx

`readXlsx` reads a sheet's table parts into its `tables`: the name formulas use (`displayName`), the range, the header and totals rows, and the columns with their calculated column formulas and totals. Formulas come in the file's spelling, which is the stored one, except for a column name starting with `@`, which a file writes without its escape. A header cell that does not contain its column's name as text is given it, and a table the model cannot represent is left out; both are reported with the code `table`.

`writeXlsx` writes each table as a table part, as Excel does. A formula in a totals row is written as a plain formula, the only kind Excel accepts there, and recorded in the table part as its column's custom totals formula. A table's name alone in a defined name is written as `Sales[]`.

## Checked against Excel

Tested in Excel 16.113 for Mac, en-US and ru-RU, on 176 workbooks, and checked in Excel for Windows; the held-out sets came out as predicted in all but a few cases, which became the rules above or known differences.

## Known differences

- Excel's AutoCorrect extends a table when a value is typed immediately below or to the right of it, and fills a calculated column when a formula is typed into an empty column of a table. `setCells` leaves the table and the other cells unchanged; a host extends a table with `editStructure` (`insertRows` inside it fills the calculated columns). Formulas pasted into part of a calculated column fill the entire column in Excel too; `copyRange` pastes only the cells.
- A table renamed to a name Excel does not accept (`AB12`, `1Ставка`, `?Продажи`) or to an existing defined name: Excel's object model adjusts the name to one it accepts (`AB12_`, `_1Ставка`, `_?Продажи`, `Rate_1`), while its Name Box rejects it; `renameTable` rejects it with `table_name`.
- In ru-RU, Excel for Mac keeps a table renamed `ИСТИНА` or `ложь` and turns `TRUE` into `TRUE_`; Excel for Windows turns `ИСТИНА` into `ИСТИНА_` and keeps `TRUE`, as both do with a defined name. `renameTable` rejects all three: it follows Windows for `ИСТИНА`, and stores formulas in en-US syntax, where `TRUE` is the boolean.
- Excel parses text typed into a cell as a number, a date or a boolean where it can (`2025-01-31` in ru-RU is a date, displayed as `31.01.2025`), and a header cell typed this way names its column by what it displays; `setCells` takes a value as given, so the column is named `2025-01-31`.
- Excel's Fill Right and fill handle shift the column of a structured reference (`SUM(Sales[Qty])` filled right becomes `SUM(Sales[Price])`); `copyRange`, like copy and paste, keeps it.
- A legacy array formula over several cells cannot be in a table, as in Excel; one over a single cell evaluates `[#This Row]` by its own row.
