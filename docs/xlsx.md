# Reading and writing xlsx

The subpath `xlcalc/xlsx` reads an xlsx file into a book the engine computes, and writes a book with its computed values back to xlsx. It has no dependencies and runs wherever the web platform's deflate streams exist: Node 20.19 or newer and current browsers. Neither function computes anything: reading gives you the values Excel stored beside the book, and writing takes the result you computed.

```js
import { compute } from 'xlcalc'
import { readXlsx, writeXlsx } from 'xlcalc/xlsx'

const { book, cached, issues } = await readXlsx(bytes)
const result = compute(book)
const { data } = await writeXlsx(book, { result })
```

## Reading

`readXlsx(data)` takes the file's bytes, a `Uint8Array` or an `ArrayBuffer`, and resolves to three things.

- `book` is a `BookInput`: sheets with their cells, columns, rows, merged areas and filter mode, names, the style table, the locale, the date system and iterative calculation.
- `cached` holds the values the file stores for formulas and for the cells their arrays spilled into, by sheet in the book's order. These are what Excel computed when it saved the file, so you can compare them with the engine's result.
- `issues` lists everything the book cannot hold, with the sheet and the cells or the name it is about. Nothing is dropped without an issue.

A file that is not a zip archive, an encrypted one, a ZIP64 archive and a file without a workbook part are errors.

### Cells

Text, numbers and booleans become values. Dates stay numbers, as Excel stores them, and their number format makes them dates on screen. The ISO dates of Strict files (`t="d"`) become serial numbers too.

Text keeps its spaces and line breaks. Excel's escapes of characters XML cannot hold are decoded: `_x000B_` is a vertical tab, and `_x005F_x0041_` is the literal text `_x0041_`. Rich text is read as its plain text and reported. Phonetic readings of East Asian text are not part of the text.

The model has no error values, so an error stored as a cell's value is read as a formula of that error: a cell holding `#N/A` gets the formula `#N/A`.

A cell is placed by its address, or after the previous cell when it has none. A `dimension` that does not match the cells is ignored.

### Formulas

Formulas come in the engine's spelling.

| in the file | in the book |
|---|---|
| `_xlfn.XLOOKUP(…)`, `_xlfn._xlws.FILTER(…)` | `XLOOKUP(…)`, `FILTER(…)` |
| `_xlfn.LET(_xlpm.x,1,_xlpm.x+1)` | `LET(x,1,x+1)` |
| `_xlfn.ANCHORARRAY(E2)` | `E2#` |
| `_xlfn.SINGLE(A1:A3)` | `@A1:A3` |
| `Jan:Mar!A1`, `'Jan 1:Mar 2'!A1` | the same reference across sheets |
| `_xludf.MYFN(1)` | `MYFN(1)`, reported |

A shared formula is written once in a file and taken by a block of cells. Each of these cells gets its own formula, made as a copy and paste would make it: `A1*2` in B1 becomes `A2*2` in B2. A master cell below or right of its copies works too, and a reference pushed past the sheet's edge becomes `#REF!`.

An array formula spills from its top-left cell. The other cells of its range hold only stored values in the file, so they go to `cached`, and they keep their own style in the book. Excel 365 marks its dynamic array formulas with cell metadata. An array formula without that mark is of the old kind, entered with Ctrl+Shift+Enter. Its cell gets the formula and `array`, the range it covers, and the formula is computed over that range as Excel computes it ([docs/legacy.md](legacy.md)).

A plain formula, without the mark, is of the old kind too. Excel 365 computes it with implicit intersection and shows `@` where it takes one value of a range. The reader gives the formula Excel shows, with `fromLegacy`:

| in the file | in the book |
|---|---|
| `A1:A3*2`, `ABS(A1:A3)` | `@A1:A3*2`, `ABS(@A1:A3)` |
| `INDEX(A1:B5,0,1)` | `@INDEX(A1:B5,0,1)` |
| `span*1`, `span` a name over a range | `@span*1` |
| `SUM(A1:A3)`, `SUMPRODUCT(A1:A3*2)`, `IF(A1>0,1,2)` | as written |

A formula that needs no `@` keeps its text as written. Excel 365 writes a dynamic formula as a plain one when the two kinds agree, so most formulas of a file Excel 365 saved come through unchanged.

Excel 365 stores `#SPILL!` and `#CALC!` as `#VALUE!` with a rich value that names the real error. The reader reads them as `#SPILL!` and `#CALC!`.

### Styles and number formats

Each cell format of the file becomes a style of the book's table and a number format. Equal styles are shared. A style holds a font, a solid fill, borders and alignment. Font fields equal to the book's default font are left out, as the model takes them from the default.

Colours become `#RRGGBB`. A theme colour with a tint is computed as Excel shows it, and the legacy palette or the book's own palette gives indexed colours.

A custom number format is read as its code. A built-in one is stored in the file as a number, and the reader gives that number, 1 to 49, which the book shows by its locale as Excel does: format 14 shows 2025-01-31 as `1/31/25` in en-US and `31.01.2025` in ru-RU ([docs/formats.md](formats.md)). Formats 50 to 58, East Asian dates, the model does not have: they are read as the code Excel for Mac en-US shows, `m/d/yy`, and reported.

Excel for Mac saves some built-in formats otherwise than by their number: the money and accounting ones (5 to 8, 41 to 44) as codes in the language it runs in, 23 to 26 as General, and the East Asian 27 to 36 as 14 and 21. A file Excel saved holds them so.

### Columns, rows and merged areas

A column gets its width when the file sets it by hand, and a row its height likewise. A width or height Excel fits by itself is left out. A hidden row or column gets `hidden: true` and keeps its size, the one it shows again when unhidden. Excel takes a row of height 0 or a column of width 0 as hidden, and so does the reader, whether the file marks it hidden or not. A row or column with its own look gets its style and number format. Neighbouring lines with equal settings become one span, such as `B:D` or `5:9`.

```js
// <row r="5" ht="30" hidden="1" customHeight="1">, <col min="5" max="5" width="12" hidden="1" customWidth="1"/>
rows: { 5: { height: 30, hidden: true } }, columns: { E: { width: 12, hidden: true } }
```

A merged area that overlaps another is left out and reported.

### Filters

A sheet whose rows a filter hid gets `filtered: true`. Then `SUBTOTAL` with codes 1 to 11 leaves out its hidden rows too, as Excel does, and so do structure operations ([docs/structure.md](structure.md)). Excel takes a sheet as filtered by what its filters hold, not by the mark `filterMode` alone:

| in the file | the sheet |
|---|---|
| an autofilter with criteria, with the mark or without it | filtered |
| a table's filter with criteria | filtered |
| the mark without an autofilter, as an advanced filter leaves it | filtered |
| an autofilter without criteria, with the mark or without it | not filtered |

Every hidden row of a filtered sheet counts as one the filter hid, whatever hid it. The book keeps the filter mode alone: the autofilter's range, its criteria and its sort, a table's criteria and an advanced filter are reported.

### The date system and calculation

A file with `workbookPr date1904` gives a book with `dateSystem: 1904`. Numbers stay as the file stores them: in such a book serial 0 is 1 January 1904, and ISO dates of Strict files count from there too.

`calcPr iterate` gives `iteration`, with `maxIterations` from `iterateCount` and `maxChange` from `iterateDelta` when the file has them. A count outside 1 to 32,767 or a change below 0 is read as Excel's default and reported. The count and the change of a book without iterative calculation are not kept: they change nothing.

```js
// <calcPr iterate="1" iterateCount="7" iterateDelta="0.5"/>
iteration: { maxIterations: 7, maxChange: 0.5 }
```

The manual calculation mode, "automatic except for data tables" and precision as displayed are reported.

### Names

| in the file | in the book |
|---|---|
| `Data!$A$1`, `'Data'!$D:$D` | the reference |
| `#REF!`, `Data!#REF!` | the lost reference, as the engine keeps it |
| `localSheetId="1"` | `scope`, the second sheet's name |
| `0.2`, `Data!$A$1*2`, `Data!$A$1:$A$3,Data!$C$1:$C$3` | the constant, the formula, the union ([docs/names.md](names.md)) |
| `_xlnm.Print_Area` and other built-in names | left out and reported |
| `_xlfn.XLOOKUP` (Excel's placeholder for a newer function) | left out |

### Tables

A sheet's table parts become its `tables`: the name formulas use, the range, the header and totals rows, each column with its calculated column formula and its totals function or label. Structured references come in the stored spelling, which is a file's but for a column name that starts with `@` (`Sales[@at]` in a file is `Sales['@at]`). A header cell that does not hold its column's name as text is given it, and a table the model cannot hold, over another or over a merged area, is left out; both are reported. [docs/tables.md](tables.md) has the rules.

### The locale

xlsx has no place for the book's locale. The writer keeps it in a custom document property named `xlcalc.locale`, and the reader reads it back. Excel keeps the property when it saves the file. A file without it gives a book without a locale, that is en-US.

### What is reported

| code | what the book cannot hold |
|---|---|
| `richText` | formatted runs of text |
| `userFunction` | a user-defined function (`_xludf.`) |
| `table` | a table left out (over another table or a merged area, its name taken), a header cell given its column's name, a structured reference without a table name |
| `externalLink` | links to other workbooks, and formulas that use them |
| `formula` | a formula the engine does not parse, kept as written |
| `errorConstant` | an error stored as a value, read as a formula of it |
| `errorValue` | an error the model has no code for, such as `#GETTING_DATA` |
| `dataTable` | a what-if data table, read as its stored values |
| `name` | built-in names, hidden names, names of a sheet left out |
| `calculation` | the manual mode, automatic but for data tables, precision as displayed, an iteration count or change out of range |
| `style` | pattern and gradient fills, diagonal borders, rotation, shrinking, sub- and superscript, protection |
| `defaultFont` | a default font other than Calibri 11 |
| `defaultWidth` | a default column width other than 8 digits |
| `numberFormat` | an East Asian built-in format (50 to 58), or an unknown one |
| `hidden` | hidden sheets, read as visible ones |
| `autoFilter` | an autofilter's range, criteria and sort, a table's criteria, an advanced filter |
| `sheetKind` | chart, dialog and macro sheets, left out |
| `cellMetadata` | a rich value in a cell, such as a data type or a picture |
| `conditionalFormat`, `dataValidation`, `comment`, `hyperlink`, `drawing`, `pivotTable`, `protection`, `frozenPanes`, `outline`, `print`, `macros`, `sparkline`, `connection` | parts of the file the model does not have |
| `locale` | a locale the package does not have |
| `part` | a missing sheet part, a cell outside the sheet |

View settings such as zoom, gridlines and the selection are not reported. A page setup is reported only when someone made one, for example a landscape page or a fit to width.

## Writing

`writeXlsx(book, opts)` resolves to `{ data, issues }`: the file's bytes and what the file cannot hold. The same book and result give the same bytes.

| option | |
|---|---|
| `result` | the book's `ComputeResult`, from `compute` or `Session.result()`: the values of formulas and of the cells their arrays spilled into |
| `mode: 'formulas'` | the default: formulas with their values from `result`. Without `result` the file has no values, and Excel computes them when it opens the file |
| `mode: 'values'` | values only: a formula's cell gets its value from `result`, which is then required |

```js
const values = await writeXlsx(book, { result, mode: 'values' })
```

### Formulas

Every formula is written as a dynamic array formula, as Excel 365 saves one: with cell metadata and the range its array spilled into. The engine computes every formula by the rules of dynamic arrays, and so does Excel with this mark. Excel 2019 and older show such formulas as array formulas in braces.

The spilled cells are written with their values. `#SPILL!` and `#CALC!` are written as Excel 365 writes them, so Excel reads them back as those errors.

Functions newer than Excel 2007 get their prefix (`_xlfn.XLOOKUP`, `_xlfn._xlws.SORT`), LET's names become `_xlpm.x`, `E2#` becomes `_xlfn.ANCHORARRAY(E2)` and `@x` becomes `_xlfn.SINGLE(x)`. Sheet names are always quoted.

A cell with `array` is written as an array formula of the old kind over its range, without the cell metadata, as Excel writes a formula entered with Ctrl+Shift+Enter. A formula in a table's totals row is written as a plain formula, the only kind Excel opens there. A name's formula is written as the book has it: a constant, a formula or a union. The list of prefixes covers every function Microsoft lists as newer than 2007, not only the whitelist.

A file with formulas asks Excel to compute the book when it opens it, so Excel shows its own values over the stored ones.

### Tables

Each table is written as a table part, as Excel writes one: its name, range, header and totals rows, an autofilter over the header and data, its columns with their calculated column formulas and totals. A formula of the totals row that no totals function of its column gives is recorded as the column's custom totals formula, without which Excel does not open the file, and a table's name alone in a defined name is written `Sales[]`.

### Styles, columns and rows

Styles, number formats, widths, heights, hidden rows and columns and merged areas are written as the book has them. The default font is Calibri 11, and a column without a width gets Excel's standard width of 8 digits.

A built-in format given by its number is written as that number, and Excel shows it by its own locale. A code that a locale shows otherwise than the built-in format it equals, such as `m/d/yy` (format 14), is written as a code of the file's own, so every locale shows what `format` says. A code that every locale shows alike, such as `0.00`, is written by its built-in number, as Excel writes it, and comes back as that number, `2`.

### Filters

A filtered sheet is written so that Excel takes it as filtered. Its first table with a header row gets criteria in its filter that every row meets, a cell equal to nothing or not equal to it, as Excel keeps a table's filter. A sheet without such a table gets Excel's mark `filterMode` and an autofilter with those criteria over its cells and hidden rows. Excel leaves the rows as the file has them, and its Reapply shows every row. A sheet whose tables have no header row and meet its cells cannot have an autofilter there: its filter mode is left out and reported.

### The date system and calculation

`dateSystem: 1904` is written as `workbookPr date1904`, and dates typed in (`{ date: '2025-01-31' }`) as serial numbers of 1904, 44226. `iteration` is written as `calcPr iterate` with the count and the change the book gives. Every file has `calcPr`: without it Excel keeps the calculation settings it has from a book opened before, iterative calculation among them.

### What the file cannot hold

| code | what happens |
|---|---|
| `meta` | a cell's `meta` is left out |
| `errorValue` | `#CYCLE!` is written as 0, as Excel stores a circular reference |
| `textLength` | text over 32,767 characters is cut |
| `formula` | a formula that does not parse, or names a sheet, a table or a column the book does not have, is written as its value |
| `name` | a name Excel does not take, given twice, of a missing sheet, or whose reference does not parse or names a table or a column the book does not have, is left out |
| `autoFilter` | the filter mode of a sheet whose tables have no header row and meet its cells is left out |

## Speed

Measured on an Apple Silicon Mac with `test/fixtures/xlsx-bench-child.ts`, the median of three runs. The ranges are over several runs on a busy machine.

| book | file | read | write |
|---|---|---|---|
| 20,000 rows × 40 values | 3.0 MB | 0.7 to 1.05 s | 1.2 to 1.8 s |
| the synthetic book, 3,400 rows, 3,900 formulas with values | 0.41 MB | 0.1 s | 0.1 to 0.15 s |

A read book of 800,000 values holds about 100 MB of heap. `FORMULA_BENCH=1 npx vitest run test/xlsx-budget.test.ts` checks that the large book reads in 2 s at most.

## How the code is organised

`src/xlsx/` holds the code, and `src/xlsx.ts` is the subpath's entry.

- `zip/` reads and writes the archive. Inflating and deflating go through `DecompressionStream` and `CompressionStream`.
- `xml/` holds the text layer: entities, Excel's escapes, and elements found by regular expressions.
- `read/` reads the package, the workbook, the sheets and their cells, and the cell metadata.
- `styles/` turns styles.xml into the style table and holds the codes of the East Asian built-in formats, the codes the writer stores by number, and the colours.
- `formulas/` changes formulas between the file's spelling and the engine's. `prefixes.ts` is generated by `bench/coverage/xlsx/generate.ts`, so the reader needs neither the harness nor the engine's function definitions. Formulas of the old kind go through `fromLegacy` (`src/legacy/`), whose table of argument classes needs no function definitions either.
- `write/` writes the workbook, the sheets and styles.xml.

Reading uses the parser and copy and paste of the structure operations. It does not load the engine's functions or evaluator, and a minified browser bundle of the subpath is about 120 KB.

## How it was checked

`npx tsx bench/coverage/run.ts xlsx <step>` sends files to Excel 16.113 for Mac through AppleScript. Excel opens each file, computes it fully and saves it under another name. The repair window "We found a problem with some content" would mean a broken file, and none appeared. The saved files are in `test/excel/xlsx/`, and `test/xlsx-excel.test.ts` checks them without Excel.

| step | what | result |
|---|---|---|
| `roundtrip` | 181 books written by `writeXlsx`: values, escapes, formulas, spills, styles, sheet names, both locales, the synthetic book, the books of the structure, spill and table sets (names of any formula, `@`, references across sheets, array formulas of the old kind and tables with structured references among them), hidden rows and columns with their sizes, lines of size 0, filtered sheets with and without a table, the 1904 date system, iterative calculation and the built-in formats by number in en-US and ru-RU | every file opens, every book reads back as written, and Excel's values equal the engine's; Excel hides a row of height 0, which the engine counts in a book that does not say `hidden` |
| `filters` | 8 files with a filter in each form a file holds it, `SUBTOTAL` and `AGGREGATE` over the rows it hid | read with the filter mode Excel took, every formula computes the value Excel stored |
| `saved` | 91 books of the case sets, 28,831 cases, as the harness sends them to Excel | read back, every case computes as in the book sent, and the values Excel stored equal its stored answers |
| `intersect` | 75 formulas of the old kind | read with `@` where Excel puts it, every formula computes the value Excel stored, but the two with `OFFSET` and the one with the intersection operator |
| `colors` | 233 fills: theme colours with Excel's tints and the legacy palette | 212 equal to Excel's colour, the rest within 3 of 255 per channel |
| `formats` | the built-in number formats 1 to 58 | the codes Excel en-US shows equal the engine's for 1 to 49 and the reader's for 50 to 58 |

The round trip compares books by what they mean: a style by its fields, a formula by its canonical text, spans of rows and columns line by line, a line of size 0 as hidden, a code that every locale shows as a built-in format as that format's number.

The 98 files of the workbook sets (`bench/coverage/workbook/`: hidden rows and columns, autofilters with and without criteria and the filter mark, the 1904 date system, iterative calculation), which Excel opened and computed step by step, read into the books the engine was checked with against Excel's answers, and compute the same values (`test/xlsx-workbook.test.ts`).

The answers to the filter files and to nine books of the round trip were predicted and committed before Excel ran them (`bench/coverage/xlsx/predicted.json`): the filter mode of 8 files and 118 values, all as Excel gave them. The books were predicted to come back without differences, and three did not: Excel saved the built-in formats of two by codes and escaped the literal characters of the date formats of the third.

## Known differences

- Excel saves the literal characters of a number format escaped, so `dd.mm.yyyy` comes back as `dd\.mm\.yyyy`. It is the same format.
- Excel saves a range over whole columns or rows as such, so `A1:A1048576` comes back as `A:A`.
- Excel saves a global name qualified by a sheet that has no name of its own, `'S'!x`, as the book's name `[0]!x`. It comes back as `x`, the same name.
- Excel saves a LET that ends with a value with its last name as the calculation. With 127 names that makes 255 arguments, one more than the engine takes, so such a formula does not parse after Excel saved it.
- Excel for Mac saves the built-in money and accounting formats (5 to 8, 41 to 44) as codes in the language it runs in, 23 to 26 as General and the East Asian 27 to 36 as 14 and 21, so a book read from a file Excel saved has these codes and numbers.
- Some theme colours with a tint differ from Excel's by up to 3 of 255 in a channel.
- In a few places Excel computes a formula of the old kind otherwise than the dynamic formula it shows for it. The engine gives the dynamic formula's value. [docs/legacy.md](legacy.md) lists them.
- The intersection operator, a space between two ranges, is not read: such a formula is reported.
- Excel takes a row of height 0 as hidden. The reader reads it so, but a book that gives a row `height: 0` without `hidden` computes `SUBTOTAL` 101 to 111 and `AGGREGATE` as if the row were shown.
- An autofilter's range and criteria are not kept. A filtered sheet comes back with the writer's criteria, which every row meets, and Excel's Reapply then shows every row.
- A style, a height or `hidden` given to a span of rows is written row by row, so a span to the last row writes a million rows.
- The model has no default font or default column width of its own. The reader reports a file with others, and the writer always writes Calibri 11 and 8 digits.
