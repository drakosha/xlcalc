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

- `book` is a `BookInput`: sheets with their cells, columns, rows and merged areas, names, the style table and the locale.
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
| `_xludf.MYFN(1)` | `MYFN(1)`, reported |

A shared formula is written once in a file and taken by a block of cells. Each of these cells gets its own formula, made as a copy and paste would make it: `A1*2` in B1 becomes `A2*2` in B2. A master cell below or right of its copies works too, and a reference pushed past the sheet's edge becomes `#REF!`.

An array formula spills from its top-left cell. The other cells of its range hold only stored values in the file, so they go to `cached`, and they keep their own style in the book. Excel 365 marks its dynamic array formulas with cell metadata. An array formula without that mark is of the old kind, entered with Ctrl+Shift+Enter, and is reported because its array now spills.

Excel 365 computes a plain formula of the old kind with implicit intersection: where such a formula puts a range in a place that takes one value, Excel takes the value in the formula's row. The engine computes it as a dynamic array, so the reader reports these formulas. `A1:A3*2`, `ABS(A1:A3)`, `SUM(A1:A3*2)` and `IF(TRUE,A1:A3)` are reported, while `SUM(A1:A3)`, `COUNTIF(A:A,">0")`, `VLOOKUP(1,A1:C5,2,FALSE)` and `SUMPRODUCT(A1:A3*2)` are not. Excel 365 writes a plain formula only where both ways agree: none of the 194 files it saved for the checks below has such an issue.

Excel 365 stores `#SPILL!` and `#CALC!` as `#VALUE!` with a rich value that names the real error. The reader reads them as `#SPILL!` and `#CALC!`.

### Styles and number formats

Each cell format of the file becomes a style of the book's table and a number format. Equal styles are shared. A style holds a font, a solid fill, borders and alignment. Font fields equal to the book's default font are left out, as the model takes them from the default.

Colours become `#RRGGBB`. A theme colour with a tint is computed as Excel shows it, and the legacy palette or the book's own palette gives indexed colours.

A custom number format is read as its code. A built-in one is stored in the file as a number, and the reader gives it the code Excel for Mac en-US shows: format 14 is `m/d/yy`, format 4 is `#,##0.00`. Excel shows the date, time and currency ones by the regional settings, so they are reported.

### Columns, rows and merged areas

A column gets its width when the file sets it by hand, and a row its height likewise. A width or height Excel fits by itself is left out. Hidden rows and columns get a size of 0 and are reported. A row or column with its own look gets its style and number format. Neighbouring lines with equal settings become one span, such as `B:D` or `5:9`.

A merged area that overlaps another is left out and reported.

### Names

| in the file | in the book |
|---|---|
| `Data!$A$1`, `'Data'!$D:$D` | the reference |
| `#REF!`, `Data!#REF!` | the lost reference, as the engine keeps it |
| `localSheetId="1"` | `scope`, the second sheet's name |
| `0.2`, `Data!$A$1*2` | kept as written and reported: the engine takes only references |
| `_xlnm.Print_Area` and other built-in names | left out and reported |
| `_xlfn.XLOOKUP` (Excel's placeholder for a newer function) | left out |

### The locale

xlsx has no place for the book's locale. The writer keeps it in a custom document property named `xlcalc.locale`, and the reader reads it back. Excel keeps the property when it saves the file. A file without it gives a book without a locale, that is en-US.

### What is reported

| code | what the book cannot hold |
|---|---|
| `richText` | formatted runs of text |
| `legacyArray` | an array formula of the old kind |
| `implicitIntersection` | a formula of the old kind that Excel computes with `@` |
| `singleOperator` | the `@` operator (`_xlfn.SINGLE`): the formula computes to `#NAME?` |
| `userFunction` | a user-defined function (`_xludf.`) |
| `table` | tables, and structured references in formulas |
| `reference3d` | references across sheets, such as `Sheet1:Sheet3!A1` |
| `externalLink` | links to other workbooks, and formulas that use them |
| `formula` | a formula the engine does not parse, kept as written |
| `errorConstant` | an error stored as a value, read as a formula of it |
| `errorValue` | an error the model has no code for, such as `#GETTING_DATA` |
| `dataTable` | a what-if data table, read as its stored values |
| `name` | constants, formulas, built-in and hidden names |
| `date1904` | the 1904 date system: numbers in date formats move to the 1900 system |
| `calculation` | iterative calculation, the manual mode |
| `style` | pattern and gradient fills, diagonal borders, rotation, shrinking, sub- and superscript, protection |
| `defaultFont` | a default font other than Calibri 11 |
| `defaultWidth` | a default column width other than 8 digits |
| `numberFormat` | a built-in format shown by the regional settings, or an unknown one |
| `hidden` | hidden sheets, rows and columns |
| `sheetKind` | chart, dialog and macro sheets, left out |
| `cellMetadata` | a rich value in a cell, such as a data type or a picture |
| `conditionalFormat`, `dataValidation`, `comment`, `hyperlink`, `drawing`, `autoFilter`, `pivotTable`, `protection`, `frozenPanes`, `outline`, `print`, `macros`, `sparkline`, `connection` | parts of the file the model does not have |
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

Functions newer than Excel 2007 get their prefix (`_xlfn.XLOOKUP`, `_xlfn._xlws.SORT`), LET's names become `_xlpm.x`, and `E2#` becomes `_xlfn.ANCHORARRAY(E2)`. Sheet names are always quoted. The list of prefixes covers every function Microsoft lists as newer than 2007, not only the whitelist.

A file with formulas asks Excel to compute the book when it opens it, so Excel shows its own values over the stored ones.

### Styles, columns and rows

Styles, number formats, widths, heights and merged areas are written as the book has them. The default font is Calibri 11, and a column without a width gets Excel's standard width of 8 digits. A number format that Excel would show by the regional settings, such as `m/d/yy`, is written as a custom code, so every locale shows what `format` says.

### What the file cannot hold

| code | what happens |
|---|---|
| `meta` | a cell's `meta` is left out |
| `errorValue` | `#CYCLE!` is written as 0, as Excel stores a circular reference |
| `textLength` | text over 32,767 characters is cut |
| `formula` | a formula that does not parse, or names a sheet the book does not have, is written as its value |
| `name` | a name Excel does not take, given twice, of a missing sheet, or whose reference does not parse, is left out |

## Speed

Measured on this Mac with `test/fixtures/xlsx-bench-child.ts`, the median of three runs.

| book | file | read | write |
|---|---|---|---|
| 20,000 rows × 40 values | 3.0 MB | 0.75 to 0.9 s | 1.2 to 1.6 s |
| the synthetic book, 3,400 rows, 3,900 formulas with values | 0.41 MB | 0.1 s | 0.1 to 0.15 s |

A read book of 800,000 values holds about 100 MB of heap. `FORMULA_BENCH=1 npx vitest run test/xlsx-budget.test.ts` checks that the large book reads in 2 s at most.

## How the code is organised

`src/xlsx/` holds the code, and `src/xlsx.ts` is the subpath's entry.

- `zip/` reads and writes the archive. Inflating and deflating go through `DecompressionStream` and `CompressionStream`.
- `xml/` holds the text layer: entities, Excel's escapes, and elements found by regular expressions.
- `read/` reads the package, the workbook, the sheets and their cells, the cell metadata, and checks implicit intersection.
- `styles/` turns styles.xml into the style table and holds the built-in number formats and the colours.
- `formulas/` changes formulas between the file's spelling and the engine's. `prefixes.ts` and `modes.ts` are generated by `bench/coverage/xlsx/generate.ts`, so the reader needs neither the harness nor the engine's function definitions.
- `write/` writes the workbook, the sheets and styles.xml.

Reading uses the parser and copy and paste of the structure operations. It does not load the engine's functions or evaluator, and a minified browser bundle of the subpath is about 120 KB.

## How it was checked

`npx tsx bench/coverage/run.ts xlsx <step>` sends files to Excel 16.113 for Mac through AppleScript. Excel opens each file, computes it fully and saves it under another name. The repair window "We found a problem with some content" would mean a broken file, and none appeared. The saved files are in `test/excel/xlsx/`, and `test/xlsx-excel.test.ts` checks them without Excel.

| step | what | result |
|---|---|---|
| `roundtrip` | 106 books written by `writeXlsx`: values, escapes, formulas, spills, styles, sheet names, both locales, the synthetic book, and the books of the structure and spill sets | every file opens, every book reads back as written, and Excel's values equal the engine's |
| `saved` | 88 books of the case sets, 27,694 cases, as the harness sends them to Excel | read back, every case computes as in the book sent, and the values Excel stored equal its stored answers |
| `intersect` | 75 formulas of the old kind | every formula whose value Excel computes otherwise is reported |
| `colors` | 233 fills: theme colours with Excel's tints and the legacy palette | 212 equal to Excel's colour, the rest within 3 of 255 per channel |
| `formats` | the built-in number formats 1 to 58 | the reader's codes equal Excel's |

The round trip compares books by what they mean: a style by its fields, a formula by its canonical text, spans of rows and columns line by line.

## Known differences

- Excel saves the literal characters of a number format escaped, so `dd.mm.yyyy` comes back as `dd\.mm\.yyyy`. It is the same format.
- Excel saves a range over whole columns or rows as such, so `A1:A1048576` comes back as `A:A`.
- Excel saves a LET that ends with a value with its last name as the calculation. With 127 names that makes 255 arguments, one more than the engine takes, so such a formula does not parse after Excel saved it.
- A built-in date, time or currency format reads as the code Excel for Mac en-US shows. Excel for Windows en-US shows format 14 as `m/d/yyyy`, and ru-RU as `dd.mm.yyyy`.
- Some theme colours with a tint differ from Excel's by up to 3 of 255 in a channel.
- The implicit intersection check knows the functions of the whitelist. A formula of the old kind with a function outside it, such as `OFFSET`, is not checked.
- Hidden rows and columns lose the size they had before they were hidden.
- A style or height given to a span of rows is written row by row, so a span to the last row writes a million rows.
- The model has no default font or default column width of its own. The reader reports a file with others, and the writer always writes Calibri 11 and 8 digits.
