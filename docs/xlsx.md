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

- `book` is a `BookInput`: sheets with their cells (notes and hyperlinks among them) and whether they are hidden, columns and rows (their outline levels among them), merged areas, filter mode, frozen panes, conditional formats, rules of data validation and outline settings, names, the style table, the locale, the date system and iterative calculation.
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

A custom number format is read as its code. A built-in one is stored in the file as a number, and the reader gives that number, 1 to 49, which the book shows by its locale as Excel does: format 14 shows 2025-01-31 as `1/31/2025` in en-US and `31.01.2025` in ru-RU ([docs/formats.md](formats.md)). Formats 50 to 58, East Asian dates, the model does not have: they are read as the code Excel for Mac en-US shows, `m/d/yy`, and reported. That code is the Mac's short date, format 14; Excel for Windows shows them as its own 14 too, `m/d/yyyy` in en-US and `ДД.ММ.ГГГГ` in ru-RU, so where the engine shows 2025-01-31 as `1/31/25` (`1.31.25` in ru-RU) Excel for Windows shows `1/31/2025` (`31.01.2025`), as the engine shows format 14 (`test/xlsx-win.test.ts`).

Excel for Mac saves some built-in formats otherwise than by their number: the money and accounting ones (5 to 8, 41 to 44) as codes in the locale it runs in, 23 to 26 as General, and the East Asian 27 to 36 as 14 and 21. A file Excel saved holds them so.

### Columns, rows and merged areas

A column gets its width when the file sets it by hand, and a row its height likewise. A width or height Excel fits by itself is left out, but a column's width that AutoFit or typing set, which Excel saves with `bestFit`, is read with `bestFit: true` and written so ([Fitting columns](formats.md#fitting-columns)). A hidden row or column gets `hidden: true` and keeps its size, the one it shows again when unhidden. Excel takes a row of height 0 or a column of width 0 as hidden, whether the file marks it hidden or not, and so do the reader and the engine. A row or column with its own look gets its style and number format. Neighbouring lines with equal settings become one span, such as `B:D` or `5:9`.

```js
// <row r="5" ht="30" hidden="1" customHeight="1">, <col min="5" max="5" width="12" hidden="1" customWidth="1"/>
rows: { 5: { height: 30, hidden: true } }, columns: { E: { width: 12, hidden: true } }
```

A grouped row or column gets its `outlineLevel`, 1 to 7, and the summary line of a collapsed group `collapsed: true`; the group's lines are hidden lines ([Data validation and outlines](#data-validation-and-outlines)).

A cell keeps of its look only what its row or column does not give it ([The cells of formatted rows and columns](#the-cells-of-formatted-rows-and-columns)): a date typed into a column of dates comes back as a value alone, and shows the column's format.

A merged area that overlaps another is left out and reported.

### Hidden sheets

A sheet a file marks `state="hidden"` gets `hidden: true`, one marked `state="veryHidden"` gets `hidden: 'very'`: a very hidden sheet is one Excel's window does not offer to show again, only its object model. Formulas read hidden sheets as any other. A book keeps one visible sheet, so a file whose worksheets are all hidden beside a visible chart sheet, which the book leaves out, shows its first worksheet, reported.

```js
// <sheet name="Rates" sheetId="2" state="hidden" r:id="rId2"/>, <sheet name="Keys" sheetId="3" state="veryHidden" r:id="rId3"/>
{ name: 'Rates', cells: { … }, hidden: true }, { name: 'Keys', cells: { … }, hidden: 'very' }
```

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
| `hidden` | a file whose every worksheet is hidden (beside chart sheets): its first worksheet is read as visible |
| `autoFilter` | an autofilter's range, criteria and sort, a table's criteria, an advanced filter |
| `sheetKind` | chart, dialog and macro sheets, left out |
| `cellMetadata` | a rich value in a cell, such as a data type or a picture |
| `comment` | a thread of comments, read as a note; a note of runs with different looks, read as its plain text; a note shown all the time, read as one shown on hover |
| `hyperlink` | a hyperlink whose address is no external relationship of the sheet; one over more than 4,096 cells, read on its cells that hold something |
| `frozenPanes` | split panes that are not frozen; panes frozen over a sheet scrolled down or right, read from the first row and column |
| `conditionalFormat` | a rule of a kind the model does not have, a pivot table's rules, a rule without its formulas or thresholds; what a rule's differential format holds beyond the model's style (alignment, protection, a font's name and size, pattern and gradient fills); an average rule with both equal and deviations, read with equal |
| `dataValidation` | a rule without a range of the sheet or the formula its type needs, or of an unknown type, left out; cells a rule shares with an earlier one, which keep the earlier one; an IME mode; a list whose items hold commas (`x12ac:list`), read split at them |
| `outline` | automatic styles of an outline |
| `drawing`, `pivotTable`, `protection`, `print`, `macros`, `sparkline`, `connection` | parts of the file the model does not have |
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

Styles, number formats, widths, heights, hidden rows and columns, outline levels and merged areas are written as the book has them. The default font is Calibri 11, and a column without a width gets Excel's standard width of 8 digits. Every cell is written with the look it shows, its row's or column's where it has none of its own, as Excel writes it ([The cells of formatted rows and columns](#the-cells-of-formatted-rows-and-columns)).

A built-in format given by its number is written as that number, and Excel shows it by its own locale. A code that a locale shows otherwise than the built-in format it equals, such as `m/d/yyyy` (format 14), is written as a code of the file's own, so every locale shows what `format` says. A code that every locale shows alike, such as `0.00`, is written by its built-in number, as Excel writes it, and comes back as that number, `2`.

### Hidden sheets

A hidden sheet is written with its state, `hidden` or `veryHidden`. The book's active sheet, the one Excel shows on opening, is its first visible sheet: Excel shows a hidden sheet that a file names as active, so a file whose first sheet is hidden names the next visible one (`<workbookView activeTab="1"/>` and that sheet's tab selected).

### Filters

A filtered sheet is written so that Excel takes it as filtered. Its first table with a header row gets criteria in its filter that every row meets, a cell equal to nothing or not equal to it, as Excel keeps a table's filter. A sheet without such a table gets Excel's mark `filterMode` and an autofilter with those criteria over its cells and hidden rows. Excel leaves the rows as the file has them, and its Reapply shows every row. A sheet whose tables have no header row and meet its cells cannot have an autofilter there: its filter mode is left out and reported.

### The date system and calculation

`dateSystem: 1904` is written as `workbookPr date1904`, and dates typed in (`{ date: '2025-01-31' }`) as serial numbers of 1904, 44226. `iteration` is written as `calcPr iterate` with the count and the change the book gives. Every file has `calcPr`: without it Excel keeps the calculation settings it has from a book opened before, iterative calculation among them.

### What the file cannot hold

| code | what happens |
|---|---|
| `meta` | a cell's `meta` is left out |
| `errorValue` | `#CYCLE!` is written as 0, as Excel stores a circular reference |
| `textLength` | text over 32,767 characters is cut, a note's too, a hyperlink's tooltip over 255, a rule's titles over 32, its input message over 255 and its alert's text over 225 |
| `dataValidation` | a rule whose formula does not parse, or names a sheet, a table or a column the book does not have, is left out |
| `formula` | a formula that does not parse, or names a sheet, a table or a column the book does not have, is written as its value |
| `name` | a name Excel does not take, given twice, of a missing sheet, or whose reference does not parse or names a table or a column the book does not have, is left out |
| `autoFilter` | the filter mode of a sheet whose tables have no header row and meet its cells is left out |
| `conditionalFormat` | a rule with a formula the file cannot hold (it does not parse, or names a sheet, a table or a column the book does not have) is left out |

## The cells of formatted rows and columns

A book gives a cell without a style or a format of its own its row's, else its column's, the style and the format each separately: `{ value: 46096 }` in a column of `dddd, mmmm d, yyyy` shows `Sunday, March 15, 2026`. Excel reads a file otherwise. The writer and the reader turn one into the other, so a book written and opened in Excel shows what the engine shows, and a file Excel saved reads into a book that shows what Excel showed.

### Cells the file stores

Excel shows a cell the file stores (a `<c>` element) by the cell's own cell format alone, whatever its row and column have. So the writer writes each cell with the whole look it shows. A value in that column of dates is written with the column's cell format, the one Excel gives a value typed there, and a cell with a bold style of its own in a filled column of dates is written bold, a date and without the fill, as the book shows it.

```js
// columns: { B: { format: 'dddd, mmmm d, yyyy', style: 1 } }, cells: { B2: { value: 46096 } }
// <col min="2" max="2" … style="1"/> and <c r="B2" s="1"><v>46096</v></c>
```

The reader keeps of a stored cell only what its row or column does not give it. The cell above comes back as `{ value: 46096 }`, without a format of its own. Where a cell stores less than its line gives, the reader spells the default out: a cell stored in General in a column of dates gets `format: 'General'`, and one without a fill in a filled column the book's plain style, a style with nothing set.

### Where a row crosses a column

An empty cell the file does not store shows its row's whole cell format when the row has one (`customFormat`), else its column's. So where a row with only a style crosses a column with only a format, Excel shows the empty cell, and a value typed there, without the column's format: `46096` typed where a bold row crosses a column of percents shows `46096`, bold, where the book shows `4609600.00%`.

Excel never leaves such a crossing empty itself: formatting a column across a formatted row, or a row across a column, it writes the cells where they cross, empty, with both lines' formats. The writer does the same: where a row with a look crosses a column with one and Excel would show the empty cell otherwise than the book, it writes that cell with the look the book shows. The reader reads a file without these cells as Excel shows it: a row with a cell format of its own that leaves out a crossing column's style or format, where a cell of that column is missing, gets the plain style or General.

```js
// <row r="6" s="1" customFormat="1"> (bold), <col min="7" max="7" style="3"/> (0.00%), no <c r="G6">
rows: { 6: { style: 0, format: 'General' } }
```

### Lines Excel formats

Excel, formatting a line across a formatted one, merges their formats field by field: a row made bold across a filled column of dates gives the cells where they cross bold, filled and dates. A book's style is one whole, and `setLayout` sets a line's style alone: where a row with a style crosses a column with one, the book shows the row's. A host that formats lines as Excel does gives those cells a style of their own.

### How it was checked

`npx tsx bench/coverage/run.ts lines <set>` writes each book of `bench/coverage/lines/<set>.ts` with `writeXlsx`; Excel for Mac opens it, takes the case's steps through its object model (a number format, bold, italic or a fill set on a column, a row or a cell, a value typed into a cell), reads each listed cell's text, number format and the fill, font colour, bold and italic of its display format, and saves the book. `npx tsx bench/windows/run.ts lines <set>` does the same in Excel for Windows through COM (`NumberFormat`, `Font`, `Interior`, `Value2`, `Range.Text`, `DisplayFormat`). A bare case sends the file without the cells the writer writes where lines cross, as other writers leave it. `test/lines-excel.test.ts` checks, without Excel, that the engine shows what Excel showed after each case, on the book written or on the one the reader reads from a bare file, that it shows the same on the book read from the file Excel saved, and that a book written and saved by Excel reads back as written.

| set | cases | cells | what |
|---|---|---|---|
| `files` | 9 | 586 | formatted columns and rows over numbers, dates, text, booleans, errors, formulas and a spill without a format of their own, cells with their own style or General, empty cells with a style; rows with only a style or a format crossing columns with the other, written and bare, values typed there; Excel formatting a column, a row and lines across each other, values typed into them |
| `held` | 6 | 501 | spans of rows and columns, a row with only a style crossing columns with a format and a style, dates, a spill across formatted columns, cells with their own General, plain style or italic, written and bare, values typed where lines cross; Excel formatting a row and then columns across it |

The engine's answers were committed before each run. The first run of `files`, on the writer that did not write the cells where lines cross, held in 314 of 320 cells: four were empty or typed cells where a row with only a style or a format crosses a column with the other, which both Excels showed by the row's whole cell format, and two were the cells where Excel formatted a filled column across a bold row, which it gave both. The writer and the reader then took the rules above, and the second run of `files` held in 584 of 586 cells, the two being those Excel formatted; `held`, made after it, in 500 of 501, the other being the same merge. Excel for Windows showed what Excel for Mac showed in every cell of both runs, opened every file without the repair window, and the engine shows what Excel showed on every book read from a file Excel saved. The fit set `autofit`, whose cases `colformat` hold a number in a formatted column, now shows the column's format in both Excels and fits the column for it, as the engine does.

## Conditional formats

A sheet's conditional formats (`conditionalFormats`) go both ways. The reader reads the rules of `<conditionalFormatting>` with their differential formats (`<dxfs>` of styles.xml) and Excel 2010's rules of the sheet's extension list (`x14:conditionalFormatting`), and gives them in priority order; the writer writes them so, each rule's priority by its place in the list. Excel 2010's part of a data bar (its fill, border, direction, negative bars and axis) is joined to its 2007 rule by their id; the icon sets `3Stars`, `3Triangles` and `5Boxes`, custom icons and formulas over other sheets are Excel 2010's rules alone, with their style inline. A text, blank or period rule always has its formula in a file; the reader leaves out the one Excel writes for it, which the model makes itself. Ranges come as the model writes them: Excel's `A1:A1048576` is `A:A`. A dxf's number format is read by its code, as Excel saves `0.00%` under another id. [docs/conditional-formats.md](conditional-formats.md) has the rules, what a cell shows by them and how it was checked.

## Notes, hyperlinks and frozen panes

A cell's note and hyperlink and a sheet's frozen panes go both ways, so a grid can show them and a file saved from it keeps them. They are a cell's `note` and `link` and a sheet's `frozen`.

```js
{
  name: 'Sales',
  cells: {
    A1: { value: 'Region', note: { author: 'Ann', text: 'From the CRM\nexport of Monday' } },
    B2: { value: 'Report', link: { url: 'https://example.com/report', tooltip: 'Open the report' } },
    B3: { value: 'Totals', link: { location: "'Totals'!A1" } },
  },
  frozen: { rows: 1, columns: 1 },
}
```

### Notes

A note has its author and its text, lines separated by `\n`. Excel keeps notes in a comments part with a VML drawing that shows them, and the reader takes the comments part. The text is read as plain text: a note made in Excel's window has its author's name in bold at the start, and such formatting is reported. A note shown all the time comes back as one shown when the pointer rests on its cell, also reported. The size and place of a note's box are not kept.

The writer writes both parts, as Excel writes them: each note gets a hidden box beside its cell. A note without an author is written with an empty one, as Excel keeps it.

### Threads of comments

Excel 365's comments are threads: a first comment and its replies, each with its author. The model has notes alone, so the reader takes a thread as a note. The note's author is the first comment's author, and its text is every comment's text, each reply after an empty line with its author:

```
Is this the final figure?

Boris: Yes, from the audit.
```

A thread's dates, its mentions and whether it is resolved are not kept, and each thread is reported. Excel writes, beside each thread, a note of its own for older versions of Excel (`[Threaded comment] Your version of Excel allows you to read...`). The reader leaves those notes out. An Excel that has no threads moves such a note with its cell when rows are inserted and leaves the thread where it was, so the reader puts the thread on the cell of its note.

The writer writes a thread read so as a note. Excel shows it as a note, with the thread's text in it.

### Hyperlinks

A hyperlink has `url`, an address outside the book, or `location`, a place in it (`'Totals'!A1`, a defined name), or both: then the place is inside what the address points to. A `tooltip` is the text Excel shows when the pointer rests on the cell. The cell shows its value, not the link, so a link is set on a cell that already holds its text. A hyperlink over a range is read on each of its cells.

Excel takes the part of an address after `#` as its place: a file holding `https://example.com/page#top` comes back from Excel as `https://example.com/page` with the place `top`. Excel also saves a space in an address as `%20`, and the reader gives the address as the file holds it. A tooltip over 255 characters is cut, as Excel does not open a file with a longer one.

### Frozen panes

`frozen: { rows: 2, columns: 1 }` keeps the first two rows and the first column in view, as Excel's Freeze Panes does with C3 selected. Split panes that are not frozen are not in the model and are reported.

### Structure operations

A note and a hyperlink go with their cell: a move takes them along, a copy copies them, a deleted cell and a cell a move or a copy lands on lose them. Inserted rows and columns get none. A hyperlink's place is kept as written when its sheet is renamed or deleted, as Excel keeps it. Rows inserted above the split of frozen panes push it down, and deleted ones draw it up, but one row stays frozen at least; the same goes for columns. [docs/structure.md](structure.md) has the rules.

### How it was checked

`npx tsx bench/coverage/run.ts parts <set>` writes each book of `bench/coverage/parts/<set>.ts` with `writeXlsx`. Excel for Mac opens it, makes what the case makes through its object model (`add comment`, `make new hyperlink`, `freeze panes`), applies the case's structure operation, reads every sheet's notes, hyperlinks and panes, and saves the book. `npx tsx bench/windows/run.ts parts <set>` does the same in Excel for Windows through COM (`AddComment`, `Hyperlinks.Add`, `FreezePanes`). `test/sheet-parts-excel.test.ts` and `test/sheet-parts-win.test.ts` check, without Excel, that Excel read what the engine holds after each case and that the reader reads every file Excel saved as Excel read it.

| set | cases | what |
|---|---|---|
| `files` | 10 | books written with every kind of note, link and frozen pane, 1,100 notes on one sheet, a long tooltip and a long address, books written after a structure operation in the engine; notes, links and frozen panes Excel made itself |
| `struct-held` | 26 | rows and columns inserted and deleted around frozen panes, notes and links; ranges moved and copied with them and onto them; sheets renamed and deleted under links to them |
| `frozen-held` | 9 | frozen rows and columns deleted past the split and inserted several at once |
| `threads` | 2 | threads of comments put into a file as Excel 365 writes them, with a row inserted above them |

Both Excels opened every book without the repair window. Excel for Mac refused the first version of the book with a tooltip of 300 characters; since then the writer cuts tooltips to 255, and both open it. The engine's answers to `struct-held`, `frozen-held` and `threads` were committed before Excel ran them (`bench/coverage/parts/*.predicted.json`), and so were those to `files`, which held but for the address with `#` and the long tooltip: `struct-held` held in 24 of 26 cases, the two others showing that deleting every frozen row or column leaves one frozen; `frozen-held`, made after that, held in 9 of 9; `threads` in 1 of 2, the other being the inserted row below. Excel for Windows read what Excel for Mac read in every case, but for the author of the notes each made itself, its own user.

Neither Excel takes threads of comments into its object model: Excel 2024 for Windows refuses to make one through COM, Excel for Mac's AppleScript has none, and both show the placeholder note of a thread they open. After inserting a row both moved the placeholder note and left the thread at its cell, which is why the reader follows the placeholder. A file Excel 365 made with threads was not available, so the threads were put into the file by hand (`bench/coverage/parts/threaded.ts`).

## Data validation and outlines

A sheet's rules of data validation and its outline go both ways, so a grid can offer a cell's list and check what is typed, and a file saved from it keeps them. They are a sheet's `validations`, the rows' and columns' `outlineLevel` and `collapsed`, and a sheet's `outline`.

```js
{
  name: 'Budget',
  cells: { A1: { value: 'Item' }, B1: { value: 'Amount' } },
  validations: [{ range: 'B2:B50', type: 'decimal', operator: 'greaterThanOrEqual', formula1: '0', error: 'No negative amounts' }],
  rows: { '3:6': { outlineLevel: 1, hidden: true }, 7: { collapsed: true } },
  outline: { summaryBelow: true },
}
```

### Data validation

A file keeps a rule in `<dataValidations>` with xlsx's defaults false: a rule without `allowBlank`, `showInputMessage` or `showErrorMessage` is read with `allowBlank: false`, `showPrompt: false` or `showError: false`, as the model's defaults are those of a rule Excel makes. `showDropDown="1"` hides a list's dropdown and is read as `dropdown: false`. A rule whose formulas refer to another sheet Excel writes in the sheet's extension, `<x14:dataValidations>`, which the reader reads beside the others. Formulas come in the engine's spelling, as a cell's do. Excel saves a whole column as `F1:F1048576`, read as `F:F`.

Two rules over one cell are not in the model. Excel writes them when a move puts a rule over a cell that keeps another one in its range, and shows the one written first, which is the one the reader keeps; the later rule loses the cell, and that is reported. A list typed where the comma is no list separator (in ru-RU) can hold an item with a comma: Excel keeps that list in an extension (`x12ac:list`), which the reader does not read, and gives the list split at its commas as a fallback, which it reads and reports.

The writer writes every rule in `<dataValidations>`, references to other sheets too, which Excel opens. A rule's titles over 32 characters, its input message over 255 and its alert's text over 225 are cut, as Excel does not open longer ones.

### Outlines

A row or column's `outlineLevel` and `collapsed` are `<row>`'s and `<col>`'s attributes, the summary settings `<sheetPr><outlinePr summaryBelow summaryRight>`. The writer writes the deepest levels into `<sheetFormatPr outlineLevelRow outlineLevelCol>`, where Excel reserves room for the outline's buttons, and the reader takes them from the lines. Automatic outline styles (`applyStyles`) are reported.

### How it was checked

The valid sets of `bench/coverage/valid/` were written by `writeXlsx` and opened by Excel for Mac and for Windows: every book loaded, and Excel read the rules, each cell's check and the outlines the engine holds but for texts cut to Excel's limits. Rules and outlines Excel made itself, and the books Excel saved after structure operations, read back as Excel read them ([docs/validation.md](validation.md#how-it-was-checked)).

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
- `read/` reads the package, the workbook, the sheets and their cells, the cell metadata, the notes, threads, hyperlinks and frozen panes (`parts.ts`), conditional formats (`conditional.ts`) and data validation (`validation.ts`).
- `styles/` turns styles.xml into the style table and the differential formats (`dxf.ts`), and holds the codes of the East Asian built-in formats, the codes the writer stores by number, and the colours.
- `formulas/` changes formulas between the file's spelling and the engine's. `prefixes.ts` is generated by `bench/coverage/xlsx/generate.ts`, so the reader needs neither the harness nor the engine's function definitions. Formulas of the old kind go through `fromLegacy` (`src/legacy/`), whose table of argument classes needs no function definitions either.
- `write/` writes the workbook, the sheets, styles.xml with its differential formats, the notes, hyperlinks and frozen panes (`parts.ts`), conditional formats (`conditional.ts`) and data validation (`validation.ts`).

Reading uses the parser and copy and paste of the structure operations. It does not load the engine's functions or evaluator, and a minified browser bundle of the subpath is about 120 KB.

## How it was checked

`npx tsx bench/coverage/run.ts xlsx <step>` sends files to Excel 16.113 for Mac through AppleScript. Excel opens each file, computes it fully and saves it under another name. The repair window "We found a problem with some content" would mean a broken file, and none appeared. The saved files are in `test/excel/xlsx/`, and `test/xlsx-excel.test.ts` checks them without Excel.

| step | what | result |
|---|---|---|
| `roundtrip` | 181 books written by `writeXlsx`: values, escapes, formulas, spills, styles, sheet names, both locales, the synthetic book, the books of the structure, spill and table sets (names of any formula, `@`, references across sheets, array formulas of the old kind and tables with structured references among them), hidden rows and columns with their sizes, lines of size 0, filtered sheets with and without a table, the 1904 date system, iterative calculation and the built-in formats by number in en-US and ru-RU | every file opens, every book reads back as written, and Excel's values equal the engine's |
| `filters` | 8 files with a filter in each form a file holds it, `SUBTOTAL` and `AGGREGATE` over the rows it hid | read with the filter mode Excel took, every formula computes the value Excel stored |
| `saved` | 97 books of the case sets, 29,089 cases, as the harness sends them to Excel | read back, every case computes as in the book sent, and the values Excel stored equal its stored answers |
| `intersect` | 75 formulas of the old kind | read with `@` where Excel puts it, every formula computes the value Excel stored, but the two with `OFFSET` |
| `colors` | 233 fills: theme colours with Excel's tints and the legacy palette | 212 equal to Excel's colour, the rest within 3 of 255 per channel |
| `formats` | the built-in number formats 1 to 58 | the codes Excel en-US shows equal the engine's for 1 to 49 and the reader's for 50 to 58 |

The round trip compares books by what they mean: a style by its fields, a formula by its canonical text, spans of rows and columns line by line, a line of size 0 as hidden, a code that every locale shows as a built-in format as that format's number.

The 98 files of the workbook sets (`bench/coverage/workbook/`: hidden rows and columns, autofilters with and without criteria and the filter mark, the 1904 date system, iterative calculation), which Excel opened and computed step by step, read into the books the engine was checked with against Excel's answers, and compute the same values (`test/xlsx-workbook.test.ts`).

The answers to the filter files and to nine books of the round trip were predicted and committed before Excel ran them (`bench/coverage/xlsx/predicted.json`): the filter mode of 8 files and 118 values, all as Excel gave them. The books were predicted to come back without differences, and three did not: Excel saved the built-in formats of two by codes and escaped the literal characters of the date formats of the third.

### Hidden sheets

`npx tsx bench/coverage/run.ts hidden <set>` writes each book of `bench/coverage/hidden/<set>.ts` with `writeXlsx`; Excel for Mac opens it, notes the active sheet, hides and shows sheets through its object model (`visible`), applies a sheet operation (rename, delete, insert, move), reads each sheet's visibility, the active sheet and the values of formulas over the sheets, and saves the book. `npx tsx bench/windows/run.ts hidden <set>` does the same in Excel for Windows (`Worksheet.Visible`). `test/hidden-sheets-excel.test.ts` checks, without Excel, that Excel read what the engine holds after each case and that the reader reads every file Excel saved with the visibility Excel read.

| set | cases | what |
|---|---|---|
| `sheets-held` | 29 | books written with hidden and very hidden sheets, the first ones too; sheets hidden, very hidden and shown by Excel; the last visible sheet hidden and deleted; hidden sheets renamed, deleted and moved, sheets inserted and moved beside them; two files changed by hand, one whose active tab is a hidden sheet and one whose every sheet is hidden |
| `sheets-held-2` | 11 | sheets moved before two hidden sheets in a row, a very hidden one and a hidden last one, a hidden sheet moved, sheets inserted before hidden ones |
| `sheets-held-3` | 10 | ends of references across sheets moved inward, outward and past each other, sheets moved past hidden ones at the start and the end |

Both Excels opened every file the writer wrote without the repair window, read each sheet's state as written, showed the first visible sheet, and computed a reference across sheets through a hidden one and `SUBTOTAL` and `AGGREGATE` over a hidden sheet's hidden row as over visible ones. A file whose active tab names a hidden sheet opened with that sheet shown, and one whose every sheet is hidden was repaired. The engine's answers were committed before Excel ran each set (`bench/coverage/hidden/<set>.predicted.json`): `sheets-held` held in 25 of 29 cases on both Excels, `sheets-held-2`, made after it, in 6 of 11, and `sheets-held-3`, made after that, in 10 of 10. The misses became the rules of [docs/structure.md](structure.md#hidden-sheets): a very hidden sheet is neither deleted nor moved, a moved sheet is placed by the visible sheets, and an end of a reference across sheets that moves stays its end until it passes the other. Both Excels answered alike in every case; Excel for Mac refused some deletes without an error, leaving the sheet.

## Known differences

- Excel saves the literal characters of a number format escaped, so `dd.mm.yyyy` comes back as `dd\.mm\.yyyy`. It is the same format.
- Excel saves a range over whole columns or rows as such, so `A1:A1048576` comes back as `A:A`.
- Excel saves a global name qualified by a sheet that has no name of its own, `'S'!x`, as the book's name `[0]!x`. It comes back as `x`, the same name.
- Excel saves a LET that ends with a value with its last name as the calculation. With 127 names that makes 255 arguments, one more than the engine takes, so such a formula does not parse after Excel saved it.
- Excel for Mac saves the built-in money and accounting formats (5 to 8, 41 to 44) as codes in the locale it runs in, 23 to 26 as General and the East Asian 27 to 36 as 14 and 21, so a book read from a file Excel saved has these codes and numbers.
- Some theme colours with a tint differ from Excel's by up to 3 of 255 in a channel.
- The East Asian date formats 50 to 58 are read as the code `m/d/yy`, which Excel for Mac en-US shows for them; Excel for Windows shows them as its short date, format 14, in en-US and ru-RU.
- In a few places Excel computes a formula of the old kind otherwise than the dynamic formula it shows for it. The engine gives the dynamic formula's value. [docs/legacy.md](legacy.md) lists them.
- An autofilter's range and criteria are not kept. A filtered sheet comes back with the writer's criteria, which every row meets, and Excel's Reapply then shows every row.
- A style, a height or `hidden` given to a span of rows is written row by row, so a span to the last row writes a million rows.
- The model has no default font or default column width of its own. The reader reports a file with others, and the writer always writes Calibri 11 and 8 digits.
- Excel takes the part of a hyperlink's address after `#` as its place, so `https://example.com/page#top` comes back as the address `https://example.com/page` and the place `top`. It is the same link.
- Excel saves a space in a hyperlink's address as `%20` and shows it as a space. The reader gives the address as the file holds it.
- Excel for Windows saves a relative address of a hyperlink (`other.xlsx`) relative to the folder it saves the file to.
- A thread of comments is read as a note, without its dates, mentions and state, and written as a note.
- Excel keeps an Excel 2010 data bar's lengths only when they are 0 and 100: a bar of 5 and 95 comes back read as 10 and 90.
- A dxf's number format that is a built-in one comes back from Excel as its code (`0.00%` for 10).
- A list of data validation whose item holds a comma (typed in ru-RU) is read split at the comma. Excel saves a list's reference to a deleted sheet as `#REF!` alone, which the engine keeps as `#REF!$A$1:$A$4`.
