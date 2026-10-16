# Reading and writing xlsx

The subpath `xlcalc/xlsx` reads an xlsx file into a book the engine computes, and writes a book with its computed values back to xlsx. It has no dependencies and runs wherever the web platform's deflate streams exist: Node 20.19 or newer and current browsers. Neither function computes anything: reading gives you the values Excel stored beside the book, and writing takes the result you computed.

```js
import { compute } from 'xlcalc'
import { readXlsx, writeXlsx } from 'xlcalc/xlsx'

const book = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*SEQUENCE(2)' } } }] }
const { data } = await writeXlsx({ book, result: compute(book) }) // a Uint8Array, for writeFile or a download
// read back, as from `await readFile('book.xlsx')`: the book, the values Excel would show, what the book cannot hold
const read = await readXlsx(data)
console.log(read.book.sheets[0].cells, read.cached.sheets[0].cells.A3.value, read.issues)
// → { A1: { value: 2 }, A2: { formula: 'A1*SEQUENCE(2)' } } 4 []
```

## In short

`readXlsx(data)` from `xlcalc/xlsx` reads a file's bytes (a `Uint8Array`, a Node `Buffer` or an `ArrayBuffer`: `await readFile(path)`, `await file.arrayBuffer()`) into a book. Beside the book it gives `cached`, the values Excel stored for formulas and spilled cells in the shape of a result, and `issues`, everything the book cannot hold, such as rich text, with the sheet and cells. `compute(read.book, { cached: read.cached })` and `createSession` with the same option open a book with a circular reference as Excel opens the file ([Circular references](cycles.md)); nothing else is taken from `cached`. `writeXlsx({ book, result })` writes a book with the values of its formulas: `result` is `compute`'s, a session's (`session.getSnapshot()` gives the book and its result in one call) or the `cached` of a file read. Formulas are written as dynamic array formulas with their spills, and `mode: 'values'` writes values only.

```js
import { compute, createSession } from 'xlcalc'
import { readXlsx, writeXlsx } from 'xlcalc/xlsx'

const book = { sheets: [{ name: 'Sales', cells: { A1: { value: 120 }, A2: { value: 80 }, B1: { formula: 'SUM(A1:A2)' }, C1: { formula: 'SEQUENCE(2)' } } }] }
const { data } = await writeXlsx({ book, result: compute(book) })

const read = await readXlsx(data)
console.log(read.book.sheets[0].cells.B1) // → { formula: 'SUM(A1:A2)' }
console.log(read.cached.sheets[0].cells.B1, read.cached.sheets[0].cells.C2) // → { value: 200 } { value: 2, spillOf: 'C1' }
console.log(read.issues) // → []

const session = createSession(read.book)
session.setCells([{ sheet: 'Sales', cell: 'A1', input: { value: 150 } }])
const saved = await readXlsx((await writeXlsx(session.getSnapshot())).data)
console.log(saved.cached.sheets[0].cells.B1) // → { value: 230 }

const values = await readXlsx((await writeXlsx({ book, result: compute(book) }, { mode: 'values' })).data)
console.log(values.book.sheets[0].cells.C2) // → { value: 2 }
console.log((await writeXlsx({ book })).issues.map((i) => i.code)) // → [ 'noCachedValues' ]
```

`values` says which values of formulas the file stores. With a result it is `'require'`: every formula and every cell of an array has its value in the result, else the result is not the book's and `writeXlsx` throws `InputError`. `'preserve'` writes the values the result has and reports the formulas without one (`noCachedValues`). Without a result it is `'omit'`: the file stores no values, reported as `noCachedValues` when the book has formulas, but not when `values: 'omit'` asks for it. A file with formulas asks Excel to compute them when it opens it, whichever values it stores. A formula a circular reference stops stores its `held` value where the result gives none (with `omit` too), as Excel opens the file holding it. A session is not a snapshot (an asynchronous one lives in a thread): pass `session.getSnapshot()`, awaited for an asynchronous session. Nor is a result alone, which has no book: `writeXlsx(compute(book))` throws `InputError` with `pass { book, result }`.

Both functions return promises: the archive is inflated and deflated by the platform's streams. Hidden and very hidden sheets (`hidden: true` and `hidden: 'very'`; the writer makes the first visible sheet the active one), styles, number formats (built-in ones by their number), widths, heights, hidden rows and columns, outline levels, merged areas, notes, hyperlinks, frozen panes, conditional formats (Excel 2010's rules too), rules of data validation, names, tables, a sheet's filter mode, the date system, iterative calculation and the book's locale go both ways. A cell without a style or a format of its own is written with its row's, else its column's, as Excel writes a value typed there, so a date in a column of dates shows as a date in Excel, and it reads back without one. A cell the file stores with a look alone, as Excel stores each cell of a formatted area, is read into its sheet's `looks` ([Looks of many cells](grid.md#looks-of-many-cells)), and the cells of `looks` are written back as such cells. A group's collapsed mark is written as Excel keeps and saves it: on the summary row of a group whose rows are all hidden, rows hidden through `setLayout` too, and not on one whose rows show ([docs/outline.md](outline.md#in-xlsx-files)).

```js
import { readXlsx, writeXlsx } from 'xlcalc/xlsx'

const book = {
  sheets: [{
    name: 'Sales',
    cells: {
      A1: { value: 'Region', note: { author: 'Ann', text: 'From the CRM' } },
      B1: { value: 'Report', link: { url: 'https://example.com/report', tooltip: 'Open the report' } },
    },
    frozen: { rows: 1 },
  }],
}
const back = (await readXlsx((await writeXlsx({ book })).data)).book.sheets[0]
console.log(back.cells.A1.note, back.frozen) // → { author: 'Ann', text: 'From the CRM' } { rows: 1 }
```

A thread of comments, as Excel 365 makes them, is read as a note: its first author, and every comment's text with the replies' authors. [docs/xlsx.md](xlsx.md) lists what is read, written and reported, and how it was checked against Excel.

An issue's `message` is for a person, in English by default. `opts.locale` of either function writes it in another language: the tag of a loaded locale or a locale object, as the other functions take a locale. The code, sheet, cells and name of an issue stay as they are, so a program reads them alike in every language.

```js
import { registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'
import { readXlsx, writeXlsx } from 'xlcalc/xlsx'

registerLocale(ruRU)
const book = { locale: 'ru-RU', sheets: [{ name: 'План', cells: { A1: { value: 2 }, B1: { formula: 'A1*2' } } }] }
const { data, issues } = await writeXlsx({ book }, { locale: 'ru-RU' })
console.log(issues[0].code, issues[0].message) // → noCachedValues Результат не передан: файл не хранит значения формул, Excel вычислит их при открытии файла
console.log((await writeXlsx({ book })).issues[0].message) // → no result was given: the file stores no values of formulas, and Excel computes them when it opens the file
console.log((await readXlsx(data, { locale: ruRU })).book.locale) // → ru-RU
```

Pass the locale the person reading the messages chose. A file does not say it: Excel keeps no language in xlsx, and the locale `xlcalc.locale` that the writer stores, which `readXlsx` gives as `book.locale`, is the book's, the rules it computes by, which a person working in another language may open too. So neither function takes the language from the book.

## Reading

`readXlsx(data)` takes the file's bytes, a `Uint8Array` (a Node `Buffer` is one, as `readFile` gives it), another view of bytes or an `ArrayBuffer` (`await file.arrayBuffer()` in a browser), and resolves to three things.

- `book` is a `BookInput`: sheets with their cells (notes and hyperlinks among them), the looks of their cells of a look alone ([Cells of a look alone](#cells-of-a-look-alone)) and whether they are hidden, columns and rows (their outline levels among them), merged areas, filter mode, frozen panes, conditional formats, rules of data validation and outline settings, names, the style table, the locale, the date system, iterative calculation and the links to other workbooks with the values stored of them ([Links to other workbooks](#links-to-other-workbooks)).
- `cached` holds the values the file stores for formulas and for the cells their arrays spilled into, by sheet in the book's order, in the shape of a result: `cached.sheets[0].cells.B1.value` as `compute(book).sheets[0].cells.B1.value`. A formula whose array spilled has `spill`, its range, and each other cell of the range `spillOf`, the formula's address, as in a result. These are what Excel computed when it saved the file, so you can compare them with the engine's result cell by cell, or write the file again with them (`writeXlsx({ book, result: cached })`). A book with a circular reference opens with them as Excel opens the file, `compute(book, { cached })` or `createSession(book, { cached })`: the formulas the cycle stops hold the stored values ([cycles.md](cycles.md)).
- `issues` lists everything the book cannot hold, with the sheet and the cells or the name it is about. Nothing is dropped without an issue. Each issue has a code for programs and a message for people, in English or in the language of `opts.locale` ([The language of the messages](#the-language-of-the-messages)).

A file that is not a zip archive, an encrypted one, a ZIP64 archive and a file without a workbook part are errors. Data that is no bytes, such as the file's path, is an `InputError` that says to pass the bytes.

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

A prefix is one only at a word's start: a name or a table's name goes on through `?` as through letters, so `a?_xlfn.b` is the name `a?_xlfn.b` and `T?[Qty]` the table `T?`, as Excel loads them.

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

A fill is read as Excel shows it. A solid fill shows its `fgColor`; without one, or with one that names no colour (`auto`, a system colour), it is white, Excel's automatic colour. A fill without a pattern type is no fill; of a gradient and a pattern in one fill the pattern counts. A gradient and patterns other than solid are not in the model: the cell is read without a fill, and that is reported.

A custom number format is read as its code. A built-in one is stored in the file as a number, and the reader gives that number, 1 to 49, which the book shows by its locale as Excel does: format 14 shows 2025-01-31 as `1/31/2025` in en-US and `31.01.2025` in ru-RU ([docs/formats.md](formats.md)). Formats 50 to 58, East Asian dates, the model does not have: they are read as the code Excel for Mac en-US shows, `m/d/yy`, and reported. That code is the Mac's short date, format 14; Excel for Windows shows them as its own 14 too, `m/d/yyyy` in en-US and `ДД.ММ.ГГГГ` in ru-RU, so where the engine shows 2025-01-31 as `1/31/25` (`1.31.25` in ru-RU) Excel for Windows shows `1/31/2025` (`31.01.2025`), as the engine shows format 14 (`test/xlsx-win.test.ts`).

Excel for Mac saves some built-in formats otherwise than by their number: the money and accounting ones (5 to 8, 41 to 44) as codes in the locale it runs in, 23 to 26 as General, and the East Asian 27 to 36 as 14 and 21. A file Excel saved holds them so.

### Cells of a look alone

Excel stores each cell of a formatted area, a row formatted to column XFD among them, as a cell with a cell format and nothing else (`<c r="B2" s="3"/>`). Such a cell does not become an entry of `cells`: the reader puts its look into its sheet's `looks` ([Looks of many cells](grid.md#looks-of-many-cells)) as it reads the row, by runs of columns of one look, and rows with the same runs share them. A cell with a value or a formula stays an entry with its own look, as before.

```js
// <row r="20"><c r="A20" s="1"/><c r="B20" s="1"/> … <c r="XFD20" s="1"/></row>, and so rows 21 and 22
looks: [{ range: 'A20:XFD22', style: 0 }]
```

So a sheet of the wide shape, 344 rows of such cells from A to XFD in two looks, the second in every 12th column, 5.6 million cells, reads into 5,571 rectangles and about a megabyte of heap, where it took 815 MB as entries. A cell that gets content later in the reading, a shared formula's copy, a note, a hyperlink or a table's header, takes its look of the file as its own, and its place among the stored cells, as an entry of that look would; a cell an array spilled into is left with its look alone, in the looks. A hyperlink over more cells than it is given one by one goes to the cells of the looks in its range as to the stored ones.

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

### The saved active sheet and view

A grid can reopen a workbook on its saved sheet and at its saved selection, scroll position and zoom. `book.activeSheet` names the active visible sheet. Each sheet's `view` holds `activeCell`, `selection` (an array of A1 cells or ranges), `topLeft`, and `zoom` (percent, 10 to 400). The active sheet's tab is the one selected.

```js
const book = {
  activeSheet: 'Plan',
  sheets: [{ name: 'Plan', cells: {}, view: {
    activeCell: 'C4', selection: ['C4:D5', 'F7'], topLeft: 'B3', zoom: 125,
  } }],
}
```

The reader takes `workbookView.activeTab` from the first workbook view and the first `sheetView` of each sheet. For frozen panes it takes the selection of the active pane. The writer restores those fields; a book without them keeps the previous writer's first-visible-sheet default. Unknown, hidden or out-of-range saved active tabs are left out. Malformed optional view fields are left out one field at a time. Other windows, inactive panes' selections and gridline settings are outside this model, and so is a group of selected tabs: it is read as the active sheet alone, so a book never opens with its sheets grouped, where an edit goes to all of them.

Views change neither values nor `inputHash`. `getLayout()`, `getBook()` and snapshots retain them in every kind of session. Renaming the active sheet updates its name; deleting or hiding it chooses the first visible sheet. The host keeps live selection and scrolling itself and overlays `activeSheet` and `view` on the book before saving. They have no `setLayout` edit, so moving around a grid creates no undo steps.

#### How it is tested

By round trips through the writer and the reader, in every kind of session; not measured in Excel.

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

A sheet's table parts become its `tables`: the name formulas use, the range, the header and totals rows, each column with its calculated column formula and its totals function or label. Structured references come in the stored spelling, which is a file's but for a column name that starts with `@` (`Sales[@at]` in a file is `Sales['@at]`). A header cell that does not hold its column's name as text is given it, and a table the model cannot hold, over another or over a merged area, is left out; both are reported, and so is what the engine cannot take of a calculated column formula, as of a cell's (at the column's first cell). A calculated column formula that differs from its cells' formulas is kept as Excel keeps it: a row inserted into the table gets it. [docs/tables.md](tables.md) has the rules.

### The locale

xlsx has no place for the book's locale. The writer keeps it in a custom document property named `xlcalc.locale`, and the reader reads it back. Excel keeps the property when it saves the file. A file without it gives a book without a locale, that is en-US.

### What is reported

| code | what the book cannot hold |
|---|---|
| `richText` | formatted runs of text |
| `userFunction` | a user-defined function (`_xludf.`) |
| `table` | a table left out (over another table or a merged area, its name taken), a header cell given its column's name, a structured reference without a table name |
| `externalLink` | a DDE or OLE link, or a link whose part is missing, kept as a book without sheets; a formula over a link the file does not hold or over such a link |
| `formula` | a formula the engine does not parse, kept as written |
| `errorConstant` | an error stored as a value, read as a formula of it |
| `errorValue` | an error the model has no code for, such as `#GETTING_DATA` |
| `dataTable` | a what-if data table, read as its stored values |
| `name` | built-in names, hidden names, names of a sheet left out; a linked book's name of a sheet it does not have |
| `calculation` | the manual mode, automatic but for data tables, precision as displayed, an iteration count or change out of range |
| `style` | pattern and gradient fills, diagonal borders, rotation, shrinking, sub- and superscript, protection |
| `defaultFont` | a default font other than Calibri 11 |
| `defaultWidth` | a default column width other than 8 digits |
| `numberFormat` | an East Asian built-in format (50 to 58), or an unknown one |
| `hidden` | a file whose every worksheet is hidden (beside chart sheets): its first worksheet is read as visible |
| `autoFilter` | an autofilter's range, criteria and sort, a table's criteria, an advanced filter |
| `sheetKind` | chart, dialog and macro sheets, left out |
| `cellMetadata` | a rich value in a cell, such as a data type or a picture, or among a link's stored values |
| `comment` | a thread of comments, read as a note; a note of runs with different looks, read as its plain text; a note shown all the time, read as one shown on hover |
| `hyperlink` | a hyperlink whose address is no external relationship of the sheet; one over more than 4,096 cells, read on its cells that hold something |
| `frozenPanes` | split panes that are not frozen; panes frozen over a sheet scrolled down or right, read from the first row and column |
| `conditionalFormat` | a rule of a kind the model does not have, a pivot table's rules, a rule without its formulas or thresholds; what a rule's differential format holds beyond the model's style (alignment, protection, a font's name and size, pattern and gradient fills); an average rule with both equal and deviations, read with equal |
| `dataValidation` | a rule without a range of the sheet or the formula its type needs, or of an unknown type, left out; cells a rule shares with an earlier one, which keep the earlier one; an IME mode; a list whose items hold commas (`x12ac:list`), read split at them |
| `outline` | automatic styles of an outline |
| `drawing`, `pivotTable`, `protection`, `print`, `macros`, `sparkline`, `connection` | parts of the file the model does not have |
| `locale` | a locale the package does not have |
| `part` | a missing sheet part, a cell outside the sheet |

Gridlines and other view settings outside `SheetView` are not reported. A page setup is reported only when someone made one, for example a landscape page or a fit to width.

### The language of the messages

An issue's `message` is English by default. `readXlsx(data, { locale })` and `writeXlsx(snapshot, { locale })` write it in the language of a locale: the tag of a loaded locale (`'ru-RU'` once `registerLocale(ruRU)` ran) or a locale object (`ruRU`), which needs no loading. A locale object without messages of its own speaks as the loaded locale of its tag, else in English. The code, sheet, cells and name are the same in every language, and so are the issues and their order.

| code | en-US | ru-RU |
|---|---|---|
| `richText` | formatted runs of text are read as plain text | Форматированный текст прочитан как простой |
| `style` | a style's text rotation is not in the model | Поворот текста в стиле не поддерживается |
| `calculation` | the maximum of iterations "0" is not one of 1 to 32,767: the default, 100, is read | Предельное число итераций «0» вне пределов от 1 до 32 767: прочитано значение по умолчанию, 100 |
| `noCachedValues` | no result was given: the file stores no values of formulas, and Excel computes them when it opens the file | Результат не передан: файл не хранит значения формул, Excel вычислит их при открытии файла |

The host passes the language the person reading the messages chose. A file does not say it: Excel keeps no language in xlsx, and the locale the writer keeps in `xlcalc.locale` is the book's, the rules it computes by, which a person working in another language may open too. So neither function takes the language from the book.

The messages are part of the locale's module, beside its other messages ([docs/locale.md](locale.md#messages)), and the English ones part of the subpath, so the subpath carries no locale but en-US, and a locale without its own text of an issue gives the English one: a host that shows Russian loads `xlcalc/locales/ru-RU` as it does for books in ru-RU. One part of one message stays English: why a table is left out, which is the check of tables' own text (`Таблица «Over» не прочитана, модель не может её хранить (B1:C2 overlaps the table T)`).

## Writing

`writeXlsx(snapshot, opts)` resolves to `{ data, issues }`: the file's bytes and what the file cannot hold. The same book and result give the same bytes. The snapshot is `{ book, result }`: the book and the values of its formulas and of the cells their arrays spilled into, from `compute`, from a session (`session.getSnapshot()` gives both in one call, so no edit comes between them) or from the `cached` of a file read.

| option | |
|---|---|
| `values: 'require'` | the default with a result: every formula and every cell of an array has its value in the result, else the result is not the book's and the call throws `InputError` |
| `values: 'preserve'` | the values the result has; a formula without one is written without it and reported (`noCachedValues`) |
| `values: 'omit'` | the default without a result: no values of formulas, and Excel computes them when it opens the file; reported (`noCachedValues`) when the book has formulas and `values` was not given |
| `mode: 'formulas'` | the default: formulas, with the values `values` asks for |
| `mode: 'values'` | values only: a formula's cell gets its value from the result, which is then required |
| `locale` | the language of the issues' messages, as `readXlsx` takes it: English by default ([The language of the messages](#the-language-of-the-messages)) |

```js
import { compute, createSession } from 'xlcalc'
import { readXlsx, writeXlsx } from 'xlcalc/xlsx'

const book = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*10' } } }] }
const session = createSession(book)
const file = await writeXlsx(session.getSnapshot())
const values = await writeXlsx({ book, result: compute(book) }, { mode: 'values' })
const later = await writeXlsx({ book }, { values: 'omit' })
const a2 = async (f) => {
  const r = await readXlsx(f.data)
  return [r.book.sheets[0].cells.A2, r.cached.sheets[0]?.cells.A2?.value]
}
// the formula with its value; the value alone; the formula, for Excel to compute
console.log(await a2(file), await a2(values), await a2(later))
// → [ { formula: 'A1*10' }, 20 ] [ { value: 20 }, undefined ] [ { formula: 'A1*10' }, undefined ]
// without a result and without values, the file says so
console.log(later.issues, (await writeXlsx({ book })).issues.map((i) => i.code))
// → [] [ 'noCachedValues' ]
```

A session is no snapshot: an asynchronous session is a proxy of a thread, and its book and result are taken there by `await session.getSnapshot()`. Given one, a result alone (`writeXlsx(compute(book))`, told by its hashes) or a result of another shape, `writeXlsx` throws `InputError` saying what to pass. The older form, the book first and its result in the options (`writeXlsx(book, { result, mode })`), is an `InputError` too: a book is not a snapshot.

### Formulas

Every formula is written as a dynamic array formula, as Excel 365 saves one: with cell metadata and the range its array spilled into. The engine computes every formula by the rules of dynamic arrays, and so does Excel with this mark. Excel 2019 and older show such formulas as array formulas in braces.

The spilled cells are written with their values. `#SPILL!` and `#CALC!` are written as Excel 365 writes them, so Excel reads them back as those errors.

Functions newer than Excel 2007 get their prefix (`_xlfn.XLOOKUP`, `_xlfn._xlws.SORT`), LET's names become `_xlpm.x`, `E2#` becomes `_xlfn.ANCHORARRAY(E2)` and `@x` becomes `_xlfn.SINGLE(x)`. Sheet names are always quoted.

A cell with `array` is written as an array formula of the old kind over its range, without the cell metadata, as Excel writes a formula entered with Ctrl+Shift+Enter. A formula in a table's totals row is written as a plain formula, the only kind Excel opens there. A name's formula is written as the book has it: a constant, a formula or a union. The list of prefixes covers every function Microsoft lists as newer than 2007, not only the whitelist.

A file with formulas asks Excel to compute the book when it opens it, so Excel shows its own values over the stored ones, whichever values the file stores, but for the cells a circular reference stops, which keep the stored ones: a formula with `held` stores it where the result gives no value, with `values: 'omit'` too ([cycles.md](cycles.md)). The stored values are what a reader that does not compute shows, such as a preview or a spreadsheet library, and what `readXlsx` gives as `cached`.

### Tables

Each table is written as a table part, as Excel writes one: its name, range, header and totals rows, an autofilter over the header and data, its columns with their calculated column formulas and totals. A formula of the totals row that no totals function of its column gives is recorded as the column's custom totals formula, without which Excel does not open the file, and a table's name alone in a defined name is written `Sales[]`.

### Styles, columns and rows

Styles, number formats, widths, heights, hidden rows and columns, outline levels and merged areas are written as the book has them. The default font is Calibri 11, and a column without a width gets Excel's standard width of 8 digits. Every cell is written with the look it shows, its row's or column's where it has none of its own, as Excel writes it ([The cells of formatted rows and columns](#the-cells-of-formatted-rows-and-columns)).

The cells of a sheet's looks are written as Excel stores them, each a cell with the cell format of the whole look it shows (`<c r="B2" s="3"/>`), row by row with the entries, without an object per cell. The cell formats are numbered as the cells come, the looks' cells among the entries as the entries of a file read come, so a book read from a file and written again gives the file the reader gave when such cells were entries. A sheet without looks is written as it is given, its entries in their order.

A built-in format given by its number is written as that number, and Excel shows it by its own locale. A code that a locale shows otherwise than the built-in format it equals, such as `m/d/yyyy` (format 14), is written as a code of the file's own, so every locale shows what `numberFormat` says. A code that every locale shows alike, such as `0.00`, is written by its built-in number, as Excel writes it, and comes back as that number, `2`.

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
| `errorValue` | `#CYCLE!`, which a name that uses itself gives, is written as 0, as Excel stores a circular reference, among a link's stored values too |
| `textLength` | text over 32,767 characters is cut, a note's too, a hyperlink's tooltip over 255, a rule's titles over 32, its input message over 255 and its alert's text over 225 |
| `dataValidation` | a rule whose formula does not parse, or names a sheet, a table or a column the book does not have, is left out |
| `formula` | a formula that does not parse, or names a sheet, a table, a column or a link to another workbook the book does not have, is written as its value |
| `noCachedValues` | the file stores no values of formulas, as no result was given, or a formula has no value in the result (`values: 'preserve'`) |
| `name` | a name Excel does not take, given twice, of a missing sheet, or whose reference does not parse or names a table, a column or a link the book does not have, is left out |
| `autoFilter` | the filter mode of a sheet whose tables have no header row and meet its cells is left out |
| `conditionalFormat` | a rule with a formula the file cannot hold (it does not parse, or names a sheet, a table or a column the book does not have) is left out |

## The cells of formatted rows and columns

A book gives a cell without a style or a format of its own its row's, else its column's, the style and the format each separately: `{ value: 46096 }` in a column of `dddd, mmmm d, yyyy` shows `Sunday, March 15, 2026`. Excel reads a file otherwise. The writer and the reader turn one into the other, so a book written and opened in Excel shows what the engine shows, and a file Excel saved reads into a book that shows what Excel showed.

### Cells the file stores

Excel shows a cell the file stores (a `<c>` element) by the cell's own cell format alone, whatever its row and column have. So the writer writes each cell with the whole look it shows. A value in that column of dates is written with the column's cell format, the one Excel gives a value typed there, and a cell with a bold style of its own in a filled column of dates is written bold, a date and without the fill, as the book shows it.

```js
// columns: { B: { numberFormat: 'dddd, mmmm d, yyyy', style: 1 } }, cells: { B2: { value: 46096 } }
// <col min="2" max="2" … style="1"/> and <c r="B2" s="1"><v>46096</v></c>
```

The reader keeps of a stored cell only what its row or column does not give it: `B2` of that file comes back as `{ value: 46096 }`, without a format of its own. Where a cell stores less than its line gives, the reader spells the default out: a cell stored in General in a column of dates gets `numberFormat: 'General'`, and one without a fill in a filled column the book's plain style, a style with nothing set.

### Where a row crosses a column

An empty cell the file does not store shows its row's whole cell format when the row has one (`customFormat`), else its column's. So where a row with only a style crosses a column with only a format, Excel shows the empty cell, and a value typed there, without the column's format: `46096` typed where a bold row crosses a column of percents shows `46096`, bold, where the book shows `4609600.00%`.

Excel never leaves such a crossing empty itself: formatting a column across a formatted row, or a row across a column, it writes the cells where they cross, empty, with both lines' formats. The writer does the same: where a row with a look crosses a column with one and Excel would show the empty cell otherwise than the book, it writes that cell with the look the book shows. The reader reads a file without these cells as Excel shows it: a row with a cell format of its own that leaves out a crossing column's style or format, where a cell of that column is missing, gets the plain style or General.

```js
// <row r="6" s="1" customFormat="1"> (bold), <col min="7" max="7" style="3"/> (0.00%), no <c r="G6">
rows: { 6: { style: 0, numberFormat: 'General' } }
```

### Lines Excel formats

Excel, formatting a line across a formatted one, merges their formats field by field: a row made bold across a filled column of dates gives the cells where they cross bold, filled and dates. A book's style is one whole, and `setLayout` sets a line's style alone: where a row with a style crosses a column with one, the book shows the row's. `patchCells` formats lines as Excel does and gives those cells a style of their own ([Editing a grid](edits.md#whole-rows-and-columns)).

### Checked against Excel

Measured in Excel for Mac and for Windows: Excel opened the written books, formatted rows, columns and cells through its object model and showed what the engine shows.

## Conditional formats

A sheet's conditional formats (`conditionalFormats`) go both ways. The reader reads the rules of `<conditionalFormatting>` with their differential formats (`<dxfs>` of styles.xml) and Excel 2010's rules of the sheet's extension list (`x14:conditionalFormatting`), and gives them in priority order; the writer writes them so, each rule's priority by its place in the list. Excel 2010's part of a data bar (its fill, border, direction, negative bars and axis) is joined to its 2007 rule by their id; the icon sets `3Stars`, `3Triangles` and `5Boxes`, custom icons and formulas over other sheets are Excel 2010's rules alone, with their style inline. A text, blank or period rule always has its formula in a file; the reader leaves out the one Excel writes for it, which the model makes itself. Ranges come as the model writes them: Excel's `A1:A1048576` is `A:A`, and a range another writer spells in lower case or with `$` is read as Excel reads it, `$a$1:b2` as `A1:B2`. A dxf's number format is read by its code, as Excel saves `0.00%` under another id and shows a dxf by its code in every locale (`m/d/yyyy` under the id 14 is `3.15.2026` in ru-RU). A dxf's fill is read as Excel reads it, otherwise than a cell's: a solid fill shows its `bgColor` alone (white for one that names no colour), a fill without a pattern type is solid, and an empty one changes nothing. `patternType="none"` takes the cell's own fill away, and a solid fill without `bgColor` keeps it (white on a cell without one); the model can say neither, so such a rule is read without a fill, and that is reported. Excel opens no file whose dxf holds a gradient beside a pattern; the reader reads the pattern and reports the gradient. [docs/conditional-formats.md](conditional-formats.md) has the rules, what a cell shows by them and how it was checked.

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

A note has its author and its text, lines separated by `\n`. Excel keeps notes in a comments part with a VML drawing that shows them, and the reader takes the comments part. The text is read as plain text: a note made in Excel's window has its author's name in bold at the start, and such formatting is reported. A note shown all the time comes back as one shown when the pointer rests on its cell, also reported. The size and place of a note's box are not kept. A note's cell is read as Excel reads it: `A01`, `$A$1`, `a1` and `A1:A1` are all A1; a note on a range of cells is left out and reported, as Excel does not open such a file.

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

### Checked against Excel

Measured in Excel for Mac and for Windows: both opened every written book with notes, hyperlinks and frozen panes (but one with a note over a range, which neither opens) and saved what the engine holds.

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

A file keeps a rule in `<dataValidations>` with xlsx's defaults false: a rule without `allowBlank`, `showInputMessage` or `showErrorMessage` is read with `allowBlank: false`, `showPrompt: false` or `showError: false`, as the model's defaults are those of a rule Excel makes. `showDropDown="1"` hides a list's dropdown and is read as `dropdown: false`. A rule whose formulas refer to another sheet Excel writes in the sheet's extension, `<x14:dataValidations>`, which the reader reads beside the others. Formulas come in the engine's spelling, as a cell's do. Excel saves a whole column as `F1:F1048576`, read as `F:F`; a range in lower case or with `$` is read as Excel reads it, `$a$1:b2` as `A1:B2`.

Two rules over one cell are not in the model. Excel writes them when a move puts a rule over a cell that keeps another one in its range, and shows the one written first, which is the one the reader keeps; the later rule loses the cell, and that is reported. A list typed where the comma is no list separator (in ru-RU) can hold an item with a comma: Excel keeps that list in an extension (`x12ac:list`), which the reader does not read, and gives the list split at its commas as a fallback, which it reads and reports.

The writer writes every rule in `<dataValidations>`, references to other sheets too, which Excel opens. A rule's titles over 32 characters, its input message over 255 and its alert's text over 225 are cut, as Excel does not open longer ones.

### Outlines

A row or column's `outlineLevel` and `collapsed` are `<row>`'s and `<col>`'s attributes, the summary settings `<sheetPr><outlinePr summaryBelow summaryRight>`. The writer writes the deepest levels into `<sheetFormatPr outlineLevelRow outlineLevelCol>`, where Excel reserves room for the outline's buttons, and the reader takes them from the lines. The writer writes `collapsed` as Excel keeps and saves it, on the summary line of a group whose lines are all hidden, whatever marks the book holds; the reader reads a file's marks as they are ([docs/outline.md](outline.md#in-xlsx-files)). Automatic outline styles (`applyStyles`) are reported.

### Checked against Excel

Measured as [Data validation](validation.md) says: both Excels opened every written book and read the rules, each cell's check and the outlines the engine holds.

## Links to other workbooks

A formula may read another workbook: `[1]Sheet1!A1+1`, `SUM('[1]My Sheet'!A1:B9)`, `[1]!Rate`. Excel keeps in the file the values of the other book it read last, and computes such formulas from them when it opens the file without updating its links. The reader gives these links as the book's `externalBooks`, so the engine computes the formulas as Excel does, and the writer writes them back, so a file saved from a grid keeps its links.

```js
externalBooks: [
  {
    target: 'Book2.xlsx',
    sheets: [
      { name: 'Sheet1', cells: { A1: 5, A2: 'east', A3: true, A4: { error: 'NA' }, A5: null, A6: '' } },
      { name: 'Rates', cells: { B1: 0.2 }, refreshError: true },
      { name: 'Old' },
    ],
    names: [{ name: 'Rate', ref: 'Rates!$B$1' }, { name: 'Local', ref: 'Sheet1!$C$2', sheet: 0 }],
  },
]
```

Excel shows a cell the link holds as its stored value, `null` as an empty cell and `''` as empty text. A cell the link does not hold is empty, or `#REF!` on a sheet with `refreshError`, the mark Excel puts on every sheet of a link when a formula over it is entered. Every cell of a sheet without `cells` is `#REF!`. [docs/links.md](links.md) has how the engine computes and shows such formulas.

### Reading

The first link the workbook lists (`<externalReferences>`) is `[1]`, the second `[2]`, whatever its part is named.

| in the file | in the book |
|---|---|
| the target of the link's own relationship, of any kind (`xlPathMissing` too): `Book2.xlsx`, `../data/Book3.xlsx`, `/xlcalc/work/Book2.xlsx`, `file:///C:/data/Book4.xlsx`, `https://example.com/Book6.xlsx` | `target`, as the file holds it (`Sub%20Folder/Book7.xlsx` keeps its `%20`) |
| Excel's other addresses of the book (`xxl21:alternateUrls`) | left out: the target stays the relationship's |
| `<sheetName val="Sheet1"/>` | a sheet of `sheets`, in the file's order |
| `<definedName name="Local" refersTo="=Sheet1!$C$2" sheetId="0"/>` | a name of `names`: `ref` without `=`, `sheet` the index of the sheet whose own name it is |
| `<sheetData sheetId="0" refreshError="1">` | that sheet's `cells` by address, and `refreshError: true` |
| a sheet without `<sheetData>` | a sheet without `cells` |
| `<cell r="A1"><v>5</v></cell>` (also `t="n"`) | `5` |
| `t="str"` (and `t="s"` or `t="inlineStr"`, which another writer may write) | text; `<v/>` is empty text `''`, as Excel writes the value of `=""` |
| `t="b"`, `t="e"` | a boolean, an error (`{ error: 'NA' }`) |
| `<cell r="A5"/>` | `null`, a cell stored empty |

A link to another program's data (DDE) or to an object (OLE), and a link whose part is missing, are not in the model. Each keeps its number as a book with an empty target and no sheets, so the links after it keep theirs, and each is reported (`externalLink`). A stored value with a rich value (`vm`, a data type or #SPILL! that Excel 365 stores as `#VALUE!`) is read as the value the cell stores, which is what Excel shows, and reported.

Formulas and names over links come as the file writes them: `[1]Sheet1!A1+1`, `'[1]My Sheet'!A1`, `[1]!Rate` (a name of the book), `[1]Sheet1!Local` (a name of its sheet), `[1]Sheet1!#REF!` (a reference into it that Excel lost); this book's name `Ext` over `[1]Sheet1!$A$2` keeps its reference. A shared formula's copies move their relative references into the link's sheet, as a paste does: `[1]Sheet1!A1` copied one row down is `[1]Sheet1!A2`. Copies the file holds are no edits, so they put no mark on the link. A formula over a link the file does not hold (`[3]Sheet1!A1` in a file of two links), or over a DDE or OLE link, is read as written and reported.

### Writing

Each link of `externalBooks` is written as Excel writes it: a part `xl/externalLinks/externalLinkN.xml` with the sheets' names, the names (`refersTo` with its `=`) and the stored values; its relationship to the other book, external, with the target as the book holds it; the workbook's `<externalReferences>` after its sheets. A number is written without a type, text as `t="str"` (`''` as `<v></v>`), a boolean as `t="b"`, an error as `t="e"`, `null` as `<cell r="A5"/>`, the mark as `refreshError="1"`, and a sheet without `cells` gets no data. A book without links is written as before.

A formula or a name over a link is written in the file's spelling, the sheet in quotes only where its name needs them: `'[1]Sheet1'!A1` of the book is `[1]Sheet1!A1` in the file, `'[1]My Sheet'!A1` stays quoted, `[1]!Rate`, `[1]Sheet1!Local` and `[1]Sheet1!#REF!` as they are. A reference by the other book's file name, as a person types it (`[Book2.xlsx]Sheet1!A1`), is written with its link's number. A formula over a link or a link's sheet the book does not have is written as its value, and a name over one is left out; both are reported.

### What is lost

- Excel's other addresses of a book (`xxl21:alternateUrls`): the writer writes the target alone, which Excel opens.
- A stored value's rich value (`vm`).
- #SPILL! and #CALC! among a link's values, written as `#VALUE!`. Excel keeps no other there: it repairs a file that holds their text, and shows and saves `#VALUE!` for them stored with a rich value, as a sheet's cells store them.
- The mark of a sheet without `cells`, which changes nothing: every cell of such a sheet is `#REF!` with the mark or without.
- A DDE or OLE link, but for its number.

Excel itself leaves out, when it saves, the values of the other book's cells no formula reads and a link no formula refers to, numbering the links after it anew in the formulas.

### Checked against Excel

Measured in Excel for Mac and for Windows, which opened the files without updating their links.

## Speed

Measured on an Apple Silicon Mac with `test/fixtures/xlsx-bench-child.ts`, the median of three runs. The ranges are over several runs on a busy machine.

| book | file | read | write |
|---|---|---|---|
| 20,000 rows × 40 values | 3.0 MB | 0.7 to 1.05 s | 1.2 to 1.8 s |
| the synthetic book, 3,400 rows, 3,900 formulas with values | 0.41 MB | 0.1 s | 0.1 to 0.15 s |
| a book of the wide shape, 5.6 million cells of a look alone (`bench/memory/wide-shape.ts`) | 15 MB | 1.5 s | |

A read book of 800,000 values holds about 100 MB of heap; the book of the wide shape about 1 MB, and a session over it opens in 7 ms (6.2 s and 815 MB, then 10.5 s and 545 MB, when such cells were entries). `FORMULA_BENCH=1 npx vitest run test/xlsx-budget.test.ts` checks that the large book reads in 2 s at most.

## How the code is organised

`src/xlsx/` holds the code, and `src/xlsx.ts` is the subpath's entry.

- `zip/` reads and writes the archive. Inflating and deflating go through `DecompressionStream` and `CompressionStream`.
- `xml/` holds the text layer: entities, Excel's escapes, and elements found by regular expressions.
- `read/` reads the package, the workbook, the sheets and their cells, the cell metadata, the notes, threads, hyperlinks and frozen panes (`parts.ts`), conditional formats (`conditional.ts`), data validation (`validation.ts`) and the links to other workbooks (`links.ts`); `refs.ts` reads the references of a rule's cells and a note's cell as Excel reads them.
- `styles/` turns styles.xml into the style table and the differential formats (`dxf.ts`), whose fills one function reads (`readFill`), and holds the codes of the East Asian built-in formats, the codes the writer stores by number, and the colours.
- `formulas/` changes formulas between the file's spelling and the engine's. `prefixes.ts` is generated by `bench/coverage/xlsx/generate.ts`, so the reader needs neither the harness nor the engine's function definitions. Formulas of the old kind go through `fromLegacy` (`src/legacy/`), whose table of argument classes needs no function definitions either.
- `write/` writes the workbook, the sheets, styles.xml with its differential formats, the notes, hyperlinks and frozen panes (`parts.ts`), conditional formats (`conditional.ts`), data validation (`validation.ts`) and the links to other workbooks (`links.ts`).

Reading uses the parser and copy and paste of the structure operations. It does not load the engine's functions or evaluator, and a minified browser bundle of the subpath is about 175 KB. The issues' messages are kept as a key and parameters (`src/xlsx/issues.ts`) and written in a locale when listed: a locale's text is in its file beside its other messages (`src/locales/ru-RU.ts`), the English one in the subpath (`src/xlsx/messages.ts`), so that `compute`, the root and `xlcalc/text` do not carry it.

## Checked against Excel

Excel 16.113 for Mac and Excel for Windows opened every file the writer wrote without the repair window, computed it and saved it, and the reader read the files Excel saved back as Excel read them.

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
- A gradient fill is read as no fill; Excel shows the gradient. A conditional format's fill of `patternType="none"`, which takes a cell's own fill away in Excel, and its solid fill without a colour, which Excel shows white on a cell without a fill, are read as none. Each is reported.
- A list of data validation whose item holds a comma (typed in ru-RU) is read split at the comma. Excel saves a list's reference to a deleted sheet as `#REF!` alone, which the engine keeps as `#REF!$A$1:$A$4`.
