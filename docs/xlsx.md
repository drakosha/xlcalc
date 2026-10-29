# Reading and writing xlsx

The `xlcalc/xlsx` subpath reads an xlsx file into a workbook the engine can calculate, and writes a workbook with its calculated values back to xlsx. It has no dependencies and runs wherever the web platform's deflate streams exist: Node 20.19 or newer and current browsers. Neither function calculates anything: reading gives you the values Excel stored alongside the workbook, and writing takes the result you calculated.

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

`readXlsx(data)` from `xlcalc/xlsx` reads a file's bytes (a `Uint8Array`, a Node `Buffer` or an `ArrayBuffer`: `await readFile(path)`, `await file.arrayBuffer()`) into a workbook. Alongside the workbook it returns `cached`, the values Excel stored for formulas and spilled cells in the shape of a result, and `issues`, everything the workbook cannot represent, such as rich text, with the sheet and cells. `compute(read.book, { cached: read.cached })`, and `createSession` with the same option, open a workbook with a circular reference as Excel opens the file ([Circular references](cycles.md)); nothing else is taken from `cached`. `writeXlsx({ book, result })` writes a workbook with the values of its formulas: `result` comes from `compute`, from a session (`session.getSnapshot()` returns the workbook and its result in one call), or is the `cached` of a file that was read. Formulas are written as dynamic array formulas with their spills, and `mode: 'values'` writes values only.

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

`values` controls which formula values the file stores. With a result, it is `'require'`: every formula and every cell of an array must have its value in the result, otherwise the result does not belong to the workbook and `writeXlsx` throws `InputError`. `'preserve'` writes the values the result has and reports the formulas without one (`noCachedValues`). Without a result, it is `'omit'`: the file stores no values, which is reported as `noCachedValues` when the workbook has formulas, but not when `values: 'omit'` is requested explicitly. A file with formulas asks Excel to calculate them when it opens the file, whichever values it stores. A formula stopped by a circular reference stores its `held` value where the result has none (with `omit` too), so Excel opens the file with that value. A session is not a snapshot (an asynchronous one lives in a thread): pass `session.getSnapshot()`, awaited for an asynchronous session. Nor is a result on its own, which has no workbook: `writeXlsx(compute(book))` throws `InputError` with `pass { book, result }`.

Both functions return promises: the archive is inflated and deflated by the platform's streams. The following round-trip: hidden and very hidden sheets (`hidden: true` and `hidden: 'very'`; without `activeSheet`, the writer makes the first visible sheet the active one), styles, number formats (built-in ones by their number), widths, heights, hidden rows and columns, outline levels, merged areas, notes, hyperlinks, frozen panes, conditional formats (including Excel 2010's rules), data validation rules, names (hidden ones too), tables, a sheet's filter mode, the date system, iterative calculation, and the workbook's locale. A cell without its own style or number format is written with its row's, otherwise its column's, as Excel writes a value typed there, so a date in a column of dates displays as a date in Excel, and it reads back without one. A cell the file stores with formatting only, as Excel stores each cell of a formatted area, is read into its sheet's `looks` ([Formatting of many cells](grid.md#formatting-of-many-cells)), and the cells in `looks` are written back as such cells. A group's collapsed marker is written as Excel keeps and saves it: on the summary row of a group whose rows are all hidden, including rows hidden through `setLayout`, and not on one whose rows are visible ([docs/outline.md](outline.md#in-xlsx-files)).

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

A comment thread, as Excel 365 creates them, is read as a note: its first author, and every comment's text with the authors of the replies. [docs/xlsx.md](xlsx.md) lists what is read, written and reported, and how it was checked against Excel.

An issue's `message` is for users, in English by default. The `opts.locale` of either function writes it in another language: the tag of a loaded locale or a locale object, as other functions accept a locale. The code, sheet, cells and name of an issue stay the same, so a program reads them the same way in every language.

```js
import { registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'
import { readXlsx, writeXlsx } from 'xlcalc/xlsx'

registerLocale(ruRU)
const book = { locale: 'ru-RU', sheets: [{ name: 'План', cells: { A1: { value: 2 }, B1: { formula: 'A1*2' } } }] }
const { data, issues } = await writeXlsx({ book }, { locale: 'ru-RU' })
console.log(issues[0].code, issues[0].message) // → noCachedValues Результат не передан: файл не хранит значения формул, Excel вычислит их при открытии файла
console.log((await writeXlsx({ book })).issues[0].message) // → no result was passed: the file stores no formula values, and Excel calculates them when it opens the file
console.log((await readXlsx(data, { locale: ruRU })).book.locale) // → ru-RU
```

Pass the locale the user reading the messages has chosen. A file does not record it: Excel stores no language in xlsx, and the locale `xlcalc.locale` that the writer stores, which `readXlsx` returns as `book.locale`, is the workbook's, the rules it calculates by, and a user working in another language may open the same workbook. So neither function takes the language from the workbook.

## Reading

`readXlsx(data)` takes the file's bytes, a `Uint8Array` (a Node `Buffer` is one, as `readFile` returns it), another view of bytes, or an `ArrayBuffer` (`await file.arrayBuffer()` in a browser), and resolves to three things.

- `book` is a `BookInput`: sheets with their cells (including notes and hyperlinks), the formatting of their formatting-only cells ([Formatting-only cells](#formatting-only-cells)) and whether they are hidden; columns and rows (including their outline levels), merged areas, filter mode, frozen panes, conditional formats, data validation rules and outline settings; names, the style table, the locale, the date system, iterative calculation, and the links to other workbooks with their stored values ([Links to other workbooks](#links-to-other-workbooks)).
- `cached` holds the values the file stores for formulas and for the cells their arrays spilled into, by sheet in the workbook's order, in the shape of a result: `cached.sheets[0].cells.B1.value` like `compute(book).sheets[0].cells.B1.value`. A formula whose array spilled has `spill`, its range, and each other cell of the range has `spillOf`, the formula's address, as in a result. These are what Excel calculated when it saved the file, so you can compare them with the engine's result cell by cell, or write the file again with them (`writeXlsx({ book, result: cached })`). A workbook with a circular reference opens with them as Excel opens the file, with `compute(book, { cached })` or `createSession(book, { cached })`: the formulas the cycle stops keep the stored values ([cycles.md](cycles.md)).
- `issues` lists everything the workbook cannot represent, with the sheet and the cells or the name it concerns. Nothing is dropped without an issue. Each issue has a code for programs and a message for users, in English or in the language of `opts.locale` ([The language of the messages](#the-language-of-the-messages)).

A file that is not a zip archive, a damaged archive (an entry cut short, two entries of one name, bytes that do not match their CRC-32), a part that is not valid XML because a character reference in it is past U+10FFFF, an encrypted file, a ZIP64 archive and a file without a workbook part cause errors. Data that is not bytes, such as the file's path, causes an `InputError` that says to pass the bytes.

`opts.limits` bounds how large the file's parts may be once inflated: `entryBytes` for one part, 512 MiB by default; `totalBytes` for all parts read together, 1 GiB by default; and `ratio`, how many times its compressed size a part larger than 16 MiB may inflate to, 200 by default. A file beyond these limits causes an error, as a damaged file does. A host that reads files uploaded by users passes lower limits:

```js
const read = await readXlsx(data, { limits: { entryBytes: 16 * 2 ** 20, totalBytes: 64 * 2 ** 20 } })
```

A zip archive can contain a part that inflates to a thousand times its compressed size: a 400 KB file can hold 400 MB of spaces. The reader stops inflating a part as soon as it exceeds the size the file declares for it, so a part containing more than it declares is rejected without using more memory than it declared. A part declared larger than `entryBytes`, a part that would take the total past `totalBytes`, and a part larger than 16 MiB declared at more than `ratio` times its compressed size are rejected before they are inflated: the parts of xlsx files inflate some 5 to 40 times, and a zip bomb's a thousand times, so a 400 KB file that honestly declares 400 MB of spaces is rejected immediately. Two archive records over the same bytes, which let a small file inflate many times over, are rejected too. A part larger than about 512 MiB is not read whatever the limits: the reader reads a part as a single JavaScript string, and Node and Chromium support strings of at most 536,870,888 characters. The largest part among the files measured here is about 112 MiB, the sheet of the wide-shape workbook ([Speed](#speed)).

### Cells

Text, numbers and booleans become values. Dates stay numbers, as Excel stores them, and their number format makes them display as dates. The ISO dates of Strict files (`t="d"`) become serial numbers too.

Text keeps its spaces and line breaks. Excel's escapes for characters XML cannot represent are decoded: `_x000B_` is a vertical tab, and `_x005F_x0041_` is the literal text `_x0041_`. A CDATA section, which some other writers use for text with `<` or `&`, gives its content as written. Rich text is read as its plain text and reported. Phonetic readings of East Asian text are not part of the text.

The model has no error values, so an error stored as a cell's value is read as a formula consisting of that error: a cell containing `#N/A` gets the formula `#N/A`.

A cell is placed by its address, or after the previous cell when it has none. A `dimension` that does not match the cells is ignored.

### Formulas

Formulas are converted to the engine's spelling.

| in the file | in the workbook |
|---|---|
| `_xlfn.XLOOKUP(…)`, `_xlfn._xlws.FILTER(…)` | `XLOOKUP(…)`, `FILTER(…)` |
| `_xlfn.LET(_xlpm.x,1,_xlpm.x+1)` | `LET(x,1,x+1)` |
| `_xlfn.ANCHORARRAY(E2)` | `E2#` |
| `_xlfn.SINGLE(A1:A3)` | `@A1:A3` |
| `Jan:Mar!A1`, `'Jan 1:Mar 2'!A1` | the same reference across sheets |
| `_xludf.MYFN(1)` | `MYFN(1)`, reported |

A prefix counts only at the start of a word: a name or a table name continues through `?` as through letters, so `a?_xlfn.b` is the name `a?_xlfn.b` and `T?[Qty]` refers to the table `T?`, as Excel loads them.

A shared formula is written once in a file and used by a block of cells. Each of these cells gets its own formula, adjusted as copy and paste would adjust it: `A1*2` in B1 becomes `A2*2` in B2. A master cell below or to the right of its copies works too, and a reference pushed past the edge of the sheet becomes `#REF!`.

An array formula spills from its top-left cell. The other cells of its range hold only stored values in the file, so they go to `cached`, and they keep their own style in the workbook. Excel 365 marks its dynamic array formulas with cell metadata. An array formula without that mark is a legacy array formula, entered with Ctrl+Shift+Enter. Its cell gets the formula and `array`, the range it covers, and the formula is calculated over that range as Excel calculates it ([docs/legacy.md](legacy.md)).

A plain formula without the mark is a legacy formula too. Excel 365 calculates it with implicit intersection and displays `@` where it takes a single value from a range. The reader returns the formula Excel displays, using `fromLegacy`:

| in the file | in the workbook |
|---|---|
| `A1:A3*2`, `ABS(A1:A3)` | `@A1:A3*2`, `ABS(@A1:A3)` |
| `INDEX(A1:B5,0,1)` | `@INDEX(A1:B5,0,1)` |
| `span*1`, where `span` is a name that refers to a range | `@span*1` |
| `SUM(A1:A3)`, `SUMPRODUCT(A1:A3*2)`, `IF(A1>0,1,2)` | as written |

A formula that needs no `@` keeps its text as written. Excel 365 writes a dynamic formula as a plain one when the two kinds agree, so most formulas in a file saved by Excel 365 come through unchanged.

Excel 365 stores `#SPILL!` and `#CALC!` as `#VALUE!` with a rich value that names the actual error. The reader reads them as `#SPILL!` and `#CALC!`.

### Styles and number formats

Each cell format in the file becomes a style in the workbook's table and a number format. Equal styles are shared. A style holds a font, a solid fill, borders, alignment and protection ([Protection](#protection)). Its alignment holds the text's orientation too, `rotation`: degrees from -90 to 90, up for positive ones, or `vertical`, the letters one above the other (xlsx `textRotation`, 91 to 180 for the degrees down, 255 for the letters); Excel 16.0.20430 for Windows read each kind from files `writeXlsx` and `writeKept` wrote as the style holds it and saved it so (`bench/windows/run.ts fields rotation`). AutoFit fits rows and columns to turned text as Excel for Mac does ([Turned text](formats.md#turned-text)). Font fields equal to the workbook's default font are omitted, since the model takes them from the default. The default font is the file's first font (the Normal style's); one other than Calibri 11 is read into `BookInput.font` and reported (`defaultFont`), since the engine measures widths, fitted heights and `####` as for Calibri 11 whatever the font.

Colors become `#RRGGBB`. A theme color with a tint is calculated as Excel displays it, and the legacy palette or the workbook's own palette resolves indexed colors.

A fill is read as Excel displays it. A solid fill displays its `fgColor`; without one, or with one that names no color (`auto`, a system color), it is white, Excel's automatic color. A fill without a pattern type is no fill; when a fill has both a gradient and a pattern, the pattern counts. Gradients and patterns other than solid are not in the model: the cell is read without a fill, and that is reported.

A custom number format is read as its code. A built-in one is stored in the file as a number, and the reader returns that number, 1 to 49, which the workbook displays according to its locale, as Excel does: format 14 displays 2025-01-31 as `1/31/2025` in en-US and `31.01.2025` in ru-RU ([docs/formats.md](formats.md)). Formats 50 to 58, East Asian dates, are not in the model: they are read as the code Excel for Mac en-US displays, `m/d/yy`, and reported. That code is the Mac's short date, format 14; Excel for Windows displays them as its own format 14 too, `m/d/yyyy` in en-US and `ДД.ММ.ГГГГ` in ru-RU, so where the engine displays 2025-01-31 as `1/31/25` (`1.31.25` in ru-RU), Excel for Windows displays `1/31/2025` (`31.01.2025`), as the engine displays format 14 (`test/xlsx-win.test.ts`).

Excel for Mac saves some built-in formats other than by their number: the currency and accounting ones (5 to 8, 41 to 44) as codes in the locale it runs in, 23 to 26 as General, and the East Asian 27 to 36 as 14 and 21. A file saved by Excel contains them that way.

### Formatting-only cells

Excel stores each cell of a formatted area, including a row formatted out to column XFD, as a cell with a cell format and nothing else (`<c r="B2" s="3"/>`). Such a cell does not become an entry in `cells`: the reader puts its formatting into its sheet's `looks` ([Formatting of many cells](grid.md#formatting-of-many-cells)) as it reads the row, as runs of columns with the same formatting, and rows with the same runs share them. A cell with a value or a formula stays an entry with its own formatting, as before.

```js
// <row r="20"><c r="A20" s="1"/><c r="B20" s="1"/> … <c r="XFD20" s="1"/></row>, and so rows 21 and 22
looks: [{ range: 'A20:XFD22', style: 0 }]
```

So a sheet of the wide shape, 344 rows of such cells from A to XFD in two formats, the second in every 12th column, 5.6 million cells, is read into 5,571 rectangles and about a megabyte of heap, where it took 815 MB as entries. A cell that gets content later during reading (a shared formula's copy, a note, a hyperlink or a table's header) takes its formatting from the file as its own, and its position among the stored cells, as an entry with that formatting would; a cell an array spilled into keeps only its formatting, in `looks`. A hyperlink over more cells than are listed individually applies to the cells in `looks` within its range as it does to the stored ones.

### Columns, rows and merged areas

A column gets its width when the file sets it manually. A width Excel fits automatically is omitted, but a column width that AutoFit or typing set, which Excel saves with `bestFit`, is read with `bestFit: true` and written that way ([Fitting columns](formats.md#fitting-columns)). A row gets its height either way: one Excel fitted (AutoFit, or wrapping text typed into it), which the file stores without `customHeight`, is read with `bestFit: true` and written back without `customHeight` ([Fitting rows](formats.md#fitting-rows)). A hidden row or column gets `hidden: true` and keeps its size, which it returns to when unhidden. Excel treats a row of height 0 or a column of width 0 as hidden, whether or not the file marks it hidden, and so do the reader and the engine. A row or column with its own formatting gets its style and number format. Adjacent rows or columns with equal settings become one span, such as `B:D` or `5:9`.

```js
// <row r="5" ht="30" hidden="1" customHeight="1">, <col min="5" max="5" width="12" hidden="1" customWidth="1"/>
rows: { 5: { height: 30, hidden: true } }, columns: { E: { width: 12, hidden: true } }
```

A grouped row or column gets its `outlineLevel`, 1 to 7, and the summary row or column of a collapsed group gets `collapsed: true`; the group's rows or columns are hidden ([Data validation and outlines](#data-validation-and-outlines)).

A cell keeps only the part of its formatting that its row or column does not provide ([The cells of formatted rows and columns](#the-cells-of-formatted-rows-and-columns)): a date typed into a column of dates comes back as a value alone, and displays with the column's format.

A merged area that overlaps another is left out and reported.

### Sheet names

A file from another program may name a sheet as Excel would not: `History`, more than 31 characters, an empty name, a name another sheet has in another case, an apostrophe at either end, or one of `: \ / ? * [ ]`. The sheet is read under a name the workbook can hold, and reported with code `part`, the new name as `sheet` and the file's as `name`. Each of `: \ / ? * [ ]` and an apostrophe at either end becomes `_`, the name is cut to 31 characters, an empty name becomes `Sheet` and the sheet's place (`Sheet2`), and a name that is reserved (`History`) or taken by another sheet gets a number, as Excel numbers a copied sheet: `History (2)`, `s (2)`. The first of names that differ only in case keeps its name. Formulas, names, calculated columns, conditional formats and validation rules that refer to the sheet refer to its new name, as after `renameSheet`, and so do the locations of hyperlinks, which `renameSheet` keeps as written ([Hyperlinks](#hyperlinks)); a reference to a name several sheets share goes to the first of them. A reference to a sheet whose name holds `[` or `:` reads as a reference to another workbook or across sheets and is not rewritten. This rule is the reader's own: what Excel does when it opens such a file is not recorded here.

```js
// <sheet name="History" sheetId="1" r:id="rId1"/>, a formula History!A1*2 on another sheet
{ name: 'History (2)', cells: { … } }, { formula: "'History (2)'!A1*2" }
```

### Hidden sheets

A sheet marked `state="hidden"` in a file gets `hidden: true`, and one marked `state="veryHidden"` gets `hidden: 'very'`: a very hidden sheet is one that Excel's window does not offer to unhide, only its object model can. Formulas read hidden sheets like any other. A workbook keeps at least one visible sheet, so a file whose worksheets are all hidden next to a visible chart sheet, which the workbook leaves out, has its first worksheet made visible, which is reported.

```js
// <sheet name="Rates" sheetId="2" state="hidden" r:id="rId2"/>, <sheet name="Keys" sheetId="3" state="veryHidden" r:id="rId3"/>
{ name: 'Rates', cells: { … }, hidden: true }, { name: 'Keys', cells: { … }, hidden: 'very' }
```

### The saved active sheet and view

A grid can reopen a workbook on its saved sheet, with its saved selection, scroll position and zoom. `book.activeSheet` names the active visible sheet. Each sheet's `view` holds `activeCell`, `selection` (an array of A1 cells or ranges), `topLeft`, `zoom` (a percentage, 10 to 400), and `gridlines` and `headings`, false where View > Show hides the sheet's gridlines or its row and column headings (xlsx `showGridLines`, `showRowColHeaders`). The active sheet's tab is the selected one, and a sheet's `tabColor` (`#RRGGBB`, xlsx `<tabColor>`, a theme color read as its color) colors its tab.

```js
const book = {
  activeSheet: 'Plan',
  sheets: [{ name: 'Plan', cells: {}, view: {
    activeCell: 'C4', selection: ['C4:D5', 'F7'], topLeft: 'B3', zoom: 125,
  } }],
}
```

The reader takes `workbookView.activeTab` from the first workbook view, and the first `sheetView` of each sheet. It takes the selection of the active pane, the top left one where the file names none; a cursor on A1, which Excel saves as a pane's selection without a cell, is no saved cursor. For frozen panes it takes a scroll position past the split from the pane after it, where Excel keeps it and the writer writes it, so frozen panes and `topLeft` round-trip together. On a frozen axis the scroll counts from the first row or column: over columns A to C frozen, `topLeft: 'D10'` (the pane after the split not scrolled) reads back as `A10`, the same window. Frozen panes without a saved cursor are written with the cursor on A1, as Excel saves them. The writer restores those fields; a workbook without them keeps the writer's default of the first visible sheet. Unknown, hidden or out-of-range saved active tabs are left out. Malformed optional view fields are left out one field at a time. Other windows, the selections of inactive panes, and the gridlines' color are outside this model, and so is a group of selected tabs: it is read as the active sheet alone, so a workbook never opens with its sheets grouped, where an edit would apply to all of them.

#### A split window

Excel's View > Split divides the window into two or four panes that scroll apart, without freezing them. A sheet's `view.split` holds it: `x`, the vertical line's distance from the window's left edge, and `y`, the horizontal line's from its top, in points, at least one of them; `topLeft`, the top left cell of the pane after both lines; and `activePane`, the pane that holds the cursor and the selection (`topRight`, `bottomLeft`, `bottomRight`; the first pane, `topLeft`, without it).

```js
const sheet = { name: 'Plan', cells: {}, view: {
  topLeft: 'A1', activeCell: 'G21', selection: ['G21'],
  split: { x: 120, y: 60, topLeft: 'F20', activePane: 'bottomRight' },
} }
```

The four panes scroll together by rows and by columns: the top left one from the view's `topLeft`, the top right one from its row and the split's column, the bottom left one from the split's row and its column, the bottom right one from the split's `topLeft`. With one line there are two panes, the second scrolled on that line's axis. The points are the file's (xlsx `<pane xSplit ySplit>` in twentieths of a point): they count the row and column headings in and are points on the screen at the view's zoom, so a grid draws the vertical line at `x` less the width of its own row headings at its zoom, and the horizontal one at `y` less the height of its column headings. Excel for Windows's headings at 100% are 19.5 points wide and 15 high, Excel for Mac's 23 and 22, so the same file shows a little more of the sheet left of the line on Windows. Excel saves a line on a pixel of its screen: Excel for Windows at 96 DPI on the nearest 0.75 point (150.5 as 150.75), Excel for Mac on the whole point before it (150.5 as 150).

Frozen panes take the place of a split: a sheet with both is written with its frozen panes, and once `setLayout` unfreezes them the split is written again, as Excel restores a window split before Freeze Panes. A file's frozen panes that were split before (`state="frozenSplit"`) are read as frozen panes alone. `writeKept` writes the workbook's `<pane>` and selection in place of the file's where the split changed, and keeps the rest of the file's view (the other panes' selections where nothing changed, a PivotTable's selection, extensions).

Views change neither values nor `inputHash`. `getLayout()`, `getBook()` and snapshots keep them in every kind of session. Renaming the active sheet updates its name; deleting or hiding it selects the first visible sheet. The host tracks the live selection and scrolling itself, and applies `activeSheet` and `view` to the workbook before saving. They have no `setLayout` edit, so moving around a grid and showing or hiding gridlines create no undo steps. A tab color is the workbook's own: `setLayout` sets it with `{ sheet, tabColor: '#FF0000' }` and removes it with `{ sheet, tabColor: null }`, one undo step, without changing a value or `inputHash`.

#### How it is tested

By round trips through the writer and the reader, in every kind of session. Tab colors, gridlines and headings also in Excel 16.0.20430 for Windows (`bench/windows/run.ts fields sheets`, `test/fields-win.test.ts`): for files `writeXlsx` wrote, and `writeKept` wrote setting, changing and removing them, Excel read through its object model each sheet's tab color, gridlines and headings as the workbook holds them, and saved them so, as the reader reads them back. Split windows in Excel for Windows and in Excel 16.114 for Mac (`fields split`, `test/fields-win.test.ts`, `test/fields-excel.test.ts`): lines on both axes and on one, scrolled, zoomed, each pane active, under frozen panes, and set, changed and removed by `writeKept`; each Excel opened each window with the panes and the active pane the workbook holds and saved the same split, its lines on its screen's pixels. Frozen panes without a cursor in both (the parts sets `files`, `frozen-held`, `struct-held`): each Excel opened them with the cursor on A1 and saved it so.

### Filters

A sheet whose rows a filter hid gets `filtered: true`. Then `SUBTOTAL` with codes 1 to 11 also skips its hidden rows, as Excel does, and so do structure operations ([docs/structure.md](structure.md)). Excel treats a sheet as filtered based on what its filters contain, not on the `filterMode` marker alone:

| in the file | the sheet |
|---|---|
| an AutoFilter with criteria, with or without the marker | filtered |
| a table's filter with criteria | filtered |
| the marker without an AutoFilter, as an advanced filter leaves it | filtered |
| an AutoFilter without criteria, with or without the marker | not filtered |

Every hidden row of a filtered sheet counts as hidden by the filter, whatever hid it.

The sheet's AutoFilter, the filter buttons on a header row outside a table, comes into `autoFilter` with its range: `<autoFilter ref="A1:D20">` gives `autoFilter: { range: 'A1:D20' }`, with or without criteria. A host reads from it where a sheet's data under the buttons ends. Its criteria and sort, a table's criteria and an advanced filter are not in the workbook and are reported; so is an AutoFilter over a table or over no range of cells, which is left out. The hidden name Excel keeps for an AutoFilter, `_xlnm._FilterDatabase`, is not reported: the AutoFilter holds its range.

### The date system and calculation

A file with `workbookPr date1904` gives a workbook with `dateSystem: 1904`. Numbers stay as the file stores them: in such a workbook, serial 0 is 1 January 1904, and the ISO dates of Strict files count from there too.

`calcPr iterate` gives `iteration`, with `maxIterations` from `iterateCount` and `maxChange` from `iterateDelta` when the file has them. A count outside 1 to 32,767 or a change below 0 is read as Excel's default and reported. The count and the change of a workbook without iterative calculation are not kept: they have no effect.

```js
// <calcPr iterate="1" iterateCount="7" iterateDelta="0.5"/>
iteration: { maxIterations: 7, maxChange: 0.5 }
```

The calculation option is kept as `calculation`: `manual`, or `autoNoTable` (automatic except for data tables), and `writeXlsx` and `writeKept` write it; the engine calculates every edit whatever it says, and a host that offers manual calculation keeps the values it showed until the user calculates. Excel 16.0.20430 for Windows took each value from a file `writeXlsx` wrote and from files `writeKept` wrote setting and removing it, and saved it so (`bench/windows/run.ts fields book`). Precision as displayed is reported.

### Names

| in the file | in the workbook |
|---|---|
| `Data!$A$1`, `'Data'!$D:$D` | the reference |
| `#REF!`, `Data!#REF!` | the lost reference, as the engine stores it |
| `localSheetId="1"` | `scope`, the name of the second sheet |
| `hidden="1"` | `hidden: true`, a name Excel's Name Manager does not list ([Hidden names](names.md#hidden-names)) |
| `0.2`, `Data!$A$1*2`, `Data!$A$1:$A$3,Data!$C$1:$C$3` | the constant, the formula, the union ([docs/names.md](names.md)) |
| `_xlnm.Print_Area` and other built-in names | left out and reported |
| `_xlfn.XLOOKUP` (Excel's placeholder for a newer function) | left out |

### Tables

A sheet's table parts become its `tables`: the name formulas use, the range, the header and totals rows, each column with its calculated column formula and its totals function or label, and its style (`tableStyleInfo`): a table in Excel's default style, `TableStyleMedium2` with banded rows, has no `style`, and one in a custom style, which the model has no place for, is read in the default style and reported. Structured references come in the stored spelling, which matches the file's except for a column name that starts with `@` (`Sales[@at]` in a file becomes `Sales['@at]`). A header cell that does not contain its column's name as text is given it, and a table the model cannot represent, one that overlaps another table or a merged area, is left out; both are reported, as is any part of a calculated column formula the engine cannot handle, as for a cell's formula (at the column's first cell). A calculated column formula that differs from its cells' formulas is kept as Excel keeps it: a row inserted into the table gets it. [docs/tables.md](tables.md) has the rules.

### The locale

xlsx has no place for the workbook's locale. The writer stores it in a custom document property named `xlcalc.locale`, and the reader reads it back. Excel keeps the property when it saves the file. A file without it gives a workbook without a locale, that is, en-US.

### What is reported

| code | what the workbook cannot represent |
|---|---|
| `richText` | formatted runs of text |
| `userFunction` | a user-defined function (`_xludf.`): one issue per function on a sheet or a name, the function in `fn` (upper-case) and its calls in `cells`, so a host can count them without reading `message` |
| `table` | a table left out (overlapping another table or a merged area, or with a name already in use), a header cell given its column's name, a custom table style, a structured reference without a table name |
| `externalLink` | a DDE or OLE link, or a link whose part is missing, kept as a workbook without sheets; a formula over a link the file does not contain, or over such a link |
| `formula` | a formula the engine cannot parse, kept as written |
| `errorConstant` | an error stored as a value, read as a formula consisting of it |
| `errorValue` | an error the model has no code for, such as `#GETTING_DATA` |
| `dataTable` | a what-if data table, read as its stored values |
| `name` | built-in names, names of a sheet that was left out; a linked workbook's name for a sheet it does not have |
| `calculation` | manual mode, automatic except for data tables, precision as displayed, an iteration count or change out of range |
| `style` | pattern and gradient fills, diagonal borders, shrink to fit, subscript and superscript |
| `defaultFont` | a default font other than Calibri 11, kept in `font` and measured as Calibri 11 |
| `defaultWidth` | a default column width other than 8 digits |
| `numberFormat` | an East Asian built-in format (50 to 58), or an unknown one |
| `hidden` | a file whose worksheets are all hidden (next to chart sheets): its first worksheet is read as visible |
| `autoFilter` | an AutoFilter's criteria and sort, a table's criteria, an advanced filter; an AutoFilter over a table or over no range of cells, left out |
| `sheetKind` | chart, dialog and macro sheets, left out |
| `cellMetadata` | a rich value in a cell, such as a data type or a picture, or among a link's stored values |
| `comment` | a comment thread, read as a note; a note with differently formatted runs, read as plain text |
| `hyperlink` | a hyperlink whose address is not an external relationship of the sheet; one over more than 4,096 cells, read on its cells that contain something |
| `frozenPanes` | panes frozen on a sheet scrolled down or right, read from the first row and column |
| `conditionalFormat` | a rule of a kind the model does not support, a pivot table's rules, a rule without its formulas or thresholds; anything in a rule's differential format beyond the model's style (alignment, protection, a font's name and size, pattern and gradient fills); an average rule with both equal and standard deviations, read with equal |
| `dataValidation` | a rule without a range on the sheet, without the formula its type needs, or of an unknown type, left out; cells a rule shares with an earlier one, which keep the earlier one; an IME mode; a list whose items contain commas (`x12ac:list`), read split at them |
| `outline` | automatic outline styles |
| `protection` | workbook protection; a sheet protection's password (`sheetPassword`), the sheet read as protected without one; ranges a protected sheet allows to edit (`protectedRanges`) |
| `drawing`, `pivotTable`, `print`, `macros`, `sparkline`, `connection` | parts of the file the model does not support |
| `locale` | a locale the package does not have |
| `part` | a sheet name Excel would not take, read under another ([Sheet names](#sheet-names)); a missing sheet part, a cell outside the sheet, a Strict file's date that is no date, XML of the cells or the shared strings left open (an element without its closing tag), a width or height that is no size, a row number that is no whole number, a style index or a shared string the file does not have, a value that is no number or no boolean |

Gridlines and other view settings outside `SheetView` are not reported. A page setup is reported only when one was actually made, for example a landscape page or fit to width.

### The language of the messages

An issue's `message` is in English by default. `readXlsx(data, { locale })` and `writeXlsx(snapshot, { locale })` write it in the language of a locale: the tag of a loaded locale (`'ru-RU'` once `registerLocale(ruRU)` has run) or a locale object (`ruRU`), which needs no loading. A locale object without messages of its own uses those of the loaded locale with its tag, otherwise English. The code, sheet, cells and name are the same in every language, and so are the issues and their order.

| code | en-US | ru-RU |
|---|---|---|
| `richText` | formatted runs of text are read as plain text | Форматированный текст прочитан как простой |
| `style` | a style's shrink to fit is not in the model | Автоподбор ширины в стиле не поддерживается |
| `calculation` | the maximum number of iterations "0" is not between 1 and 32,767: the default, 100, is used | Предельное число итераций «0» вне пределов от 1 до 32 767: прочитано значение по умолчанию, 100 |
| `noCachedValues` | no result was passed: the file stores no formula values, and Excel calculates them when it opens the file | Результат не передан: файл не хранит значения формул, Excel вычислит их при открытии файла |

The host passes the language the user reading the messages has chosen. A file does not record it: Excel stores no language in xlsx, and the locale the writer stores in `xlcalc.locale` is the workbook's, the rules it calculates by, and a user working in another language may open the same workbook. So neither function takes the language from the workbook.

The messages are part of the locale's module, alongside its other messages ([docs/locale.md](locale.md#messages)), and the English ones are part of the subpath, so the subpath includes no locale other than en-US, and a locale without its own text for an issue falls back to the English one: a host that displays Russian loads `xlcalc/locales/ru-RU`, as it does for workbooks in ru-RU. One part of one message stays in English: why a table is left out, which comes from the table check's own text (`Таблица «Over» не прочитана, модель не может её хранить (B1:C2 overlaps the table T)`).

## Writing

`writeXlsx(snapshot, opts)` resolves to `{ data, issues }`: the file's bytes and what the file cannot represent. It writes the file anew from the workbook; to keep what the workbook cannot hold (charts, pictures, pivot tables, the page setup), write over the original file with `xlcalc/xlsx/keep` ([Writing over the original file](xlsx-keep.md)). The same workbook and result produce the same bytes. The snapshot is `{ book, result }`: the workbook, and the values of its formulas and of the cells their arrays spilled into, from `compute`, from a session (`session.getSnapshot()` returns both in one call, so no edit comes between them) or from the `cached` of a file that was read.

| option | |
|---|---|
| `values: 'require'` | the default with a result: every formula and every cell of an array must have its value in the result, otherwise the result does not belong to the workbook and the call throws `InputError` |
| `values: 'preserve'` | the values the result has; a formula without one is written without it and reported (`noCachedValues`) |
| `values: 'omit'` | the default without a result: no formula values, and Excel calculates them when it opens the file; reported (`noCachedValues`) when the workbook has formulas and `values` was not given |
| `mode: 'formulas'` | the default: formulas, with the values `values` requests |
| `mode: 'values'` | values only: a formula's cell gets its value from the result, which is then required |
| `locale` | the language of the issues' messages, as for `readXlsx`: English by default ([The language of the messages](#the-language-of-the-messages)) |

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

A session is not a snapshot: an asynchronous session is a proxy for a thread, and its workbook and result are obtained there with `await session.getSnapshot()`. Given a session, a result alone (`writeXlsx(compute(book))`, recognized by its hashes) or a result of another shape, `writeXlsx` throws an `InputError` saying what to pass. The older form, with the workbook first and its result in the options (`writeXlsx(book, { result, mode })`), throws an `InputError` too: a workbook is not a snapshot.

A workbook written and read back keeps its values and their `hash`, but usually gets another `inputHash`, since the input comes back in the file's form: formulas in canonical text (`sum(a1)` as `SUM(A1)`, sheet names quoted), typed dates as serial numbers, number formats by their built-in ids and the style table renumbered. To tell whether a workbook's data changed since it was saved, compare `hash`, or keep the `inputHash` of the workbook as it was read back.

### Formulas

Every formula is written as a dynamic array formula, as Excel 365 saves one: with cell metadata and the range its array spilled into. The engine calculates every formula by the rules of dynamic arrays, and so does Excel with this mark. Excel 2019 and older display such formulas as array formulas in braces.

The spilled cells are written with their values. `#SPILL!` and `#CALC!` are written as Excel 365 writes them, so Excel reads them back as those errors.

Functions newer than Excel 2007 get their prefix (`_xlfn.XLOOKUP`, `_xlfn._xlws.SORT`), LET names become `_xlpm.x`, `E2#` becomes `_xlfn.ANCHORARRAY(E2)`, and `@x` becomes `_xlfn.SINGLE(x)`. Sheet names are always quoted. A reference to a deleted sheet, which the workbook holds as Excel shows it (`#REF!A5`, `#REF!E2#`), is written `#REF!` (`_xlfn.ANCHORARRAY(#REF!)`), as Excel writes it: Excel does not open a file with the other form. Read back, the formula has `#REF!` (`#REF!#`).

A cell with `array` is written as a legacy array formula over its range, without the cell metadata, as Excel writes a formula entered with Ctrl+Shift+Enter. A formula in a table's totals row is written as a plain formula, the only kind Excel accepts there. A name's formula is written as the workbook has it: a constant, a formula or a union. The list of prefixes covers every function Microsoft lists as newer than 2007, not only the whitelist.

A file with formulas asks Excel to recalculate the workbook when it opens it, so Excel displays its own values over the stored ones, whichever values the file stores, except for the cells a circular reference stops, which keep the stored ones: a formula with `held` stores it where the result has no value, with `values: 'omit'` too ([cycles.md](cycles.md)). The stored values are what a reader that does not calculate displays, such as a preview or a spreadsheet library, and what `readXlsx` returns as `cached`.

### Tables

Each table is written as a table part, as Excel writes one: its name, range, header and totals rows, an AutoFilter over the header and data, its columns with their calculated column formulas and totals, and its style, Excel's default for a table without `style`. A formula in the totals row that none of its column's totals functions produces is recorded as the column's custom totals formula, without which Excel does not open the file, and a table's name alone in a defined name is written as `Sales[]`.

### Styles, columns and rows

Styles, number formats, widths, heights, hidden rows and columns, outline levels and merged areas are written as the workbook has them. The default font is the workbook's `font`, Calibri 11 without one, and a column without a width gets Excel's standard width of 8 digits. Every cell is written with the formatting it displays, its row's or column's where it has none of its own, as Excel writes it ([The cells of formatted rows and columns](#the-cells-of-formatted-rows-and-columns)).

The cells in a sheet's `looks` are written as Excel stores them, each as a cell with the cell format of the full formatting it displays (`<c r="B2" s="3"/>`), row by row together with the entries, without an object per cell. Cell formats are numbered in the order cells appear, with the `looks` cells among the entries as in a file that was read, so a workbook read from a file and written again produces the same file the reader would have produced when such cells were entries. A sheet without `looks` is written as given, its entries in their order.

A built-in format given by its number is written as that number, and Excel displays it according to its own locale. A code that a locale displays differently from the built-in format it equals, such as `m/d/yyyy` (format 14), is written as a custom code of the file, so every locale displays what `numberFormat` specifies. A code that every locale displays the same way, such as `0.00`, is written as its built-in number, as Excel writes it, and comes back as that number, `2`.

### Hidden sheets

A hidden sheet is written with its state, `hidden` or `veryHidden`. The workbook's active sheet, the one Excel shows on opening, is `activeSheet` ([The saved active sheet and view](#the-saved-active-sheet-and-view)), or without it the first visible sheet: Excel displays a hidden sheet that a file names as active, so a file whose first sheet is hidden names the next visible one (`<workbookView activeTab="1"/>`, with that sheet's tab selected).

### Filters

A sheet's `autoFilter` is written as `<autoFilter ref>` over its range, with the hidden name `_xlnm._FilterDatabase` over the same cells, as Excel writes both when it makes an AutoFilter. Excel opens the file with the filter buttons on the range's first row.

A filtered sheet is written so that Excel treats it as filtered, with Excel's `filterMode` marker: its AutoFilter gets filter criteria that every row satisfies (a cell equal to nothing, or not equal to it). A sheet without an AutoFilter puts them on its first table with a header row, as Excel keeps a table's filter, and without such a table gets an AutoFilter over its cells and hidden rows to hold them. Excel leaves the rows as the file has them, and its Reapply shows every row. A sheet without an AutoFilter whose tables have no header row and overlap its cells cannot have one there: its filter mode is left out and reported.

### The date system and calculation

`dateSystem: 1904` is written as `workbookPr date1904`, and typed dates (`{ value: { date: '2025-01-31' } }`) as serial numbers in the 1904 system, 44226. `iteration` is written as `calcPr iterate` with the count and the change the workbook specifies. Every file has `calcPr`: without it, Excel keeps the calculation settings from a workbook opened earlier, including iterative calculation.

### What the file cannot represent

| code | what happens |
|---|---|
| `meta` | a cell's `meta` is left out |
| `errorValue` | `#CYCLE!`, which a self-referencing name returns, is written as 0, as Excel stores a circular reference, including among a link's stored values; `#ERROR!`, the engine's error for a formula that does not compute, is written as `#NULL!`, as Excel has no `#ERROR!` |
| `textLength` | text longer than 32,767 characters is truncated, including a note's; also a hyperlink's tooltip over 255, a rule's titles over 32, its input message over 255 and its alert text over 225 |
| `dataValidation` | a rule whose formula does not parse, or names a sheet, a table or a column the workbook does not have, is left out |
| `formula` | a formula that does not parse, or names a sheet, a table, a column or a link to another workbook the workbook does not have, is written as its value, or as an empty cell when it has none (no result given) |
| `noCachedValues` | the file stores no formula values, because no result was given, or a formula has no value in the result (`values: 'preserve'`) |
| `name` | a name Excel does not accept, given twice, scoped to a missing sheet, or whose reference does not parse or names a table, a column or a link the workbook does not have, is left out |
| `autoFilter` | the filter mode of a sheet without an AutoFilter whose tables have no header row and overlap its cells is left out |
| `conditionalFormat` | a rule with a formula the file cannot represent (it does not parse, or names a sheet, a table or a column the workbook does not have) is left out |

## The cells of formatted rows and columns

A workbook gives a cell without its own style or number format its row's, otherwise its column's, the style and the number format each separately: `{ value: 46096 }` in a column formatted `dddd, mmmm d, yyyy` displays `Sunday, March 15, 2026`. Excel interprets a file differently. The writer and the reader convert between the two, so a workbook written and opened in Excel displays what the engine displays, and a file saved by Excel reads into a workbook that displays what Excel displayed.

### Cells the file stores

Excel displays a cell the file stores (a `<c>` element) by the cell's own cell format alone, whatever its row and column have. So the writer writes each cell with the full formatting it displays. A value in that column of dates is written with the column's cell format, the one Excel gives a value typed there, and a cell with its own bold style in a filled column of dates is written as bold, as a date, and without the fill, as the workbook displays it.

```js
// columns: { B: { numberFormat: 'dddd, mmmm d, yyyy', style: 1 } }, cells: { B2: { value: 46096 } }
// <col min="2" max="2" … style="1"/> and <c r="B2" s="1"><v>46096</v></c>
```

The reader keeps only the part of a stored cell's formatting that its row or column does not provide: `B2` of that file comes back as `{ value: 46096 }`, without its own number format. Where a cell stores less than its row or column provides, the reader writes out the default explicitly: a cell stored as General in a column of dates gets `numberFormat: 'General'`, and one without a fill in a filled column gets the workbook's empty style, a style with nothing set.

### Where a row crosses a column

An empty cell the file does not store displays its row's full cell format when the row has one (`customFormat`), otherwise its column's. So where a row with only a style crosses a column with only a number format, Excel displays the empty cell, and a value typed there, without the column's number format: `46096` typed where a bold row crosses a column formatted as percentages displays `46096`, in bold, where the workbook displays `4609600.00%`.

Excel itself never leaves such an intersection empty: when it formats a column across a formatted row, or a row across a column, it writes the cells where they cross, empty, with both formats. The writer does the same: where a formatted row crosses a formatted column and Excel would display the empty cell differently from the workbook, it writes that cell with the formatting the workbook displays. The reader reads a file without these cells as Excel displays it: a row with its own cell format that omits a crossing column's style or number format, where a cell of that column is missing, gets the empty style or General.

```js
// <row r="6" s="1" customFormat="1"> (bold), <col min="7" max="7" style="3"/> (0.00%), no <c r="G6">
rows: { 6: { style: 0, numberFormat: 'General' } }
```

### Lines Excel formats

When Excel formats a row or column across a formatted one, it merges their formats field by field: a row made bold across a filled column of dates makes the cells where they cross bold, filled and dates. A workbook style is a single unit, and `setLayout` sets only a row's or column's own style: where a row with a style crosses a column with one, the workbook displays the row's. `patchCells` formats rows and columns as Excel does and gives those cells a style of their own ([Editing a grid](edits.md#entire-rows-and-columns)).

### Checked against Excel

Tested in Excel for Mac and for Windows: Excel opened the written workbooks, formatted rows, columns and cells through its object model, and displayed what the engine displays.

## Protection

A sheet's protection and its cells' come from the file and go back into it, so a host can show which cells its users may edit and save the file protected as it was ([Protection](structure.md#protection)).

### Reading

A cell format's `<protection>` is its style's `protection`: `locked="0"` is `{ locked: false }`, `hidden="1"` is `{ hidden: true }`, and a cell format without the element is locked and not hidden. Excel takes the cell format's own element whatever its `applyProtection` says, and does not take the protection of its cell style (`cellStyleXfs`), so neither does the reader.

`<sheetProtection>` with `sheet="1"` is the sheet's `protection`. Its attributes say the opposite of the model's fields: `formatCells="0"` lets users format cells, `formatCells: true`. An attribute the element leaves out takes the schema's default, which leaves objects and scenarios unprotected: a file with `<sheetProtection sheet="1"/>` alone reads as `{ objects: true, scenarios: true }`, as Excel reads it, where the protection Excel's dialog gives is written `<sheetProtection sheet="1" objects="1" scenarios="1"/>`. An element without `sheet="1"` protects nothing, and the sheet reads unprotected. A password (`password`, or `algorithmName` with `hashValue`, `saltValue` and `spinCount`) is not in the model: the sheet reads as protected, and the password is reported (`protection`, `sheetPassword`). So are the ranges a protected sheet allows to edit (`protectedRanges`).

### Writing

`writeXlsx` writes a style's protection into its cell format with `applyProtection="1"`, and a sheet's protection after its cells, with the attributes Excel writes: `{}` becomes `<sheetProtection sheet="1" objects="1" scenarios="1"/>`, and `{ formatCells: true }` adds `formatCells="0"`. It writes no password, so a user unprotects the sheet without one. `writeKept` keeps the file's element with its password while the protection did not change, sets the attributes of what users may do in that element when it did, and removes it when the sheet is no longer protected ([docs/xlsx-keep.md](xlsx-keep.md)).

### Checked against Excel

Excel 16.114 for Mac opened, without its repair window, every file written for the protection sets, read the protection the engine reads and saved it as the writer writes it: files `writeXlsx` wrote, files changed by hand (`applyProtection="0"`, a cell style's protection, an element without `sheet="1"`, with `sheet="0"`, with `sheet="1"` alone, with a password), and files `writeKept` wrote over a protected one with a password ([docs/development.md](development.md#protection)).

## Conditional formats

A sheet's conditional formats (`conditionalFormats`) round-trip. The reader reads the rules in `<conditionalFormatting>` with their differential formats (`<dxfs>` in styles.xml) and Excel 2010's rules in the sheet's extension list (`x14:conditionalFormatting`), and returns them in priority order; the writer writes them the same way, with each rule's priority given by its position in the list. Excel 2010's part of a data bar (its fill, border, direction, negative bars and axis) is joined to its 2007 rule by id; the icon sets `3Stars`, `3Triangles` and `5Boxes`, custom icons, and formulas referring to other sheets exist only as Excel 2010 rules, with their style inline. A text, blank or time period rule always has its formula in a file; the reader omits the one Excel writes for it, which the model generates itself. Ranges come as the model writes them: Excel's `A1:A1048576` becomes `A:A`, and a range another writer spells in lowercase or with `$` is read as Excel reads it, `$a$1:b2` as `A1:B2`. A dxf's number format is read by its code, since Excel saves `0.00%` under another id and displays a dxf by its code in every locale (`m/d/yyyy` under id 14 displays `3.15.2026` in ru-RU). A dxf's fill is read as Excel reads it, which differs from a cell's: a solid fill displays only its `bgColor` (white for one that names no color), a fill without a pattern type is solid, and an empty one changes nothing. `patternType="none"` removes the cell's own fill, and a solid fill without `bgColor` keeps it (white on a cell without one); the model can express neither, so such a rule is read without a fill, and that is reported. Excel opens no file whose dxf contains a gradient alongside a pattern; the reader reads the pattern and reports the gradient. [docs/conditional-formats.md](conditional-formats.md) has the rules, how a cell is displayed under them, and how this was checked.

## Notes, hyperlinks and frozen panes

A cell's note and hyperlink and a sheet's frozen panes round-trip, so a grid can display them and a file saved from it keeps them. They are a cell's `note` and `link` and a sheet's `frozen`.

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

A note has its author and its text, with lines separated by `\n`. Excel keeps notes in a comments part with a VML drawing that displays them, and the reader takes the comments part. The text is read as plain text: a note created in Excel's window has its author's name in bold at the start, and such formatting is reported. A note shown all the time (Review > Notes > Show/Hide Note, `<x:Visible/>` in the drawing) has `shown: true`, and the writer writes it so; Excel 16.0.20430 for Windows read each note `writeXlsx` and `writeKept` wrote as shown or not as the workbook holds it, and saved it so (`bench/windows/run.ts fields notes`). The size and position of a note's box are not kept. A note's cell is read as Excel reads it: `A01`, `$A$1`, `a1` and `A1:A1` all mean A1; a note on a range of cells is left out and reported, since Excel does not open such a file.

The writer writes both parts, as Excel writes them: each note gets a hidden box next to its cell. A note without an author is written with an empty one, as Excel keeps it.

### Threads of comments

Excel 365's comments are threads: a first comment and its replies, each with its author. The model has only notes, so the reader reads a thread as a note. The note's author is the first comment's author, and its text is every comment's text, with each reply after an empty line, preceded by its author:

```
Is this the final figure?

Boris: Yes, from the audit.
```

A thread's dates, its mentions and whether it is resolved are not kept, and each thread is reported. Next to each thread, Excel writes a separate note for older versions of Excel (`[Threaded comment] Your version of Excel allows you to read...`). The reader leaves those notes out. A version of Excel without threads moves such a note with its cell when rows are inserted, and leaves the thread where it was, so the reader puts the thread on the cell of its note.

The writer writes a thread read this way as a note. Excel displays it as a note, with the thread's text in it.

### Hyperlinks

A hyperlink has a `url`, an address outside the workbook, or a `location`, a place in it (`'Totals'!A1`, a defined name), or both: then the location is inside what the address points to. A `tooltip` is the text Excel shows when the pointer hovers over the cell. The cell displays its value, not the link, so a link is set on a cell that already contains its text. A hyperlink over a range is read on each of its cells.

Excel treats the part of an address after `#` as its location: a file containing `https://example.com/page#top` comes back from Excel as `https://example.com/page` with the location `top`. Excel also saves a space in an address as `%20`, and the reader returns the address as the file stores it. A tooltip longer than 255 characters is truncated, since Excel does not open a file with a longer one.

### Frozen panes

`frozen: { rows: 2, columns: 1 }` keeps the first two rows and the first column in view, as Excel's Freeze Panes does with C3 selected. Split panes that are not frozen are not in the model and are reported.

### Structure operations

A note and a hyperlink go with their cell: a move takes them along, a copy copies them, and a deleted cell, or a cell that a move or a copy lands on, loses them. Inserted rows and columns get none. A hyperlink's location is kept as written when its sheet is renamed or deleted, as Excel keeps it. Rows inserted above the split of frozen panes push it down, and deleted ones pull it up, but at least one row stays frozen; the same applies to columns. [docs/structure.md](structure.md) has the rules.

### Checked against Excel

Tested in Excel for Mac and for Windows: both opened every written workbook with notes, hyperlinks and frozen panes (except one with a note over a range, which neither opens) and saved what the engine holds.

## Data validation and outlines

A sheet's data validation rules and its outline round-trip, so a grid can offer a cell's list and check what is typed, and a file saved from it keeps them. They are a sheet's `validations`, the rows' and columns' `outlineLevel` and `collapsed`, and a sheet's `outline`.

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

A file stores a rule in `<dataValidations>` with xlsx's defaults of false: a rule without `allowBlank`, `showInputMessage` or `showErrorMessage` is read with `allowBlank: false`, `showPrompt: false` or `showError: false`, since the model's defaults are those of a rule Excel creates. `showDropDown="1"` hides a list's dropdown and is read as `dropdown: false`. Excel writes a rule whose formulas refer to another sheet in the sheet's extension, `<x14:dataValidations>`, which the reader reads alongside the others. Formulas are converted to the engine's spelling, like a cell's. Excel saves an entire column as `F1:F1048576`, read as `F:F`; a range in lowercase or with `$` is read as Excel reads it, `$a$1:b2` as `A1:B2`.

Two rules over the same cell are not in the model. Excel writes them when a move puts a rule over a cell that keeps another one in its range, and applies the one written first, which is the one the reader keeps; the later rule loses the cell, and that is reported. A list typed where the comma is not the list separator (in ru-RU) can contain an item with a comma: Excel keeps that list in an extension (`x12ac:list`), which the reader does not read, and writes a fallback with the list split at its commas, which the reader reads and reports.

The writer writes every rule in `<dataValidations>`, including references to other sheets, which Excel opens. A rule's titles longer than 32 characters, its input message longer than 255 and its alert text longer than 225 are truncated, since Excel does not open longer ones.

### Outlines

A row's or column's `outlineLevel` and `collapsed` are attributes of `<row>` and `<col>`, and the summary settings are `<sheetPr><outlinePr summaryBelow summaryRight>`. The writer writes the deepest levels into `<sheetFormatPr outlineLevelRow outlineLevelCol>`, where Excel reserves room for the outline's buttons, and the reader takes them from the rows and columns. The writer writes `collapsed` as Excel keeps and saves it, on the summary row or column of a group whose rows or columns are all hidden, whatever markers the workbook holds; the reader reads a file's markers as they are ([docs/outline.md](outline.md#in-xlsx-files)). Automatic outline styles (`applyStyles`) are reported.

### Checked against Excel

Tested as described in [Data validation](validation.md): both versions of Excel opened every written workbook and read the rules, each cell's check and the outlines the engine holds.

## Links to other workbooks

A formula can read another workbook: `[1]Sheet1!A1+1`, `SUM('[1]My Sheet'!A1:B9)`, `[1]!Rate`. Excel stores in the file the values from the other workbook it last read, and calculates such formulas from them when it opens the file without updating its links. The reader returns these links as the workbook's `externalBooks`, so the engine calculates the formulas as Excel does, and the writer writes them back, so a file saved from a grid keeps its links.

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

Excel displays a cell stored in the link as its stored value, `null` as an empty cell, and `''` as empty text. A cell not stored in the link is empty, or `#REF!` on a sheet with `refreshError`, the marker Excel puts on every sheet of a link when a formula referring to it is entered. Every cell of a sheet without `cells` is `#REF!`. [docs/links.md](links.md) describes how the engine calculates and displays such formulas.

### Reading

The first link the workbook lists (`<externalReferences>`) is `[1]`, the second `[2]`, whatever its part is named.

| in the file | in the workbook |
|---|---|
| the target of the link's own relationship, of any kind (including `xlPathMissing`): `Book2.xlsx`, `../data/Book3.xlsx`, `/xlcalc/work/Book2.xlsx`, `file:///C:/data/Book4.xlsx`, `https://example.com/Book6.xlsx` | `target`, as the file stores it (`Sub%20Folder/Book7.xlsx` keeps its `%20`) |
| Excel's other addresses of the workbook (`xxl21:alternateUrls`) | left out: the target stays the relationship's |
| `<sheetName val="Sheet1"/>` | a sheet in `sheets`, in the file's order |
| `<definedName name="Local" refersTo="=Sheet1!$C$2" sheetId="0"/>` | a name in `names`: `ref` without `=`, and `sheet`, the index of the sheet the name belongs to |
| `<sheetData sheetId="0" refreshError="1">` | that sheet's `cells` by address, and `refreshError: true` |
| a sheet without `<sheetData>` | a sheet without `cells` |
| `<cell r="A1"><v>5</v></cell>` (also `t="n"`) | `5` |
| `t="str"` (and `t="s"` or `t="inlineStr"`, which another writer may produce) | text; `<v/>` is empty text `''`, as Excel writes the value of `=""` |
| `t="b"`, `t="e"` | a boolean, an error (`{ error: 'NA' }`) |
| `<cell r="A5"/>` | `null`, a cell stored as empty |

A link to another program's data (DDE) or to an object (OLE), and a link whose part is missing, are not in the model. Each keeps its number as a workbook with an empty target and no sheets, so the links after it keep theirs, and each is reported (`externalLink`). A stored value with a rich value (`vm`, a data type, or a #SPILL! that Excel 365 stores as `#VALUE!`) is read as the value the cell stores, which is what Excel displays, and reported.

Formulas and names referring to links come as the file writes them: `[1]Sheet1!A1+1`, `'[1]My Sheet'!A1`, `[1]!Rate` (a workbook-level name), `[1]Sheet1!Local` (a sheet-level name), and `[1]Sheet1!#REF!` (a reference into it that Excel lost); this workbook's name `Ext` referring to `[1]Sheet1!$A$2` keeps its reference. A shared formula's copies shift their relative references into the link's sheet, as a paste does: `[1]Sheet1!A1` copied one row down becomes `[1]Sheet1!A2`. Copies stored in the file are not edits, so they put no marker on the link. A formula referring to a link the file does not contain (`[3]Sheet1!A1` in a file with two links), or to a DDE or OLE link, is read as written and reported.

### Writing

Each link in `externalBooks` is written as Excel writes it: a part `xl/externalLinks/externalLinkN.xml` with the sheet names, the names (`refersTo` with its `=`) and the stored values; its relationship to the other workbook, external, with the target as the workbook stores it; and the workbook's `<externalReferences>` after its sheets. A number is written without a type, text as `t="str"` (`''` as `<v></v>`), a boolean as `t="b"`, an error as `t="e"`, `null` as `<cell r="A5"/>`, the marker as `refreshError="1"`, and a sheet without `cells` gets no data. A workbook without links is written as before.

A formula or a name referring to a link is written in the file's spelling, with the sheet in quotes only where its name needs them: `'[1]Sheet1'!A1` in the workbook is `[1]Sheet1!A1` in the file, `'[1]My Sheet'!A1` stays quoted, and `[1]!Rate`, `[1]Sheet1!Local` and `[1]Sheet1!#REF!` stay as they are. A reference by the other workbook's file name, as a user types it (`[Book2.xlsx]Sheet1!A1`), is written with its link's number. A formula referring to a link or a link's sheet the workbook does not have is written as its value, and a name referring to one is left out; both are reported.

### What is lost

- Excel's other addresses for a workbook (`xxl21:alternateUrls`): the writer writes only the target, which Excel opens.
- A stored value's rich value (`vm`).
- #SPILL! and #CALC! among a link's values, written as `#VALUE!`. Excel keeps no other form there: it repairs a file that contains their text, and displays and saves `#VALUE!` for them, stored with a rich value, as a sheet's cells store them.
- The marker of a sheet without `cells`, which changes nothing: every cell of such a sheet is `#REF!` with or without the marker.
- A DDE or OLE link, except for its number.

When it saves, Excel itself leaves out the values of the other workbook's cells that no formula reads, and any link no formula refers to, renumbering the following links in the formulas.

### Checked against Excel

Tested in Excel for Mac and for Windows, which opened the files without updating their links.

## Speed

Measured on an Apple Silicon Mac with `test/fixtures/xlsx-bench-child.ts`, the median of three runs. The ranges cover several runs on a busy machine.

| workbook | file | read | write |
|---|---|---|---|
| 20,000 rows × 40 values | 3.0 MB | 0.7 to 1.05 s | 1.2 to 1.8 s |
| the synthetic workbook, 3,400 rows, 3,900 formulas with values | 0.41 MB | 0.1 s | 0.1 to 0.15 s |
| a wide-shape workbook, 5.6 million formatting-only cells (`bench/memory/wide-shape.ts`) | 15 MB | 1.5 s | |

A workbook of 800,000 values, once read, takes about 100 MB of heap; the wide-shape workbook about 1 MB, and a session over it opens in 7 ms (6.2 s and 815 MB, then 10.5 s and 545 MB, when such cells were entries). `FORMULA_BENCH=1 npx vitest run test/xlsx-budget.test.ts` checks that the large workbook is read in at most 2 s.

## How the code is organized

`src/xlsx/` contains the code, and `src/xlsx.ts` is the subpath's entry point.

- `zip/` reads and writes the archive. Inflating and deflating go through `DecompressionStream` and `CompressionStream`.
- `xml/` contains the text layer: entities, Excel's escapes, and elements found by regular expressions.
- `read/` reads the package, the workbook, the sheets and their cells, the cell metadata, the notes, threads, hyperlinks and frozen panes (`parts.ts`), conditional formats (`conditional.ts`), data validation (`validation.ts`) and the links to other workbooks (`links.ts`); `refs.ts` parses the references of a rule's cells and a note's cell as Excel parses them.
- `styles/` converts styles.xml into the style table and the differential formats (`dxf.ts`), whose fills one function reads (`readFill`), and contains the codes of the East Asian built-in formats, the codes the writer stores by number, and the colors.
- `formulas/` converts formulas between the file's spelling and the engine's. `prefixes.ts` is generated by `bench/coverage/xlsx/generate.ts`, so the reader needs neither the harness nor the engine's function definitions. Legacy formulas go through `fromLegacy` (`src/legacy/`), whose table of argument classes needs no function definitions either.
- `write/` writes the workbook, the sheets, styles.xml with its differential formats, the notes, hyperlinks and frozen panes (`parts.ts`), conditional formats (`conditional.ts`), data validation (`validation.ts`) and the links to other workbooks (`links.ts`).

Reading uses the parser and the copy and paste of the structure operations. It does not load the engine's functions or evaluator, and a minified browser bundle of the subpath is about 175 KB. The issues' messages are stored as a key with parameters (`src/xlsx/issues.ts`) and rendered in a locale when listed: a locale's text is in its file alongside its other messages (`src/locales/ru-RU.ts`), and the English text is in the subpath (`src/xlsx/messages.ts`), so that `compute`, the root entry and `xlcalc/text` do not include it.

## Checked against Excel

Excel 16.113 for Mac and Excel for Windows opened every file the writer wrote without the repair dialog, calculated it and saved it, and the reader read the files Excel saved back as Excel read them. Excel 16.114 for Mac saved a sheet's AutoFilter written alone and on a filtered sheet over the range written, and the AutoFilters it made itself read with their ranges; Excel for Windows was not run on AutoFilters.

## Known differences

- Excel saves the literal characters of a number format escaped, so `dd.mm.yyyy` comes back as `dd\.mm\.yyyy`. It is the same format.
- Excel saves a range covering entire columns or rows as such, so `A1:A1048576` comes back as `A:A`.
- Excel saves a global name qualified by a sheet that has no name of its own, `'S'!x`, as the workbook-level name `[0]!x`. It comes back as `x`, the same name.
- Excel saves a LET that ends with a value with its last name as the calculation. With 127 names that makes 255 arguments, one more than the engine accepts, so such a formula does not parse after Excel saved it.
- Excel for Mac saves the built-in currency and accounting formats (5 to 8, 41 to 44) as codes in the locale it runs in, 23 to 26 as General, and the East Asian 27 to 36 as 14 and 21, so a workbook read from a file saved by Excel has these codes and numbers.
- Some theme colors with a tint differ from Excel's by up to 3 of 255 in a channel.
- The East Asian date formats 50 to 58 are read as the code `m/d/yy`, which Excel for Mac en-US displays for them; Excel for Windows displays them as its short date, format 14, in en-US and ru-RU.
- In a few places Excel calculates a legacy formula differently from the dynamic formula it displays for it. The engine returns the dynamic formula's value. [docs/legacy.md](legacy.md) lists them.
- An AutoFilter's criteria and sort are not kept. A filtered sheet comes back with the writer's criteria, which every row satisfies, and Excel's Reapply then shows every row.
- A filtered sheet without an AutoFilter or a table with a header row is written with an AutoFilter over its cells, and reads back with it.
- A style, a height or `hidden` given to a span of rows is written row by row, so a span to the last row writes a million rows.
- The model has no default column width of its own: the reader reports a file with another one, and the writer always writes 8 digits. The default font is kept (`font`), but the engine measures as for Calibri 11, so a column's `####`, `entryWidth` and AutoFit can differ from Excel's in a workbook of another default font.
- Excel treats the part of a hyperlink's address after `#` as its location, so `https://example.com/page#top` comes back as the address `https://example.com/page` and the location `top`. It is the same link.
- Excel saves a space in a hyperlink's address as `%20` and displays it as a space. The reader returns the address as the file stores it.
- Excel for Windows saves a relative hyperlink address (`other.xlsx`) relative to the folder it saves the file to.
- A comment thread is read as a note, without its dates, mentions and state, and written as a note.
- Excel keeps an Excel 2010 data bar's lengths only when they are 0 and 100: a bar of 5 and 95 comes back and is read as 10 and 90.
- A dxf number format that is a built-in one comes back from Excel as its code (`0.00%` for 10).
- A gradient fill is read as no fill; Excel displays the gradient. A conditional format fill with `patternType="none"`, which removes a cell's own fill in Excel, and a solid fill without a color, which Excel displays as white on a cell without a fill, are read as no fill. Each is reported.
- A data validation list whose item contains a comma (typed in ru-RU) is read as split at the comma. Excel saves a list's reference to a deleted sheet as `#REF!` alone, which the engine keeps as `#REF!$A$1:$A$4`.
