# A grid over a session

What a grid reads from a session to draw a sheet: a window of cells with their looks, the sizes of rows and columns, merged areas and tables; and how it changes the layout and the looks of many cells at once.

## A window for a grid

A grid draws a window of a sheet with `getViewport(sheet, range, opts?)`: in one call it gives the cells with their look resolved, the sizes of the window's lines, the merged areas and tables that meet it, the frozen panes and where the sheet ends.

```js
import { createSession, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const session = createSession({
  locale: 'ru-RU',
  styles: [{ font: { bold: true } }],
  sheets: [{
    name: 'Plan',
    cells: {
      A1: { value: 'Qty' }, B1: { value: 'Total' }, D1: { formula: '1/A3' },
      A2: { value: 2 }, B2: { formula: '[@Qty]*1.5' }, D2: { formula: 'SUM(B1:D1)' },
      A3: { value: 0 }, B3: { formula: '[@Qty]*1.5' },
    },
    rows: { 1: { style: 0 } },
    columns: { B: { numberFormat: '0.00', width: 12 } },
    merges: ['E1:F2'],
    tables: [{ name: 'T', range: 'A1:B3', columns: [{ name: 'Qty' }, { name: 'Total' }] }],
    frozen: { rows: 1 },
  }],
})
const v = session.getViewport('Plan', 'A1:E2', { formulaLocal: true })
console.log(v.cells.C1, v.cells.B2)
// → { value: null, text: '', effectiveStyle: 0 } {
//     value: 3,
//     text: '3,00',
//     formula: 'T[[#This Row],[Qty]]*1.5',
//     formulaLocal: '[@Qty]*1,5',
//     numberFormat: '0.00',
//     numberFormatSource: 'column'
//   }
console.log(v.styles, v.columns.map((c) => c.width), v.rows, v.merges, v.usedRange)
// → { '0': { font: { bold: true } } } [ 8.83203125, 12, 8.83203125, 8.83203125, 8.83203125 ] [ { row: 1, height: 15 }, { row: 2, height: 15 } ] [ 'E1:F2' ] { rows: 3, columns: 6 }
console.log(v.tables, v.frozen) // → [ { name: 'T', range: 'A1:B3', headers: true, totals: false } ] { rows: 1, columns: 0 }
console.log(session.getCellIssue('Plan', 'D2'))
// → {
//     error: 'DIV0',
//     reason: "division by zero, from 'Plan'!D1",
//     origin: { sheet: 'Plan', cell: 'D1' }
//   }
console.log(session.explain('Plan', 'D2').text)
// → Plan!D2 = #ДЕЛ/0!  ← СУММ(B1:D1)  (ошибка из Plan!D1)
//     Plan!B1:D1 (2 ячейки)
//       Plan!D1 = #ДЕЛ/0!  ← 1/A3  (ошибка возникает здесь)
//         Plan!A3 = 0
```

A cell of the window carries what `getCells` gives (the value, the shown `text` and `color`, the formula, the spill, `meta`, a note, a link, an old-kind array's area) and its look resolved: its own `style` and the `effectiveStyle` it shows, its own, else its row's, else its column's, as numbers of the book's styles, as `getCells` and the changes of edits give a style, with the objects of the styles the window shows in `styles`, copies by number; the `numberFormat` it shows by as a code, with `numberFormatSource` saying whether it is the cell's, the row's or the column's (a built-in format by the code of the book's locale: 14 is `dd.mm.yyyy` in ru-RU); and `appearance`, what its conditional formats give it. The style and the format are taken each on its own, so a cell of a bold row in a column formatted `0.00` is bold and shows `0.00`. An empty cell is listed when it shows something: a style of its row or column (`C1` above), a number format of its row or column (an empty cell of a column formatted `0.00` has `numberFormat: '0.00'` and `numberFormatSource: 'column'`, as Excel's list of formats shows it), an appearance, a note or a link; so are the cells of a spill, a large array's cells past its data too. A cell where a styled row crosses a styled column shows the row's style, as Excel shows a file's empty cell there. `columns` and `rows` give each line's number (`column`, `row`, from 1), width and height, Excel's default where it has none (8.83203125 and 15 points), with `hidden`; `merges` the merged areas that meet the window, whole; `tables` the tables that meet it, whole, each with whether its first row is a header row and its last a totals row; `frozen` the sheet's frozen rows and columns, 0 for none; `usedRange` is `getUsedRange(sheet)`. `formulaLocal: true` adds each formula as Excel shows it in the book's locale, inside its table as Excel shows it there (`[@Qty]`), and `diagnostics: true` adds to each cell with an error its `issue`, as `getCellIssue` gives it. A window holds at most `maxCells` cells, 20,000 by default: of a larger one it gives its first rows that fit, with `truncated: true` and the `range` it gave. It costs what the window holds, not the book: on the synthetic book of 34,000 rows a window of 20,000 cells takes 15 ms, as on the book of 3,400 rows and a little less than `getCells` and `getLayout()` together, and 26 ms from a thread.

`getCellIssue(sheet, cell)` says why a cell holds an error without building the book's result: the error, the reason, the check code and position of a formula that does not read, as the result's `diagnostics` give them, and `origin`, the cell the error began in, found by following the cells the reasons name (in a cycle, the first cell of the cycle the way comes back to); a formula a circular reference stops has one too, `CYCLE`, whatever value it holds. `explain` gives a structured `root`: local nodes carry `location: { sheet, range }`, invariant `formula`, `value` and `children`, for a panel without parsing its rendered `text`. Async `explain` also takes `signal` and `timeoutMs`, and a cancelled queued or running call rejects with `CalcError` code `aborted` ([Async calls](async.md#calls)). `explain` marks each cell with an error by that cell and leads to it: through a range by the cell the error came from, and past `depth` for `errorDepth` cells more (64 by default), never round a cycle. Its text writes an error as the book's cells show it and an empty cell as `empty`.

A grid under a filter draws rows that are not next to each other, and a form some columns of a table. `rows` and `columns` give only those lines of `range`, by their numbers on the sheet (from 1, in ascending order, each once), in that order: `cells` has the cells of the lines given, `merges` and `tables` those that meet them, whole, a merged area that reaches a row left out too, and `maxCells` counts the cells given. `editText: true` adds to each cell the text it opens with for editing, as Excel's formula bar shows it: a formula with `=` in the book's locale, a value as `entryText` writes it by the format the cell shows, with an apostrophe before text that would read as something else, in every cell of an old-kind array formula's area its formula, as F2 opens it, and nothing for a cell without a value or a formula of its own (an empty one, one a dynamic array spilled into).

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 'Code' }, A2: { value: '007' }, A3: { value: 5 }, A4: { formula: 'SUM(A3,1)' } }, merges: ['B1:B2'] }] })
const v = session.getViewport('S', 'A1:B4', { rows: [2, 4], editText: true })
console.log(v.rows, Object.keys(v.cells), v.cells.A2.editText, v.cells.A4.editText, v.merges)
// → [ { row: 2, height: 15 }, { row: 4, height: 15 } ] [ 'A2', 'A4' ] '007 =SUM(A3,1) [ 'B1:B2' ]
```

`viewportOf(saved, sheet, range, opts?)` gives the same window over a book and its result as a host saved them, `{ book, result }` (what `getSnapshot()` gives), without a session and without computing anything: a server shows a saved version of a book, or the book with a change not yet taken, in the grid that shows the live one. The values, the text, the colour and the reasons of errors come from the result, the formats, styles and layout from the book, and the options are `getViewport`'s. The window has no `revision`. A result holds no `appearance` of conditional formats, so the window has none by default and a cell's `text` is the result's, by the cell's own format, not by a rule's; with `appearance: true` the rules are computed as a session computes them, over the result's values and on its clock (TODAY in a rule is the version's day), without computing a formula, and each cell has the `appearance` and the text by a rule's number format that `getViewport` gives in a session of the book with that clock ([what can differ](conditional-formats.md#a-saved-version)). A cycle's reasons name no cell, so the issue of a cell a cycle stops begins, for `issue`, at the cell itself. The first window over a book and a result checks the book and finds where each sheet ends, a pass over its cells (0.24 s on the synthetic book of 34,000 rows); the later ones over the same two objects cost what the window holds (5 ms for 20,000 cells, 11 ms with `editText`), so a host keeps the objects of a version while it serves its windows, and changes neither. The first window with `appearance` also reads every value of the result into the state the rules are computed over (10 MB on the synthetic book of 34,000 rows), kept with the objects too: on that book with a formula rule over 4 columns, a three-colour scale and a data bar, a window of 20,000 cells took about 0.16 s more than without it the first time (0.55 s against 0.39 s at a load average of 5 to 9; 0.03 s more on 3,400 rows), and a later one 15 ms over rows read before and 27 ms over new rows, whose rule formulas are made then, against 11 to 12 ms without it.

A grid needs a few things beside its window, and the session gives them without the whole book:

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{
    name: 'Plan',
    cells: { A1: { value: 3 }, B1: { formula: 'SEQUENCE(A1)' }, B3: { numberFormat: '0.0' }, D1: { formula: 'A1*10', array: 'D1:D2' }, D2: { style: 0 } },
    merges: ['F1:G2'],
  }],
  styles: [{ font: { bold: true } }],
})
console.log(session.getUsedRange('Plan')) // → { rows: 3, columns: 7 }
console.log(session.getEnd('Plan', 'B1', 'down'), session.getEnd('Plan', 'D1', 'right')) // → B3 XFD1
console.log(session.getCells('Plan', ['B3', 'D2']))
// → {
//     D2: { value: 30, text: '30', spillOf: 'D1', style: 0 },
//     B3: {
//       value: 3,
//       text: '3.0',
//       numberFormat: '0.0',
//       spillOf: 'B1',
//       numberFormatCode: '0.0'
//     }
//   }
console.log(session.getCells('Plan', 'D1').D1.array, session.getLayout().sheets) // → D1:D2 [ { name: 'Plan', merges: [ 'F1:G2' ] } ]
```

`getUsedRange(sheet)` says where the sheet's cells end: the last row and column reached by a cell of its input (a value, a formula, or a style, a format, meta, a note or a link alone, a cell of the sheet's looks too), a spilled array (its whole range, so `B1` above reaches row 3), a merged area or a table. Settings of whole rows and columns, conditional formats and data validation do not count, as they usually cover whole lines. It costs the sheet's columns, spills and cells without a value, not a pass over all its cells. `getEnd(sheet, cell, direction)` is where Control and an arrow key go from a cell (`'up'`, `'down'`, `'left'`, `'right'`), as Excel's `Range.End` goes, measured in Excel for Windows: to the last cell of a block (the spill down to B3 above), else the next cell that holds something, else the sheet's edge (XFD1 above: a merged area without a value is empty), with hidden rows and columns not there. It reads the cells of the one line, 0.2 ms down a column of 34,000 values ([rules](live-books.md#where-control-and-an-arrow-key-go)). `getLayout()` is the book without its cells: the styles, names and settings, and every field of each sheet but `cells` and `looks` (columns, rows, merged areas, tables, frozen panes, conditional formats, data validation, outline settings, whether it is hidden), at a cost that does not depend on the cells. `getCells(sheet, ranges)` reads several rectangles at once, each cell once, and `getCell(sheet, cell)` one cell, undefined where `getCells` leaves it out. A view of an old-kind array formula's top-left cell has its `array`, beside its `spill`. A view with a number format of its own, and a change whose format changed to one, have the format as a code too (`numberFormatCode`), as the window gives it: a built-in format by the code of the book's locale (14 is `dd.mm.yyyy` in ru-RU), a code as it is.

## Layout

`setLayout(edits)` changes what a sheet has beside its cells, in one call and one step of the undo: the width, `bestFit` and hidden state of columns, the height, `bestFit` and hidden state of rows, the style and number format of whole rows and columns, outline levels and collapsed groups (`{ sheet, columns: 'B:D', set }`, `{ sheet, rows: '2:5', set }`, a field set to null removed), AutoFit of columns and rows (`{ sheet, columns, fit: true }`, `{ sheet, rows, fit: true }`), and the sheet's merged areas, filter mode, frozen panes, conditional formats, rules of data validation and outline settings (`{ sheet, merges }`, `{ sheet, filtered }`, `{ sheet, frozen }`, `{ sheet, conditionalFormats }`, `{ sheet, validations }`, `{ sheet, outline }`, each list replacing the sheet's whole list).

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{
    name: 'S',
    cells: { A1: { value: 1234567.891, numberFormat: '#,##0.00' }, A2: { value: 2 }, A3: { value: 4 }, B1: { formula: 'SUBTOTAL(109,A1:A3)' } },
    rows: { '2:3': { outlineLevel: 1 } },
  }],
})
const r = session.setLayout([
  { sheet: 'S', columns: 'A', set: { width: 6 } },
  { sheet: 'S', rows: '2:3', set: { hidden: true } },
  { sheet: 'S', rows: '4', set: { collapsed: true } },
])
console.log(r.changes.map((c) => [c.cell, c.text]), r.layout) // → [ [ 'B1', '1234568' ], [ 'A1', '######' ] ] [ 'S' ]
session.setLayout([{ sheet: 'S', columns: 'A:B', fit: true }])
console.log(session.getLayout().sheets[0].columns)
// → {
//     A: { width: 11.6640625, bestFit: true },
//     B: { width: 12.1640625, bestFit: true }
//   }
```

Excel's outline commands, Group, Ungroup, Show and Hide Detail and the level buttons, are structure operations ([Data validation and outlines](validation.md)). It recomputes what the layout changes and nothing else: hidden rows and the filter mode the `SUBTOTAL` and `AGGREGATE` formulas over them (`B1` above leaves out the hidden rows), a merged area added or removed the arrays whose spill range it meets. With the display on, the answer lists the cells whose shown text or colour changed with them: by a width (`A1` above shows `#` across a column too narrow for it), a row's or a column's format, a merged area; `layout` names the sheets whose layout changed, for a grid to read `getLayout()` again. AutoFit gives each column with something to show the width of its widest cell, as Excel for Mac fits it, and marks it `bestFit` ([Showing values](formats.md) has how). The look of cells and the style table change through `patchCells` and `setStyles` ([Editing a grid](edits.md)), and recompute nothing.

A host that keeps a sheet's `columns` or `rows` itself, beside a session or without one, patches them as `setLayout` does with `patchLines(map, span, patch)` (from `xlcalc` and `xlcalc/a1`): the span is `C`, `B:D`, `5` or `5:9`, the axis by its shape; the patch goes over each line's entry, null removing a field, and the answer is the new map `setLayout` would write, spans the patch cuts split and lines left without a field dropped.

```js
import { patchLines } from 'xlcalc/a1'

console.log(patchLines({ 'B:F': { width: 10 } }, 'D', { hidden: true }))
// → {
//     'B:C': { width: 10 },
//     D: { width: 10, hidden: true },
//     'E:F': { width: 10 }
//   }
console.log(patchLines({ '5:9': { height: 20, outlineLevel: 1 } }, '7:12', { height: null }))
// → { '5:6': { height: 20, outlineLevel: 1 }, '7:9': { outlineLevel: 1 } }
```

## Looks of many cells

A sheet formatted across many empty cells, such as whole rows to column XFD in a model laid out for printing, keeps that look in `looks`: rectangles, each with a style and a number format, a later one over the earlier ones. The book holds no object per cell for them, so a sheet of six million formatted cells reads, opens a session and goes to a thread in a few megabytes. A cell without an entry in `cells` shows the style and the number format of the last rectangle over it, both from that rectangle (one without a field gives the cell none of its own, one without either clears the look); a field the cell has none of comes from its row, else its column, as for an entry. A cell with an entry shows its own fields, whatever the rectangles say: `B2` below holds a value and no look of its own.

```js
import { createSession } from 'xlcalc'

const session = createSession({
  styles: [{ font: { bold: true } }, { fill: { color: '#FFF2CC' } }],
  sheets: [{
    name: 'Plan',
    cells: { A1: { value: 'Region', style: 0 }, B2: { value: 0.25 } },
    looks: [{ range: 'A2:XFD4', numberFormat: '#,##0' }, { range: 'B2:B4', style: 1, numberFormat: '0.0%' }, { range: 'XFD4' }],
  }],
})
console.log(session.getCell('Plan', 'B3'), session.getCell('Plan', 'B2'))
// → {
//     value: null,
//     text: '',
//     numberFormat: '0.0%',
//     style: 1,
//     numberFormatCode: '0.0%'
//   } { value: 0.25, text: '0.25' }
console.log(session.getUsedRange('Plan'), Object.keys(session.getResult().sheets[0].cells)) // → { rows: 4, columns: 16384 } [ 'A1', 'B2' ]
console.log(session.getBook().sheets[0].looks)
// → [
//     { range: 'A2:A3', numberFormat: '#,##0' },
//     { range: 'C2:XFD3', numberFormat: '#,##0' },
//     { range: 'B3', style: 1, numberFormat: '0.0%' },
//     { range: 'A4', numberFormat: '#,##0' },
//     { range: 'B4', style: 1, numberFormat: '0.0%' },
//     { range: 'C4:XFC4', numberFormat: '#,##0' }
//   ]
session.setCellText('Plan', 'B3', '12')
console.log(session.getBook().sheets[0].cells.B3) // → { style: 1, value: 0.12, numberFormat: '0.0%' }
```

An entry with nothing but a style or a number format is the same thing, and a book is taken with each such entry moved into its sheet's `looks`: a book with looks and the same book with an entry in each of their cells compute, show, hash and write alike. The books the calls give back (`getBook()`, `getSnapshot()`, `editStructure`, `readXlsx`) hold the looks in one form for one set of them, as above: rows in blocks, the runs of a group of rows as rectangles across it, then the rows of the group that differ. A cell of the looks is a cell with a look of its own wherever cells are read: a window, `getCells`, `readCells` and `findCells`, `getUsedRange`, text typed into it, its conditional formats. A result lists no such cell, as it holds no value, but where a large spill's unstored part covers it, as a result listed an entry there. Typed over, a cell keeps its look, as in Excel; `setCells` with null keeps it, `replaceCells` with null takes it away, and `patchCells`, `clearCells`, structure operations and undo change the looks' cells as they change entries. They change only the rows they touch: rows and columns inserted and deleted move the looks by their rows and runs, the new lines with the look of the line before them, and a range copied or moved takes the looks of its cells and of where it lands. On a sheet of the wide shape, 5.6 million cells of looks, a row inserted or deleted takes a few milliseconds and its undo as little, a copy of 344 × 100 of those cells 0.1 s, a row patched to XFD 0.1 s and the formats of three rows cleared 0.3 s, each leaving a few megabytes of heap at most (`bench/memory/wide-shape.ts edits`). The input hash reads a sheet's looks after its cells, so a book without them hashes as before.
