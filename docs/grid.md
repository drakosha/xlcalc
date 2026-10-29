# Rendering a grid from a session

What a grid reads from a session to draw a sheet: a viewport of cells with their formatting, the sizes of rows and columns, merged areas and tables; and how it changes the layout and the formatting of many cells at once.

## A viewport for a grid

A grid draws a viewport of a sheet with `getViewport(sheet, range, opts?)`. A single call returns the cells with their formatting resolved, the sizes of the viewport's rows and columns, the merged areas and tables that overlap it, the frozen panes and where the sheet's data ends.

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

A cell of the viewport carries what `getCells` returns (the value, the displayed `text` and `color`, the formula, the spill, `meta`, a note, a hyperlink, the range of a legacy array formula) together with its resolved formatting:
- its own `style` and the `effectiveStyle` it displays: its own, otherwise its row's, otherwise its column's, as indexes into the workbook's styles, the same way `getCells` and the changes of edits report a style; the style objects the viewport uses are in `styles`, keyed by index;
- the `numberFormat` it displays with, as a code, with `numberFormatSource` saying whether the format is the cell's, the row's or the column's (a built-in format is given by its code in the workbook's locale: 14 is `dd.mm.yyyy` in ru-RU);
- `appearance`, what its conditional formats apply to it.

The style and the number format are resolved separately, so a cell in a bold row of a column formatted `0.00` is bold and displays `0.00`. An empty cell is listed when it displays something: a style of its row or column (`C1` above), a number format of its row or column (an empty cell in a column formatted `0.00` has `numberFormat: '0.00'` and `numberFormatSource: 'column'`, as Excel's list of formats shows it), an appearance, a note or a hyperlink; so are the cells of a spill, including a large array's cells beyond its data. Where a formatted row crosses a formatted column, a cell displays the row's style, as Excel displays an empty cell of a file there.

The rest of the viewport:
- `columns` and `rows` give each line's number (`column`, `row`, from 1), width and height, with Excel's default where none is set (8.83203125 and 15 points), and `hidden`;
- `merges` lists the merged areas that overlap the viewport, in full;
- `tables` lists the tables that overlap it, in full, each with whether its first row is a header row and its last row a totals row;
- `frozen` gives the sheet's frozen rows and columns, 0 for none;
- `usedRange` is `getUsedRange(sheet)`.

`formulaLocal: true` adds each formula as Excel displays it in the workbook's locale, and inside a table as Excel displays it there (`[@Qty]`). `diagnostics: true` adds an `issue` to each cell with an error, as `getCellIssue` returns it. A viewport holds at most `maxCells` cells, 20,000 by default: for a larger range it returns as many of the first rows as fit, with `truncated: true` and the `range` it returned. Its cost depends on the size of the viewport, not of the workbook: on the synthetic workbook of 34,000 rows a viewport of 20,000 cells takes 15 ms, the same as on the workbook of 3,400 rows and slightly less than `getCells` and `getLayout()` together, and 26 ms from a worker thread.

`getCellIssue(sheet, cell)` explains why a cell has an error without building the workbook's result. It returns the error, the reason, and for a formula that does not parse the error code and position, as the result's `diagnostics` give them, and `origin`, the cell where the error originated, found by following the cells the reasons name (in a cycle, the first cell of the cycle that the path returns to). A formula stopped by a circular reference has an issue too, `CYCLE`, whatever value it holds. `explain` returns a structured `root`: local nodes carry `location: { sheet, range }`, the invariant `formula`, `value` and `children`, so a panel does not have to parse the rendered `text`. The asynchronous `explain` also accepts `signal` and `timeoutMs`, and a cancelled call, queued or running, rejects with `CalcError` code `aborted` ([Async calls](async.md#calls)). `explain` marks each cell with an error and leads to the cell where it originated: through a range by the cell the error came from, and past `depth` for up to `errorDepth` more cells (64 by default), never around a cycle. Its text writes an error as the workbook's cells display it, and an empty cell as `empty`.

A grid under a filter draws rows that are not adjacent, and a form draws only some columns of a table. `rows` and `columns` restrict the viewport to those lines of `range`, by their numbers on the sheet (from 1, in ascending order, each once): `cells` contains the cells of those lines, `merges` and `tables` the ones that overlap them, in full, including a merged area that extends into a row left out, and `maxCells` counts only the cells returned. `editText: true` adds to each cell the text it opens with in edit mode, as Excel's formula bar shows it: a formula with `=` in the workbook's locale; a value as `entryText` writes it in the format the cell displays, with an apostrophe before text that would otherwise be parsed as something else; the formula in every cell of a legacy array formula's range, as F2 opens it; and nothing for a cell without a value or formula of its own (an empty cell, or one a dynamic array spilled into).

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 'Code' }, A2: { value: '007' }, A3: { value: 5 }, A4: { formula: 'SUM(A3,1)' } }, merges: ['B1:B2'] }] })
const v = session.getViewport('S', 'A1:B4', { rows: [2, 4], editText: true })
console.log(v.rows, Object.keys(v.cells), v.cells.A2.editText, v.cells.A4.editText, v.merges)
// → [ { row: 2, height: 15 }, { row: 4, height: 15 } ] [ 'A2', 'A4' ] '007 =SUM(A3,1) [ 'B1:B2' ]
```

`viewportOf(saved, sheet, range, opts?)` returns the same viewport over a workbook and its result as a host saved them, `{ book, result }` (what `getSnapshot()` returns), without a session and without calculating anything. A server can use it to display a saved version of a workbook, or the workbook with a pending change, in the same grid that displays the live one. The values, the text, the color and the reasons of errors come from the result; the formats, styles and layout come from the workbook; the options are those of `getViewport`. The viewport has no `revision`.

A result holds no `appearance` from conditional formats, so by default the viewport has none either, and a cell's `text` is the result's, formatted by the cell's own number format rather than a rule's. With `appearance: true` the rules are evaluated as a session evaluates them, over the result's values and on its clock (TODAY in a rule is the day of the version), without calculating any formula, and each cell gets the `appearance` and the text in a rule's number format that `getViewport` returns in a session of the workbook with that clock ([what can differ](conditional-formats.md#a-saved-version)). The reasons of a cycle name no cell, so for a cell stopped by a cycle the `issue` starts at the cell itself.

The first viewport over a workbook and a result validates the workbook and finds where each sheet's data ends, which takes one pass over its cells (0.24 s on the synthetic workbook of 34,000 rows). Later viewports over the same two objects cost only what they contain (5 ms for 20,000 cells, 11 ms with `editText`), so a host should keep the objects of a version while it serves viewports of it, and modify neither. The first viewport with `appearance` also loads every value of the result into the state the rules are evaluated over (10 MB on the synthetic workbook of 34,000 rows) and keeps it with the objects. On that workbook, with a formula rule over 4 columns, a three-color scale and a data bar, a viewport of 20,000 cells took about 0.16 s longer than without `appearance` the first time (0.55 s against 0.39 s at a load average of 5 to 9; 0.03 s longer on 3,400 rows). A later one took 15 ms over rows read before and 27 ms over new rows, whose rule formulas are compiled then, against 11 to 12 ms without it.

Besides its viewport, a grid needs a few other things, and the session returns them without the whole workbook:

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

`getUsedRange(sheet)` returns where the sheet's data ends: the last row and column reached by a cell of its input (a value, a formula, or only a style, a number format, meta, a note or a hyperlink, including cells formatted through the sheet's `looks`), a spilled array (its entire range, so `B1` above reaches row 3), a merged area or a table. Formatting of entire rows and columns, conditional formats and data validation are not counted, as they usually cover entire lines. Its cost depends on the sheet's columns, spills and cells without a value; it does not scan all the cells.

`getEnd(sheet, cell, direction)` returns where Ctrl and an arrow key go from a cell (`'up'`, `'down'`, `'left'`, `'right'`), as Excel's `Range.End` does, verified in Excel for Windows: to the last cell of a block (down to B3 through the spill above), otherwise to the next cell that contains something, otherwise to the edge of the sheet (XFD1 above: a merged area without a value is empty), skipping hidden rows and columns. It reads only the cells of that one row or column, 0.2 ms down a column of 34,000 values ([rules](live-books.md#where-control-and-an-arrow-key-go)).

`getAutoSum(sheet, cell)` returns what Excel's AutoSum (Home > Editing, Alt+=) offers at a cell: the range it sums and the formula, as the engine writes it and in the workbook's locale (`{ range: 'A2:A4', formula: 'SUM(A2:A4)', formulaLocal: 'СУММ(A2:A4)' }`). Going up from the cell, it passes the cells right above that hold no number (blanks, text), takes the run of numbers there and sums from its first cell down to the cell, so a header stays out and a blank or a label right above the cell comes in (`SUM(A1:A4)` at A5 below 1, 2, 3 and a blank). Without numbers above, it does the same to the left. A date or a time is no number for it. A `SUM` or `SUBTOTAL` formula in the run is a subtotal: the run ends there (`SUM(A4)` below a subtotal and one number), and when the cell right above is one, the subtotals alone are summed, nearest first (`SUM(A6,A3)`). Without numbers the formula is `SUM()`, which Excel opens for editing and refuses to enter. A host types `formulaLocal` into the cell with `=` (`setCellText`) or opens it for editing with the range pointed; Typed so, the formula takes the number format of its range's first cell, as Excel gives it: `0%` for a sum of percentages ([Formulas](entry.md#formulas)).

`getSpecialCells(sheet, range, kind)` returns the cells Excel's Go To Special (Home > Editing > Find & Select) selects, as rectangles a grid selects at once (`['B2:D2', 'B4', 'C6']`), none where Excel says no cells were found. `kind` is `{ type: 'blanks' }`, `{ type: 'constants' }` or `{ type: 'formulas' }` with the kinds of values of the dialog's boxes (`values: ['numbers', 'text', 'logicals', 'errors']`, all of them without it; a formula by the value it answers, empty text being text), or `{ type: 'visible' }`. A range of one cell stands for the sheet's used range, as in Excel: from its first row and column of cells to its last, a row or a column of its own (hidden, of a height, a style) reaching it at column A or row 1, so blanks before the first cell are found; a larger range is taken within the used range (`C:C` finds blanks only down to the last row used). Visible cells are those of the range itself in no hidden row or column, of the whole sheet for one cell. The cells a spill fills past its formula are neither constants nor formulas, nor blank. ([How it was checked](development.md#go-to-special).)

`getLayout()` returns the workbook without its cells: the styles, names and settings, and every field of each sheet except `cells` and `looks` (columns, rows, merged areas, tables, frozen panes, conditional formats, data validation, outline settings, and whether it is hidden), at a cost that does not depend on the cells. `getCells(sheet, ranges)` reads several rectangles at once, each cell once, and `getCell(sheet, cell)` reads one cell, returning undefined where `getCells` would leave it out. The top-left cell of a legacy array formula has its `array` besides its `spill`. A cell with a number format of its own, and a change whose number format became one, also carry the format as a code (`numberFormatCode`), as the viewport gives it: a built-in format by its code in the workbook's locale (14 is `dd.mm.yyyy` in ru-RU), any other code as it is.

## Layout

`setLayout(edits)` changes everything a sheet has besides its cells, in one call and one undo step:
- the width, `bestFit` and hidden state of columns, and the height, `bestFit` and hidden state of rows;
- the style and number format of entire rows and columns, outline levels and collapsed groups (`{ sheet, columns: 'B:D', set }`, `{ sheet, rows: '2:5', set }`, where a field set to null is removed);
- AutoFit of columns and rows (`{ sheet, columns, fit: true }`, `{ sheet, rows, fit: true }`);
- the sheet's merged areas, filter mode, frozen panes, conditional formats, data validation rules and outline settings (`{ sheet, merges }`, `{ sheet, filtered }`, `{ sheet, frozen }`, `{ sheet, conditionalFormats }`, `{ sheet, validations }`, `{ sheet, outline }`), each list replacing the sheet's entire list.

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

Excel's outline commands (Group, Ungroup, Show Detail, Hide Detail and the level buttons) are structure operations ([Data validation and outlines](validation.md)). `setLayout` recalculates what the layout change affects and nothing else: hidden rows and the filter mode affect the `SUBTOTAL` and `AGGREGATE` formulas over them (`B1` above ignores the hidden rows), and a merged area added or removed affects the arrays whose spill range it overlaps. With display enabled, the result lists the cells whose displayed text or color changed: through a width (`A1` above shows `#` across a column that is too narrow for it), the format of a row or a column, or a merged area. `layout` names the sheets whose layout changed, so that a grid can call `getLayout()` again. AutoFit sets each column that has content to the width of its widest cell, as Excel for Mac fits it, and marks it `bestFit` ([Number formats and displayed text](formats.md) has the details). Cell formatting and the style table change through `patchCells` and `setStyles` ([Editing a grid](edits.md)), which recalculate nothing.

A host that keeps a sheet's `columns` or `rows` itself, alongside a session or without one, can patch them the way `setLayout` does with `patchLines(map, span, patch)` (from `xlcalc` and `xlcalc/a1`). The span is `C`, `B:D`, `5` or `5:9`, and its form tells columns from rows. The patch is applied to each line's entry, null removing a field, and the function returns the new map that `setLayout` would write: spans the patch cuts through are split, and lines left without fields are dropped.

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

## Formatting of many cells

A sheet formatted across many empty cells, such as entire rows out to column XFD in a model laid out for printing, keeps that formatting in `looks`: rectangles, each with a style and a number format, where a later one overrides earlier ones. The workbook holds no object per cell for them, so a sheet with six million formatted cells loads, opens in a session and moves to a worker thread in a few megabytes.

A cell without an entry in `cells` displays the style and the number format of the last rectangle that covers it, both taken from that rectangle (a rectangle without one of the fields gives the cell none of its own, and one without either clears the formatting). A field the cell has none of comes from its row, otherwise from its column, as for an entry. A cell with an entry displays its own fields, whatever the rectangles say: `B2` below contains a value and has no formatting of its own.

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

An entry with nothing but a style or a number format means the same thing, and when a workbook is loaded each such entry is moved into its sheet's `looks`: a workbook with `looks` and the same workbook with an entry in each of their cells calculate, display, hash and save identically. The workbooks the API returns (`getBook()`, `getSnapshot()`, `editStructure`, `readXlsx`) hold the formatting in one canonical form for a given set of rectangles, as above: rows in blocks, the runs across a group of rows as rectangles, then the rows of the group that differ.

A cell covered by `looks` counts as a cell with formatting of its own wherever cells are read: in a viewport, `getCells`, `readCells` and `findCells`, `getUsedRange`, text typed into it, and its conditional formats. A result does not list such a cell, as it holds no value, except where the unstored part of a large spill covers it, just as a result lists an entry there. When the user types over such a cell, it keeps its formatting, as in Excel; `setCells` with null keeps it, `replaceCells` with null removes it, and `patchCells`, `clearCells`, structure operations and undo change these cells the same way they change entries. They change only the rows they touch: inserted and deleted rows and columns shift the rectangles, new lines take the formatting of the line before them, and a copied or moved range takes the formatting of its cells and of its destination. On a sheet of the wide shape, with 5.6 million formatted cells, inserting or deleting a row takes a few milliseconds and undoing it as little, copying 344 × 100 of those cells 0.1 s, patching a row out to XFD 0.1 s and clearing the formats of three rows 0.3 s, each leaving at most a few megabytes on the heap (`bench/memory/wide-shape.ts edits`). The input hash reads a sheet's `looks` after its cells, so a workbook without them hashes as before.
