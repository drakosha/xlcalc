# xlcalc

A formula engine for spreadsheet-like workbooks in JavaScript. You give it a book (sheets of cells with values and Excel formulas, plus named ranges); it gives back the computed values, the errors with their reasons, and a deterministic hash of the values. Formulas use Excel syntax: English function names, commas between arguments, a decimal point; [Localization](#localization) shows them and reads them in another language, as Excel does.

What it does differently:

- **Compiles each formula once.** By default a formula becomes a tree of JavaScript closures, with no code generated, so a strict Content Security Policy does not stand in the way; with `execution: 'eval'` it becomes JavaScript text, and one `new Function` call compiles the book. A formula that exceeds the JS compiler's nesting limit falls back to closures, preserving evaluation order. Both modes give the same values and hashes, and the tests compute every case book both ways. Cells are evaluated over the reference graph in topological order. A session keeps the compiled book and recomputes only the cells that depend on an edit. On a synthetic book with 3,400 data rows and 3,907 formulas, mostly SUMIFS over whole columns, a cold compute takes about 0.3 s with the cells' shown text (0.28 s without) and an edit of a criterion cell under 10 ms (p95) in either mode on an Apple M5 Pro with Node 26; `test/budgets.test.ts` enforces 1.5 s and 100 ms.
- **Deterministic values and hashes.** The same book with the same clock gives the same values and the same `hash` (SHA-256 of a canonical text of all values plus the engine configuration version), regardless of input order, process or time zone. `inputHash` identifies the input.
- **A clock for volatile functions.** `TODAY` and `NOW` come from `clock.now` (UTC), `RAND` and `RANDBETWEEN` from `clock.seed` and the cell address. Without a clock the engine takes the current moment and a random seed, and `result.clock` reproduces the result.
- **A whitelist of functions, checked against Microsoft Excel.** A function is admitted only when it gives Excel's answer on every one of its test cases, typical calls and edge cases (text and booleans in numeric arguments, blanks, errors, date text, ranges with mixed content), computed in Excel 16 with the en-US locale; [Functions](#functions) has the list. Any other function is reported as `unknown_function` (`#NAME?`) rather than computed with different semantics.
- **Arrays inside a formula.** Operators and functions work element by element over ranges and arrays with Excel 365's rules (`SUM(B:B*C:C)`, `SUMPRODUCT((A1:A9="x")*B1:B9)`, `SUM(IF(A1:A9>0,A1:A9))`), array constants `{1,2;3,4}` are read, and functions with an array answer (`SORT`, `FILTER`, `UNIQUE`, `SEQUENCE`, `MMULT`, …) can be used inside a formula. A formula whose result is an array spills into the cells to its right and below, as in Excel 365 (`#SPILL!` where they are taken), and `E2#` refers to the range it filled (see [Arrays](#arrays)).
- **Cells shown as Excel shows them.** Each cell of a result carries, beside its value, the text Excel shows in it: its number format in the book's locale, General narrowed to the column's width, `#` across a column a number does not fit, the colour of the format's section (see [Showing values](#showing-values)). The value stays the number; `display: false` leaves the text out. A session also gives what a sheet's conditional formats show in each cell of a window: fills, fonts, colour scales, data bars and icons by the current values.
- **Errors are values.** An invalid formula never throws: its cell gets an Excel error, and the result's `diagnostics` say why, with a check code and the position in the formula.
- **xlsx in and out.** `xlcalc/xlsx` reads an xlsx file into a book, with the values Excel stored and a list of what the book cannot hold, and writes a book with its computed values back to xlsx. It has no dependencies and loads none of the engine's computation (see [Reading and writing xlsx](#reading-and-writing-xlsx)).
- **Computing without blocking.** `createInlineCalc` computes in the same thread in portions that give way to the event loop, `createCalcPool` from `xlcalc/worker` in Node worker threads or Web Workers, with soft and hard deadlines, cancellation, a thread's memory limit and limits on the size of a book; a session survives the loss of its thread (see [Computing without blocking](#computing-without-blocking)).
- **No Node APIs.** It runs in Node 20.19 or newer and in browsers (see [Limitations](#limitations) about Content Security Policy).

Function implementations are ported from [formula.js](https://github.com/formulajs/formulajs) (MIT), [jStat](https://github.com/jstat/jstat) (MIT) and [bessel](https://github.com/SheetJS/bessel) (Apache-2.0), with fixes verified against Microsoft Excel; the rest (criteria and database functions, the lookups, `TEXT` and `VALUE`, dates, `IF`, the volatile functions) are written here.

## Install

```sh
npm install xlcalc
```

The package is ESM with TypeScript types included. From CommonJS use `await import('xlcalc')`, or `require()` on Node versions that can require ES modules (20.19+ and 22.12+). It has no dependencies.

### What to import

The package is split where hosts' needs part, so a bundle takes only what it uses:

| for | import |
|---|---|
| computing a book | `compute` from `xlcalc` |
| a session over a book: edits, structure operations, undo, AutoFit, conditional formats, data validation, what cells show | `createSession` from `xlcalc` |
| reading and writing xlsx files | `readXlsx`, `writeXlsx` and `fromLegacy` from `xlcalc/xlsx` |
| computing in threads | `createCalcPool` from `xlcalc/worker`, and `createInlineCalc` from `xlcalc` for portions in the same thread |
| cell addresses in A1 style, without the engine | `xlcalc/a1` |
| the references of formula text as it is typed, for colouring them, and a number format's category as the ribbon shows it, without the engine | `xlcalc/text` |
| books and formulas in ru-RU | `ruRU` from `xlcalc/locales/ru-RU`, loaded once by `registerLocale(ruRU)` from `xlcalc` ([Localization](#localization)) |
| types alone | `xlcalc/contract` |

en-US is built in. Bundled by esbuild for the browser and minified, `compute` alone takes 254 KB (97 KB gzip), `createSession` 395 KB (148 KB), the whole root 419 KB (157 KB), `xlcalc/xlsx` 174 KB (65 KB), `xlcalc/worker` 21 KB (8 KB) in the page and the thread it starts 499 KB (168 KB), `xlcalc/a1` 8 KB (3 KB), `xlcalc/text` 45 KB (17 KB), the ru-RU module 92 KB (15 KB); `test/bundle-size.test.ts` keeps `compute` under 300 KB.

## Quick start

```js
import { compute } from 'xlcalc'

const cells = {
  A1: { value: 'North' }, B1: { value: 120 },
  A2: { value: 'South' }, B2: { value: 80 },
  A3: { value: 'North' }, B3: { value: 45.5 },
  D1: { formula: 'SUMIFS(B:B,A:A,"North")' },
  D2: { formula: 'XLOOKUP("South",A1:A3,B1:B3)' },
  D3: { formula: 'INDIRECT("B2")' },
}
const { sheets, diagnostics } = compute({ sheets: [{ name: 'Sales', cells }] })
console.log(sheets[0].cells.D1.value) // → 165.5
console.log(sheets[0].cells.D2.value) // → 80
console.log(sheets[0].cells.D3.value, diagnostics[0].reason) // → { error: 'NAME' } unknown function INDIRECT
```

A formula is stored without the leading `=` in `formula` (one given is dropped: `{ formula: '=SUM(A1:A2)' }` is `SUM(A1:A2)`, the same input and `inputHash`), separately from `value`, so a string from your data never turns into a formula. So are the formulas of names, tables, conditional formats and data validation. Dates are passed explicitly as `{ date: '2026-01-31' }` and become Excel serial numbers; `numberFormat` and `meta` pass through to the result untouched, and the result's `text` shows the value by its `numberFormat`.

## Recipes

Twelve common tasks, each with the calls it takes. The sections after them have the rules.

### Compute a book and read a cell

```js
import { compute } from 'xlcalc'

const r = compute({ sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*21' } } }] })
console.log(r.sheets[0].cells.A2.value, r.diagnostics) // → 42 []
```

### Open a session, edit a cell and take the changes

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*21' } } }] })
const { revision, changes } = session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
console.log(revision, changes.map((c) => `${c.cell}=${c.value}`)) // → 1 [ 'A1=3', 'A2=63' ]
```

### Type text into a cell as Excel reads it, in ru-RU

```js
import { createSession, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const session = createSession({ locale: 'ru-RU', sheets: [{ name: 'S', cells: {} }] })
session.setCellText('S', 'A1', '15.03.2026')
session.setCellText('S', 'A2', '1 234,5')
console.log(Object.values(session.getCells('S', 'A1:A2')).map((c) => [c.value, c.text, c.numberFormat])) // → [ [ 46096, '15.03.2026', 14 ], [ 1234.5, '1 234,50', 4 ] ]
```

### Paste text from the clipboard

```js
import { createSession, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const session = createSession({ locale: 'ru-RU', sheets: [{ name: 'S', cells: { B2: { value: 'old' } } }] })
const pasted = session.pasteText('S', 'A1', "15.03.2026\t1 234,5\t=B1*2\r\n'007\r\n")
console.log(pasted.range, pasted.changes.map((c) => `${c.cell} ${c.text}`)) // → A1:C3 [ 'A1 #########', 'B1 1 234,50', "A2 '007", 'B2 ', 'C1 2469' ]
console.log(session.undo().call, session.getCell('S', 'A1')) // → pasteText undefined
```

### Copy a rectangle as plain text

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', columns: { A: { width: 3 } }, cells: { A1: { value: 46096, numberFormat: 'yyyy-mm-dd' }, A2: { value: 1.23456789 } } }] })
console.log(JSON.stringify(session.copyText('S', 'A1:A2', { trailingNewline: false }))) // → "2026-03-15\r\n1.23456789"
```

`copyText` prepares the selection from computed values and effective formats without column clipping, skipping hidden rows and columns. The host writes the returned string to its clipboard. Windows ends the copy with CRLF; Mac uses `{ trailingNewline: false }`. Copies over `maxCells` (20,000 by default) or `maxChars` (16 Mi characters) are refused rather than truncated. `formatClipboard` from `xlcalc/text` remains available for a host that already has a rectangle of full formatted texts. [Copying plain text](docs/entry.md#copying-plain-text) covers merges, spills and why every plain-text copy cannot be read back without ambiguity.

### Enter a text into a selection (Control+Return)

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { value: 3 }, A3: { value: 4 }, C1: { numberFormat: '0%' } } }] })
console.log(session.fillCellText('S', 'B1:B3', '=A1*2').changes.map((c) => `${c.cell} ${c.formula} ${c.value}`)) // → [ 'B3 A3*2 8', 'B2 A2*2 6', 'B1 A1*2 4' ]
console.log(session.fillCellText('S', 'C1:C2', '5').changes.map((c) => `${c.cell} ${c.text}`), session.getBook().sheets[0].cells.C2) // → [ 'C1 5%', 'C2 5%' ] { value: 0.05, numberFormat: '0%' }
```

### Read a window for a grid

```js
import { createSession } from 'xlcalc'

const session = createSession({ styles: [{ font: { bold: true } }], sheets: [{ name: 'S', cells: { A1: { value: 'Total', style: 0 }, B1: { value: 1234.5, numberFormat: '#,##0.00' } }, columns: { B: { width: 12 } } }] })
const v = session.getViewport('S', 'A1:B2')
console.log(v.cells.A1.effectiveStyle, v.styles[v.cells.A1.effectiveStyle], v.cells.B1.text, v.columns.map((c) => c.width), v.usedRange) // → 0 { font: { bold: true } } 1,234.50 [ 8.83203125, 12 ] { rows: 1, columns: 2 }
```

### Read a window over a saved version of a book

```js
import { compute, viewportOf } from 'xlcalc'

const rule = { type: 'cellIs', ranges: ['A1:A2'], operator: 'greaterThan', formulas: ['1000'], style: { fill: { color: '#FFC7CE' } } }
const book = { sheets: [{ name: 'S', cells: { A1: { value: 1234.5 }, A2: { formula: 'A1/0' } }, columns: { A: { numberFormat: '#,##0.00' } }, conditionalFormats: [rule] }] }
const saved = JSON.parse(JSON.stringify({ book, result: compute(book) })) // as a host stores a version
const v = viewportOf(saved, 'S', 'A1:A2', { diagnostics: true, editText: true, appearance: true })
console.log(v.cells.A1.text, v.cells.A1.editText, v.cells.A1.appearance, v.cells.A2.issue.reason) // → 1,234.50 1234.5 { style: { fill: { color: '#FFC7CE' } } } division by zero
```

### Make a range bold and set its number format

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 0.25 }, B1: { value: 0.5 } } }] })
const r = session.patchCells([{ sheet: 'S', range: 'A1:B1', set: { style: { font: { bold: true } }, numberFormat: '0%' } }])
console.log(r.changes.map((c) => [c.cell, c.text, c.style]), r.styles) // → [ [ 'A1', '25%', 0 ], [ 'B1', '50%', 0 ] ] [ { font: { bold: true } } ]
```

### Insert a row and move the window by `moved`

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { formula: 'SUM(A1:A2)' } } }] })
const { moved } = session.editStructure({ op: 'insertRows', sheet: 'S', row: 2 })
const rowAfter = (row) => moved.reduce((x, m) => (m.kind === 'lines' && m.axis === 'rows' && x >= m.at ? x + m.count : x), row)
console.log(moved, rowAfter(1), rowAfter(3)) // → [ { kind: 'lines', sheet: 'S', axis: 'rows', at: 2, count: 1 } ] 1 4
```

### Read an xlsx file, edit it and write it back

```js
import { createSession } from 'xlcalc'
import { readXlsx, writeXlsx } from 'xlcalc/xlsx'

const bytes = (await writeXlsx({ book: { sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*21' } } }] } })).data // or await readFile(path)
const session = createSession((await readXlsx(bytes)).book)
session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
const saved = await writeXlsx(session.getSnapshot())
console.log((await readXlsx(saved.data)).cached.sheets[0].cells.A2) // → { value: 63 }
```

The reader also keeps `book.activeSheet` and each sheet's `view` (`activeCell`, `selection`, `topLeft`, `zoom`), so a grid can restore the saved window. `getLayout()` retains them; overlay live view state on the book when saving. They do not affect `inputHash`. [Saved views](docs/xlsx.md#the-saved-active-sheet-and-view) has the file mapping and limits.

### Show a formula in a locale

```js
import { toLocal } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

console.log(toLocal(`IF(A1>0.5,SUM('Data'!B1:B3),"none")`, ruRU).formula) // → ЕСЛИ(A1>0,5;СУММ(Data!B1:B3);"none")
```

### Say why a cell holds an error

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 0 }, A2: { formula: '1/A1' }, A3: { formula: 'A2*2' } } }] })
console.log(session.getCellIssue('S', 'A3'))
// → {
//     error: 'DIV0',
//     reason: "division by zero, from 'S'!A2",
//     origin: { sheet: 'S', cell: 'A2' }
//   }
```

### Compute in a pool of threads with a deadline

```js
import { createCalcPool } from 'xlcalc/worker'

const pool = createCalcPool({ threads: 2 })
const book = { sheets: [{ name: 'S', cells: Object.fromEntries(Array.from({ length: 50_000 }, (_, i) => [`A${i + 1}`, { formula: `SUM(SEQUENCE(${(i % 100) + 1}))` }])) }] }
const r = await pool.compute(book, { timeoutMs: 10_000 })
const late = await pool.compute(book, { timeoutMs: 20 }).catch((e) => e)
console.log(r.sheets[0].cells.A100.value, late.code) // → 5050 soft_timeout
await pool.close()
```

## Usage

### Sessions

Mixed cell, name and formatting operations can be applied atomically with `applyChanges(operations, { baseRevision, clock })`. `previewChanges` returns bounded before/after value impacts, including dependent cells and errors, without changing the book or its history. Use the preview's revision and clock when applying it. `readCells`, `findCells`, `getDependents` and `summarizeRanges` read at one revision where the book lives, with projection and bounded pages instead of a full snapshot. `findCells` finds as Excel's Find with `wildcards`, `wholeCell`, `byColumns`, `skipHidden` and `in: 'edit'` (Look in Formulas) or `'text'` (Values), measured in Excel for Windows; `rows` keeps the rows a grid shows, `after` begins after a cell and goes round as Find Next does, `editText` gives the text a cell opens with, and the session keeps the texts it searched, so a search typed letter by letter costs a match of kept texts (0.02 s to 0.04 s over 510,000 cells). `replaceInCells` is Excel's Replace All over the same options, measured in Excel for Windows too, in one step of the undo. These methods also work in inline sessions and pools. See [live workbook operations and queries](docs/live-books.md) for the operation shapes, clocks, cursors and limits.


A session holds a computed book and recomputes only what an edit changes. `setCells` applies edits atomically and answers the session's revision and the cells whose values changed; `explain` shows where a value comes from, as a tree and as text in the book's locale ([Localization](#localization)).

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{
    name: 'Plan',
    cells: {
      A1: { value: 100 }, A2: { value: 250 },
      B1: { formula: 'SUM(A1:A2)' },
      B2: { formula: 'ROUND(B1*1.2,0)' },
    },
  }],
})

console.log(session.setCells([{ sheet: 'Plan', cell: 'A2', input: { value: 300 } }]))
// → {
//     revision: 1,
//     changes: [
//       { sheet: 'Plan', cell: 'A2', value: 300, text: '300' },
//       { sheet: 'Plan', cell: 'B1', value: 400, text: '400' },
//       { sheet: 'Plan', cell: 'B2', value: 480, text: '480' }
//     ]
//   }
console.log(session.explain('Plan', 'B2').text)
// → Plan!B2 = 480  ← ROUND(B1*1.2,0)
//     Plan!B1 = 400  ← SUM(A1:A2)
//       Plan!A1:A2 (2 cells)
console.log(session.getResult().hash === session.getResult().hash) // → true
session.close()
```

Every call that changes the book answers an `EditResult`: the session's `revision` after it (one more for each such call, `session.revision` reads it) and `changes`, the cells whose value or shown text changed, with a cell's `formula`, own `style`, own `numberFormat`, `note` or `link` when they changed (null when the cell has none now). Changes to `meta` are not listed: the host keeps this data, and history restores it. `getResult()` returns the same shape as `compute`, `getSnapshot()` independent copies of the book and its result at one moment in every kind of session (what `writeXlsx` takes), `getHash()` its values hash alone, `getBook()` the book with every change so far (the session's own objects in the same thread: do not modify them), and `close()` releases the book. Later edits leave a snapshot unchanged, and changing a snapshot cannot change the session. A book an answer gives has the type `Book`, a `BookInput`: it goes back into any call that takes a book. A full `compute` of the edited book with `session.clock` gives the same hash as `session.getResult()`; the test suite checks this after each of 200 random edits.

A session changes its book as a grid does, each kind of change in a section of its own: the cells and their look ([Editing a grid](#editing-a-grid)), typed text ([Typed entry](#typed-entry)), widths, heights and the sheet's other settings ([Layout](#layout)), rows, columns, ranges and sheets ([Structure operations](#structure-operations)), the names (`setNames`) and the style table (`setStyles`); it reads a window without the whole book ([A window for a grid](#a-window-for-a-grid)) and undoes every change ([Undo and redo](#undo-and-redo)).

### A window for a grid

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

`getCellIssue(sheet, cell)` says why a cell holds an error without building the book's result: the error, the reason, the check code and position of a formula that does not read, as the result's `diagnostics` give them, and `origin`, the cell the error began in, found by following the cells the reasons name (in a cycle, the first cell of the cycle the way comes back to); a formula a circular reference stops has one too, `CYCLE`, whatever value it holds. `explain` gives a structured `root`: local nodes carry `location: { sheet, range }`, invariant `formula`, `value` and `children`, for a panel without parsing its rendered `text`. Async `explain` also takes `signal` and `timeoutMs`, and a cancelled queued or running call rejects with `CalcError` code `aborted` ([Async calls](docs/async.md#calls)). `explain` marks each cell with an error by that cell and leads to it: through a range by the cell the error came from, and past `depth` for `errorDepth` cells more (64 by default), never round a cycle. Its text writes an error as the book's cells show it and an empty cell as `empty`.

A grid under a filter draws rows that are not next to each other, and a form some columns of a table. `rows` and `columns` give only those lines of `range`, by their numbers on the sheet (from 1, in ascending order, each once), in that order: `cells` has the cells of the lines given, `merges` and `tables` those that meet them, whole, a merged area that reaches a row left out too, and `maxCells` counts the cells given. `editText: true` adds to each cell the text it opens with for editing, as Excel's formula bar shows it: a formula with `=` in the book's locale, a value as `entryText` writes it by the format the cell shows, with an apostrophe before text that would read as something else, in every cell of an old-kind array formula's area its formula, as F2 opens it, and nothing for a cell without a value or a formula of its own (an empty one, one a dynamic array spilled into).

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 'Code' }, A2: { value: '007' }, A3: { value: 5 }, A4: { formula: 'SUM(A3,1)' } }, merges: ['B1:B2'] }] })
const v = session.getViewport('S', 'A1:B4', { rows: [2, 4], editText: true })
console.log(v.rows, Object.keys(v.cells), v.cells.A2.editText, v.cells.A4.editText, v.merges)
// → [ { row: 2, height: 15 }, { row: 4, height: 15 } ] [ 'A2', 'A4' ] '007 =SUM(A3,1) [ 'B1:B2' ]
```

`viewportOf(saved, sheet, range, opts?)` gives the same window over a book and its result as a host saved them, `{ book, result }` (what `getSnapshot()` gives), without a session and without computing anything: a server shows a saved version of a book, or the book with a change not yet taken, in the grid that shows the live one. The values, the text, the colour and the reasons of errors come from the result, the formats, styles and layout from the book, and the options are `getViewport`'s. The window has no `revision`. A result holds no `appearance` of conditional formats, so the window has none by default and a cell's `text` is the result's, by the cell's own format, not by a rule's; with `appearance: true` the rules are computed as a session computes them, over the result's values and on its clock (TODAY in a rule is the version's day), without computing a formula, and each cell has the `appearance` and the text by a rule's number format that `getViewport` gives in a session of the book with that clock ([what can differ](docs/conditional-formats.md#a-saved-version)). A cycle's reasons name no cell, so the issue of a cell a cycle stops begins, for `issue`, at the cell itself. The first window over a book and a result checks the book and finds where each sheet ends, a pass over its cells (0.24 s on the synthetic book of 34,000 rows); the later ones over the same two objects cost what the window holds (5 ms for 20,000 cells, 11 ms with `editText`), so a host keeps the objects of a version while it serves its windows, and changes neither. The first window with `appearance` also reads every value of the result into the state the rules are computed over (10 MB on the synthetic book of 34,000 rows), kept with the objects too: on that book with a formula rule over 4 columns, a three-colour scale and a data bar, a window of 20,000 cells took about 0.16 s more than without it the first time (0.55 s against 0.39 s at a load average of 5 to 9; 0.03 s more on 3,400 rows), and a later one 15 ms over rows read before and 27 ms over new rows, whose rule formulas are made then, against 11 to 12 ms without it.

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

`getUsedRange(sheet)` says where the sheet's cells end: the last row and column reached by a cell of its input (a value, a formula, or a style, a format, meta, a note or a link alone, a cell of the sheet's looks too), a spilled array (its whole range, so `B1` above reaches row 3), a merged area or a table. Settings of whole rows and columns, conditional formats and data validation do not count, as they usually cover whole lines. It costs the sheet's columns, spills and cells without a value, not a pass over all its cells. `getEnd(sheet, cell, direction)` is where Control and an arrow key go from a cell (`'up'`, `'down'`, `'left'`, `'right'`), as Excel's `Range.End` goes, measured in Excel for Windows: to the last cell of a block (the spill down to B3 above), else the next cell that holds something, else the sheet's edge (XFD1 above: a merged area without a value is empty), with hidden rows and columns not there. It reads the cells of the one line, 0.2 ms down a column of 34,000 values ([rules](docs/live-books.md#where-control-and-an-arrow-key-go)). `getLayout()` is the book without its cells: the styles, names and settings, and every field of each sheet but `cells` and `looks` (columns, rows, merged areas, tables, frozen panes, conditional formats, data validation, outline settings, whether it is hidden), at a cost that does not depend on the cells. `getCells(sheet, ranges)` reads several rectangles at once, each cell once, and `getCell(sheet, cell)` one cell, undefined where `getCells` leaves it out. A view of an old-kind array formula's top-left cell has its `array`, beside its `spill`. A view with a number format of its own, and a change whose format changed to one, have the format as a code too (`numberFormatCode`), as the window gives it: a built-in format by the code of the book's locale (14 is `dd.mm.yyyy` in ru-RU), a code as it is.

### Editing a grid

A grid edits cells as Excel does: a value typed over a cell keeps the cell's look, Delete clears the content, Clear Formats the look, and Bold over a selection goes over what each cell shows. The session does each in one call, recomputes what depends on it and answers what the grid draws again.

| to | call |
|---|---|
| put a value or a formula into cells, keeping their look (typing over a cell, a paste of values) | `setCells([{ sheet, cell, input: { value } }])` |
| replace cells whole, or remove them (`input: null`) | `replaceCells([{ sheet, cell, input }])` |
| set the look, a note or a link of cells, a range, whole rows or columns, without their values (Bold, Fill, a number format) | `patchCells([{ sheet, range, set: { style, numberFormat } }])` |
| clear the content, the look or both (Delete, Clear Formats, Clear All) | `clearCells([{ sheet, range }], 'contents')`, `'formats'` or `'all'` |
| take text a person typed into a cell, as Excel reads it in the book's locale | `setCellText(sheet, cell, text)` |
| take texts typed into several cells, each read as typed into its cell, in one step | `setCellTexts([{ sheet, cell, text }])` |
| paste text from the clipboard at a cell, as Excel pastes it | `pasteText(sheet, cell, text)` |
| enter a text into every cell of a selection (Control+Return) | `fillCellText(sheet, range, text, { active })` |
| change widths, heights, hidden lines, a line's style, merged areas, frozen panes, conditional formats, data validation, or AutoFit columns | `setLayout([{ sheet, columns, set: { width } }])` ([Layout](#layout)) |
| replace the book's table of styles | `setStyles(styles)` |
| insert or delete rows and columns, move or copy a range, fill it as the fill handle does, merge or unmerge cells, insert, rename, move, hide or delete a sheet, add, rename, convert or delete a table | `editStructure(op)` ([Structure operations](#structure-operations)) |

A style by its fields, `{ font: { bold: true } }`, is what `patchCells` takes: it goes over the style each cell shows, and the session finds or adds the style in the book's `styles`. `setCells` and `replaceCells` take a cell's input as the book stores it, a style by its number in `styles`, as `getCells`, `getViewport` and the changes give it. A host that keeps a book's styles itself finds a style's number by the same rule: `internStyle(styles, style)` gives the first equal style's number and the same array, or a new array with a copy of the style at the end; styles are equal when their fields are, in any order.

```js
import { internStyle } from 'xlcalc'

const styles = [{ font: { bold: true } }, { fill: { color: '#FFEE99' }, font: { italic: true } }]
const found = internStyle(styles, { font: { italic: true }, fill: { color: '#FFEE99' } })
console.log(found.index, found.styles === styles) // → 1 true
const added = internStyle(styles, { font: { bold: true, italic: true } })
console.log(added.index, added.styles.length, styles.length) // → 2 3 2
```

```js
import { createSession } from 'xlcalc'

const session = createSession({
  styles: [{ font: { bold: true } }],
  sheets: [{
    name: 'Plan',
    cells: { A1: { value: 'Item', style: 0 }, A2: { value: 2, numberFormat: '0.00', note: { text: 'from the CRM' } }, A3: { formula: 'A2*10' } },
    conditionalFormats: [{ type: 'cellIs', ranges: ['A2:A3'], operator: 'greaterThan', formulas: ['20'], style: { font: { bold: true } } }],
  }],
})
const edit = session.setCells([{ sheet: 'Plan', cell: 'A2', input: { value: 5 } }])
console.log(edit.revision, edit.repaint, session.getBook().sheets[0].cells.A2)
// → 1 [ { sheet: 'Plan', range: 'A2:A3' } ] { numberFormat: '0.00', note: { text: 'from the CRM' }, value: 5 }
const fill = session.patchCells([{ sheet: 'Plan', range: 'A1:B1', set: { style: { fill: { color: '#FFEE99' } } } }])
console.log(fill.changes.map((c) => `${c.cell} ${c.style}`), fill.styles)
// → [ 'A1 1', 'B1 2' ] [
//     { font: { bold: true } },
//     { font: { bold: true }, fill: { color: '#FFEE99' } },
//     { fill: { color: '#FFEE99' } }
//   ]
session.clearCells([{ sheet: 'Plan', range: 'A2:A3' }], 'contents')
console.log(session.getBook().sheets[0].cells.A2) // → { numberFormat: '0.00', note: { text: 'from the CRM' } }
const stale = (() => {
  try {
    session.setCells([{ sheet: 'Plan', cell: 'A3', input: { value: 1 } }], { baseRevision: 1 })
  } catch (e) {
    return e.code
  }
})()
console.log(stale, session.revision) // → stale_revision 3
```

`setCells(edits)` sets the fields an edit gives and keeps the cell's others: null removes a field, and the content is one, so a new `value` takes the place of a formula and its `array`, a new `formula` that of a value. `input: null` clears the content and the hyperlink and keeps the style, the format, the note and `meta`, as Excel's Delete key. `replaceCells(edits)` replaces each cell whole, null removing it. `clearCells(ranges, what)` clears rectangles, whole rows and columns too, as Excel's ClearContents (`'contents'`: values, formulas and hyperlinks), ClearFormats (`'formats'`: each cell's style and number format, so that it shows the plain look whatever its row or column has, and the conditional formats over the cells) and Clear (`'all'`: both, and notes, `meta` and the rules of data validation over the cells); whole rows and columns lose their own style and format too.

`patchCells(patches)` sets the `style`, `numberFormat`, `meta`, `note` or `link` of a `cell` or a `range` (`B2:D9`, whole columns `B:D`, whole rows `2:5`; null removes the field) and keeps values and formulas without being given them: on a cell of an old-kind array formula's area or of a spill it leaves the array as it is, a cell left with nothing is removed, and it computes nothing. A style given by its fields goes over the style each cell shows, its own, else its row's, else its column's, as Excel formats a selection: `{ font: { bold: true } }` makes an italic cell bold and italic, null removes a field (`{ fill: null }`), and equal styles share one entry of the book's `styles`, found or added (`styles` in the answer is the table then); a style by its number replaces the cell's. An empty cell of a rectangle gets a cell with the look, up to 1,048,576 cells (one column's worth; more is an `InputError`). Whole rows and columns are formatted as Excel formats them: the lines' own style and format, the cells of theirs with a look of their own, and the cells where they cross a line of the other direction with one, which get the two merged.

Beside `revision` and `changes`, an answer has `layout`, the sheets whose layout changed (lines, merged areas, tables, rules), for a grid to read `getLayout()` again; `styles`; `moved`, after a structure operation and its undo; and `repaint`, the rectangles of conditional formats whose cells may look otherwise after a value or a rule changed. An edit made over another revision than the session's (`baseRevision`, from an answer, `session.revision` or a window, `getViewport(…).revision`) changes nothing and throws `CalcError` with the code `stale_revision`, so that a host does not apply an edit over one it has not seen. Each call is one step of the undo. [docs/edits.md](docs/edits.md) has the rules and how they were checked against Excel.

### Layout

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

Excel's outline commands, Group, Ungroup, Show and Hide Detail and the level buttons, are structure operations ([Data validation and outlines](#data-validation-and-outlines)). It recomputes what the layout changes and nothing else: hidden rows and the filter mode the `SUBTOTAL` and `AGGREGATE` formulas over them (`B1` above leaves out the hidden rows), a merged area added or removed the arrays whose spill range it meets. With the display on, the answer lists the cells whose shown text or colour changed with them: by a width (`A1` above shows `#` across a column too narrow for it), a row's or a column's format, a merged area; `layout` names the sheets whose layout changed, for a grid to read `getLayout()` again. AutoFit gives each column with something to show the width of its widest cell, as Excel for Mac fits it, and marks it `bestFit` ([Showing values](#showing-values) has how). The look of cells and the style table change through `patchCells` and `setStyles` ([Editing a grid](#editing-a-grid)), and recompute nothing.

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

### Looks of many cells

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

### Undo and redo

A session undoes and redoes its changes, as Excel's Undo and Redo do:

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{ name: 'Plan', cells: { A1: { value: 100 }, A2: { value: 250 }, A3: { formula: 'SUM(A1:A2)' } } }],
})
session.setCells([{ sheet: 'Plan', cell: 'A1', input: { value: 120 } }])
session.editStructure({ op: 'insertRows', sheet: 'Plan', row: 2 })
console.log(session.history) // → { undo: 2, redo: 0 }
const back = session.undo()
console.log(back.call, back.mapCell({ sheet: 'Plan', cell: 'A4' }), back.changes.map((c) => c.formula)) // → editStructure { sheet: 'Plan', cell: 'A3' } [ 'SUM(A1:A2)' ]
console.log(session.undo())
// → {
//     call: 'setCells',
//     revision: 4,
//     changes: [
//       { sheet: 'Plan', cell: 'A1', value: 100, text: '100' },
//       { sheet: 'Plan', cell: 'A3', value: 350, text: '350' }
//     ],
//     cells: [ { sheet: 'Plan', cell: 'A1' } ]
//   }
console.log(session.undo(), session.history) // → null { undo: 0, redo: 2 }
console.log(session.redo().changes.length, session.getBook().sheets[0].cells.A1) // → 2 { value: 120 }
```

`undo()` takes back the last call that changed the book, whole however many cells it changed: `setCells`, `replaceCells`, `clearCells`, `patchCells`, `setCellText`, `setCellTexts` and `fillCellText` (with the widths they gave the columns), `pasteText`, `setNames`, `setStyles`, `setLayout` or `editStructure`. Its answer says which call it was and is that call's answer (`EditResult`), its revision one more: the cells whose value, text, formula or look changed, the `layout`, the `styles`, the `moved` addresses. For the calls that change cells it also lists the cells whose input came back (`cells`), whose note or link a host reads again. For a structure operation it has `mapCell`, where a cell of the book before the undo is now (`A4` above is `A3` again once the inserted row is gone): a host moves its values by it and applies the changes, as after `editStructure`, and reads `getLayout()`, `getUsedRange(sheet)` and its window again, as it does after a change of the layout, the names or the styles. `redo()` does the last undone step again; a new change forgets what `redo` could do. With nothing to take back or do again they answer null, and `history` counts the steps each way, for a toolbar's buttons. `setClock` and a live session's new clock are no steps, and neither is a call that fails: an undo stopped by a deadline or an abort changes nothing.

A session keeps 100 steps, as Excel does, in 64 MB at most by the engine's estimate, or in a pool with `maxHeapMb` its share of an eighth of a thread's heap (8 MB for 384 MB and `maxSessionsPerThread: 6`): the option `undo: { depth, maxMb }` sets both, and `depth: 0` keeps none. A step keeps what its call replaced and no more: the inputs of the cells it changed, the layout of the sheets it changed (the widths, `bestFit` and hidden columns of an AutoFit too), the style table or the names. A structure operation keeps the cells it does not carry back as they were (the rows it deleted, the formulas it rewrote) and the layout of the sheets it changed, not a copy of the book, and undoing it builds the book's state again, as the operation did. A deleted sheet comes back with its cells, which Excel cannot do; a step larger than `maxMb` alone is not kept and the steps before it go, as Excel forgets what it cannot undo. On the synthetic book of 34,000 rows, an edit and its undo cost what the edit costs (107 ms for an amount on the data sheet, which its 2,800 SUMIFS read), `insertRows` keeps 0.4 MB and takes 1.5 s where it takes 1.3 s without a history, its undo and its redo 1.4 to 1.5 s each; `deleteRows` takes 1.5 s where it takes 1.4 s. Deleting the sheet of 506,594 cells would keep 59 MB, over the default, so the session forgets its history then; with `maxMb: 100` it is undone in 1.4 to 1.7 s.

### Following changes

`session.subscribe(listener)` calls the listener after each call that changes the book, whoever makes it: an edit, typed text, a paste, a structure operation, a change of layout or styles, `undo` and `redo`. The listener gets what the call answered of `EditResult` (the revision, the changed cells, the sheets whose layout changed, the style table, how addresses moved, where to repaint) and the call's name in `method`, so a grid, a total shown elsewhere and a second window can all follow one session without passing the answers around:

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{ name: 'Plan', cells: { A1: { value: 100 }, A2: { value: 250 }, A3: { formula: 'SUM(A1:A2)' } } }],
})
const unsubscribe = session.subscribe((change) => {
  console.log(change.method, change.revision, change.changes.map((c) => `${c.cell}=${c.value}`).join(' '))
})
session.setCells([{ sheet: 'Plan', cell: 'A1', input: { value: 150 } }])
session.undo()
session.previewCellText('Plan', 'A2', '300')
unsubscribe()
session.setCells([{ sheet: 'Plan', cell: 'A2', input: { value: 300 } }])
console.log(session.revision)
// → setCells 1 A1=150 A3=400
// → undo 2 A1=100 A3=350
// → 3
```

- A preview, a read, a refusal and a call that throws are not told; the listener runs when the call is done, before it returns, so the session it reads is the one after the change.
- Listeners run in the order they subscribed, and changes reach them in the order of their revisions: a change a listener makes reaches every listener after the one it answers.
- A listener that throws stops neither the call nor the other listeners; its error is thrown again in a microtask.
- `subscribe` answers the function that unsubscribes; `close()` unsubscribes every listener. An asynchronous session (`createInlineCalc`, `createCalcPool`) takes `subscribe` the same way, synchronously: a listener runs when a call's answer has come, before the call's promise settles.

### The clock and volatile functions

```js
import { compute } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { formula: 'TODAY()' }, A2: { formula: 'NOW()' }, A3: { formula: 'RANDBETWEEN(1,6)' } } }] }

const fixed = compute(book, { clock: { now: '2026-01-31T18:00:00Z', seed: 42 } })
console.log(fixed.sheets[0].cells.A1.value, fixed.sheets[0].cells.A2.value) // → 46053 46053.75

const now = compute(book) // the current moment and a random seed
console.log(compute(book, { clock: now.clock }).hash === now.hash) // → true
console.log(now.volatile.map((c) => c.cell).join(',')) // → A1,A2,A3
```

`volatile` lists the cells that depend on volatile functions, directly or through other cells, and on the clock's year through date text without a year ([Localization](#localization)). A session opened without a clock keeps the moment it was opened until `setClock`; `createSession(book, { live: true })` takes a fresh clock on every recomputation instead.

### Arrays

A range or an array constant where a function takes one value, or under an operator, is computed element by element, and functions that take ranges take arrays the same way. A formula whose result is an array of several values spills: its cell shows the first value and the cells to its right and below the others, as in Excel 365. In the result a spilled cell has `spillOf`, the formula's address, and the formula has `spill`, the range it fills; `E2#` refers to that range. Where a cell of the range holds a value or a formula, or a merged area meets it, the formula gives `#SPILL!` and its issue names the cell.

```js
import { compute } from 'xlcalc'

const cells = {
  A1: { value: 3 }, A2: { value: 5 }, A3: { value: 8 },
  B1: { formula: 'SUM(A1:A3*2)' },
  B2: { formula: 'SUMPRODUCT((A1:A3>4)*A1:A3)' },
  C1: { formula: 'SORT(A1:A3,1,-1)' },
  D1: { formula: 'SUM(C1#)' },
  E1: { formula: 'A1:A3*2' },
  E3: { value: 'x' },
}
const r = compute({ sheets: [{ name: 'S', cells }] })
const out = r.sheets[0].cells
console.log(out.B1.value, out.B2.value, out.D1.value) // → 32 13 16
console.log(out.C1, out.C2) // → { value: 8, text: '8', formula: 'SORT(A1:A3,1,-1)', spill: 'C1:C3' } { value: 5, text: '5', spillOf: 'C1' }
console.log(out.E1.value, r.diagnostics[0].reason)
// → { error: 'SPILL' } the array of 3×1 values cannot spill: 'S'!E3 is not empty
```

A spilled cell is a value of the book: it is in `sheets[].cells` and in the hash. An array of more than 65,536 values, such as `A:A*2` (1,048,576 rows), lists its cells up to where the rest holds one value and gives that value as the formula's `spillFill` (0 below the data). The design is in `docs/arrays-incell.md`, spilling and the rules measured in Excel in `docs/spill.md`.

`@A1:A10` takes the cell of the range in the formula's row or column, and an array's top-left value. `SUM((A1:A3,C1:C3))` sums a union of ranges, and `SUM('Jan:Mar'!B2)` sums B2 on every sheet from Jan to Mar. A space between two ranges is their intersection (`SUM(A1:C3 B2:D4)` sums B2:C3, `#NULL!` where they do not meet), and `:` between references that are not one range joins them (`SUM(A1:INDEX(B1:B9,3))`). [docs/references.md](docs/references.md) has the rules.

An array formula of the old kind, entered with Ctrl+Shift+Enter, is a cell with `array`, its area:

```js
import { compute } from 'xlcalc'

const cells = { A1: { value: 1 }, A2: { value: 2 }, C1: { formula: 'A1:A2*10', array: 'C1:C3' } }
const out = compute({ sheets: [{ name: 'S', cells }] }).sheets[0].cells
console.log(out.C1.value, out.C2.value, out.C3.value) // → 10 20 { error: 'NA' }
```

The array fills its area and never spills. A session refuses to change part of the area, as Excel does. `fromLegacy` from `xlcalc/xlsx` turns a formula of the old kind from a file into the dynamic formula Excel 365 shows for it, with `@` where Excel puts it, and `readXlsx` does so for every such formula of a file. [docs/legacy.md](docs/legacy.md) has both.

### Checking formulas

`validateFormula` checks a formula without a book: it returns the canonical text (sheet names quoted, function names upper-case) and what the formula uses, or a check code with a message in English and a position.

```js
import { validateFormula, validateName, validateNames } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

console.log(validateFormula("sum(Sales!b1:b3, 'Q1 data'!A1)", { sheets: ['Sales', 'Q1 data'] }).canonical)
// → SUM('Sales'!B1:B3,'Q1 data'!A1)
const bad = validateFormula('SUM(B1:B3', { sheets: ['Sales'] })
console.log(bad.ok, bad.code, bad.message, bad.at) // → false syntax expected ), found end of formula 9
console.log(validateName('Q1').ok, validateName('tax_rate').ok, validateName('Ставка').ok) // → false true true
console.log(validateName('ложь').ok, validateName('ложь', { locale: ruRU }).code) // → true boolean
const names = [{ name: 'R', ref: '1' }, { name: 'rate', ref: "'Q1 data'!B1" }, { name: 'gone', ref: "'Other'!A1" }]
console.log(validateNames(names, { sheets: ['Sales', 'Q1 data'] }).map((x) => `${x.name} ${x.dropped}`)) // → [ 'R true', 'gone false' ]
```

`validateName` checks a name by Excel's rules for a book of `locale`, a loaded locale's tag or a locale object (en-US without one; a ru-RU book also refuses `ИСТИНА` and `ЛОЖЬ`), and refuses with the rule the name breaks: `chars` (a character or a length Excel does not take), `reference` (it reads as a cell, `Q1`, or R1C1, `R2C3`) or `boolean`. `validateNames` gives the name issues a book with these sheets, tables and locale would give, without the book: `dropped` is true for a name the book leaves out (one Excel would not take, a duplicate, an unknown scope sheet, a formula that does not parse or a reference without its sheet), false for one it keeps that gives an error where it is used, so a host can refuse a name before it stores it.

Every check answers `{ ok: true, … }` or a refusal, the type `Refusal`: `{ ok: false, code, message, at? }`, a code a program reads, a message for people and, for text, the position in it. So do `validateFormula`, `validateLocalFormula`, `toLocal`, `toInvariant`, `parseEntry`, `validateName` and `editStructure`.

Check codes: `empty`, `syntax`, `too_long` (over 8,192 characters), `too_deep` (over 64 levels of nesting), `external_ref` (a reference to a workbook `ctx.externalBooks` holds no link to, [Links to other workbooks](#links-to-other-workbooks)), `unknown_function`, `unknown_sheet`, `unknown_name`, `unsupported` (an external function called with non-literal arguments). In `compute`, an invalid formula gives `#NAME?` for an unknown function or name, `#REF!` for an unknown sheet or an external reference, and `#ERROR!` otherwise. A formula whose result is an array is valid: it spills.

### Typed entry

`parseEntry` reads text typed into a cell the way Excel reads it in the book's locale, without a book, and gives the cell's new input: a number, a date or a time with the format Excel sets for it (`15.03.2026` in ru-RU gets the short date, `5%` the format `0%`, `$1,234.50` the locale's money), a boolean, an error, text, or a formula in the stored form. The cell typed into keeps its style, meta, note and link, and its format unless Excel changes it: a number typed into a percent cell is percent, a time typed into a date cell sets `h:mm`, the number formats and the book's own formats stay.

```js
import { parseEntry } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

const ru = (text, cell) => parseEntry(text, { locale: ruRU, cell, now: '2026-10-10T12:00:00Z' })
console.log(ru('1 234,5').input, ru('15.03.2026').input, ru('5%').input) // → { value: 1234.5, numberFormat: 4 } { value: 46096, numberFormat: 14 } { value: 0.05, numberFormat: 9 }
console.log(ru('ИСТИНА').input, ru('#Н/Д').input, ru("'007").input, ru('15.3').input) // → { value: true } { formula: '#N/A' } { value: '007' } { value: 46096, numberFormat: 16 }
console.log(ru('=СУММ(A1;1,5)').input, ru('=FOO(1)').input, ru('-C1').input) // → { formula: 'SUM(A1,1.5)' } { formula: 'FOO(1)' } { formula: '-C1' }
const stop = ru('=SUM(A1,1.5)')
console.log(stop.ok, stop.code, stop.message, stop.at) // → false syntax Здесь не может стоять «.» 9
console.log(ru('18:30', { numberFormat: 14, style: 2, meta: { id: 7 }, note: { text: 'checked' } }).input)
// → {
//     value: 0.7708333333333334,
//     numberFormat: 20,
//     style: 2,
//     meta: { id: 7 },
//     note: { text: 'checked' }
//   }
console.log(parseEntry('3/15/2026').input, parseEntry('$1,234.50').input.value, parseEntry('5', { cell: { numberFormat: '0%' } }).input) // → { value: 46096, numberFormat: 14 } 1234.5 { value: 0.05, numberFormat: '0%' }
```

`opts.locale` is the book's locale, a loaded locale's tag or a locale object ([Localization](#localization)), `opts.cell` the entry context (its effective `numberFormat`, `style`, `meta`, `note` and `link`), `opts.now` the moment whose year completes a date typed without one (`15.3`), and `opts.dateSystem` the book's date system. A `CellInput` or a `CellView` from `getCell` or `getCells` holds only the cell's own number format. To account for a row's or column's format, use the resolved cell from `getViewport`, or supply the effective `numberFormat` yourself:

```js
import { createSession, parseEntry } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', columns: { A: { numberFormat: '0%' } }, cells: { A1: { value: 0.1 } } }] })
const shown = session.getViewport('S', 'A1').cells.A1
console.log(parseEntry('5', { cell: shown }).input) // → { value: 0.05, numberFormat: '0%' }
session.close()
```

`parseEntry` stores the supplied format in its returned input. For editing a session, `setCellText` resolves the inherited format itself and keeps it on its row or column instead of storing it on the cell. Only a formula that does not parse stops the entry, with the check code, its position in the text and a message in the book's locale, as Excel's window does; an unknown function is taken and computes `#NAME?`. Empty text clears the value and keeps the rest of the cell (`null` when nothing is left). The rules, measured case by case in Excel for Mac and Excel for Windows, typed through the object model and on the keyboard, are in [docs/entry.md](docs/entry.md). A number typed into a column too narrow for it widens the column: `entryWidth` gives the width ([Showing values](#showing-values)).

A session takes typed text into a cell itself: `setCellText(sheet, cell, text, opts?)` reads it as `parseEntry` does, in the book's locale and date system, by the session's clock and the format the cell shows (its own, else its row's or column's, which the cell does not take as its own), keeps the rest of the cell, and widens the column in the same step of the undo as Excel does, unless `widen: false`; `previewCellText` says the same and changes nothing.

```js
import { createSession, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const session = createSession({
  locale: 'ru-RU',
  styles: [{ font: { bold: true } }],
  sheets: [{
    name: 'S',
    cells: { A1: { value: 1, style: 0 } },
    columns: { B: { numberFormat: '0%' } },
    validations: [{ range: 'C1:C9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'warning' }],
  }],
}, { clock: { now: '2026-10-10T12:00:00Z', seed: 1 } })
const typed = session.setCellText('S', 'A1', '1234567890')
console.log(typed.width, session.getBook().sheets[0].cells.A1) // → 11.1640625 { style: 0, value: 1234567890 }
console.log(session.setCellText('S', 'B1', '7').changes, session.getBook().sheets[0].cells.B1) // → [ { sheet: 'S', cell: 'B1', value: 0.07, text: '7%' } ] { value: 0.07 }
const refused = session.setCellText('S', 'A2', '=СУММ(A1;1.5)')
console.log(session.previewCellText('S', 'C1', '-2').validation.severity, refused.code, refused.at) // → warning syntax 10
```

A value the cell's rule of data validation does not take is written, with `validation`, the rule, its alert's style (`severity`) and message. A formula that does not read, a change of part of an old-kind array and a formula in a table's header cell are refused (`{ ok: false, code, message, at? }`, the message in the book's locale); text typed into a header cell renames its column, as `setCells` does.

Excel shows the rule's alert before it writes such a value: a `stop` alert never takes it (Retry or Cancel), a `warning` writes it on Yes, an `information` on OK. With `refuseInvalid: true` the call writes nothing and answers `{ ok: false, code: 'validation', sheet, cell, rule, severity, message }`, so the host shows the alert first. Where the person keeps the value, the host makes the same call again without the option, and the value is written with its alert. Without the option the value is written, and a host that refuses it afterwards calls `undo()`, which leaves a step to redo.

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{ name: 'S', cells: {}, validations: [{ range: 'A1:A9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'warning', error: 'Positive only' }] }],
})
const asked = session.setCellText('S', 'A1', '-2', { refuseInvalid: true })
console.log(asked.code, asked.cell, asked.severity, asked.message, session.revision, session.history) // → validation A1 warning Positive only 0 { undo: 0, redo: 0 }
// the person answers Yes: the same call without the option writes the value
const kept = session.setCellText('S', 'A1', '-2')
console.log(kept.validation.severity, session.getBook().sheets[0].cells.A1, session.revision) // → warning { value: -2 } 1
const many = session.setCellTexts([{ sheet: 'S', cell: 'A2', text: '5' }, { sheet: 'S', cell: 'A3', text: '1.5' }], { refuseInvalid: true })
console.log(many.code, many.cell, session.revision) // → validation A3 1
```

`setCellTexts(entries, opts?)` takes several texts typed at once: each `{ sheet, cell, text }` is read as `setCellText` reads it in its cell, in order, and all are written with one recomputation, one revision and one step of the undo. A cell given twice keeps the later text, read by the format the earlier one gave it. A column widens once, to the largest width a text typed alone into one of its cells would give it (`widths`, unless `widen: false`), and the cells whose values their rules do not take come back with their alerts (`validations`). A text `setCellText` refuses refuses the whole call, with its `sheet` and `cell`, and nothing is written; with `refuseInvalid`, so does the cell whose value breaks the strictest rule once all the texts are written (stop before warning before information, the first cell of that style). `previewCellTexts` answers each cell's input and value, the widths and the alerts, and changes nothing.

```js
import { createSession, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const session = createSession({
  locale: 'ru-RU',
  sheets: [{ name: 'S', cells: { A1: { value: 'Item' } }, validations: [{ range: 'C1:C9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'warning' }] }],
}, { clock: { now: '2026-10-10T12:00:00Z', seed: 1 } })
const typed = session.setCellTexts([
  { sheet: 'S', cell: 'B1', text: '15.03.2026' },
  { sheet: 'S', cell: 'B2', text: '1234567890' },
  { sheet: 'S', cell: 'C1', text: '-2' },
])
console.log(typed.widths, typed.validations.map((v) => `${v.cell} ${v.severity}`)) // → [ { sheet: 'S', column: 'B', width: 11.1640625 } ] [ 'C1 warning' ]
const refused = session.setCellTexts([{ sheet: 'S', cell: 'A2', text: '7' }, { sheet: 'S', cell: 'A3', text: '=СУММ(A1;1.5)' }])
console.log(refused.code, refused.cell, refused.at, session.revision) // → syntax A3 10 1
console.log(session.previewCellTexts([{ sheet: 'S', cell: 'A2', text: '5%' }, { sheet: 'S', cell: 'A3', text: '=A2*2' }]).cells.map((c) => [c.cell, c.input, c.value]))
// → [
//     [ 'A2', { value: 0.05, numberFormat: 9 }, 0.05 ],
//     [ 'A3', { formula: 'A2*2' }, 0.1 ]
//   ]
```

A cell edited in place, after F2, a double click or in the formula bar, opens with its own text, and Excel reads what comes back otherwise than text typed over the cell: its percent editor adds `%` only to a number begun by a sign. With `inPlace: true`, `setCellText`, `setCellTexts`, `fillCellText` and their previews read the text so: in a percent cell `6%` is 0.06 and `6` is 6 (600%), where typed over the cell `6%` is the text `6%%` and `6` is 6%.

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 0.05, numberFormat: '0%' } } }] })
console.log(session.previewCellText('S', 'A1', '6%').input, session.previewCellText('S', 'A1', '6%', { inPlace: true }).input)
// → { value: '6%%', numberFormat: '0%' } { value: 0.06, numberFormat: '0%' }
```

A paste and Control+Return read text otherwise than typing, and have their own calls. `pasteText(sheet, cell, text, opts?)` pastes text from the clipboard as Excel pastes it: `parseClipboard(text)` makes it a rectangle, rows at line breaks and fields at tabs, every row as wide as the longest, and the rectangle fills the cells from `cell`, an empty field, a short row's missing ones and the row after a last line break clearing theirs. A field begun by a quote runs to its closing quote, line breaks and tabs inside it, and wraps its cell. Each field is read as typed into its cell, but an apostrophe stays part of the text, a formula Excel does not take is pasted as text, a percent cell has no percent editor (`5` is 500%) and `€5` takes the euro format; no column widens and no rule is checked. The answer has the `range` filled; part of an array, a formula into a header cell and a rectangle past the sheet's edge are refused with the `sheet` and `cell`.

`fillCellText(sheet, range, text, opts?)` enters a text into every cell of a range, as Control+Return: it reads the text once in the active cell, the range's top-left one or `active`, and every cell takes what the active cell holds after it, a formula moved as a copy from the active cell moves it (`copyFormula`) and the active cell's number format where it has one. The columns widen as typing widens them (`widths`, unless `widen: false`), and only the active cell's rule is checked (`validation`, or with `refuseInvalid` the refusal before anything is written). Both are one step of the undo, and `previewPasteText` and `previewFillCellText` change nothing. The rules, as measured in Excel, are in [docs/entry.md](docs/entry.md#a-paste-and-controlreturn).

```js
import { createSession, parseClipboard, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const session = createSession({
  locale: 'ru-RU',
  sheets: [{
    name: 'S',
    cells: { A1: { value: 2 }, A2: { value: 3 } },
    columns: { C: { numberFormat: '0%' } },
    validations: [{ range: 'D1:D9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'warning' }],
  }],
}, { clock: { now: '2026-10-10T12:00:00Z', seed: 1 } })
const text = "'007\t=СУММ(A1;1.5)\t5\t-2\r\n€5"
console.log(parseClipboard(text)) // → [ [ "'007", '=СУММ(A1;1.5)', '5', '-2' ], [ '€5', '', '', '' ] ]
const pasted = session.pasteText('S', 'A4', text)
console.log(pasted.range, Object.entries(session.getCells('S', pasted.range)).map(([a, c]) => `${a} ${c.text}`), 'validation' in pasted)
// → A4:D5 [ "A4 '007", 'B4 =СУММ(A1;1.5)', 'C4 500%', 'D4 -2', 'A5 € 5' ] false
const filled = session.fillCellText('S', 'D1:D2', '=A1-5')
console.log(filled.changes.map((c) => `${c.cell} ${c.formula} ${c.value}`), filled.validation.severity) // → [ 'D2 A2-5 -2', 'D1 A1-5 -3' ] warning
const refused = session.fillCellText('S', 'E1:E3', '=СУММ(A1;', { active: 'E2' })
console.log(refused.code, refused.cell, refused.at, session.revision) // → syntax E2 9 2
```

### Addresses and helpers

A grid, a dialog or a server that keeps versions of a book needs Excel's rules for addresses, names and values without computing anything. The package gives them by the engine's own code: `xlcalc/a1` holds the addresses and names in a few kilobytes and none of the engine, `xlcalc/text` the references of formula text and the names a reference writes, for a browser that colours a formula as it is typed without the engine, and `xlcalc` exports the same functions with the helpers that need a locale. These take a loaded locale's tag or a locale object, which needs no loading; `quoteSheetIfNeeded` takes the object alone, so `xlcalc/a1` carries no list of locales.

```js
import { cellAddress, linesKey, parseLines, parseRange, quoteSheetIfNeeded, rangeAddress, sameName } from 'xlcalc/a1'
import { entryText, errorText, needsApostrophe, sameValue, validateSheetName } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'
import { refAt, scanReferences } from 'xlcalc/text'

console.log(parseRange('$d$9:B2'), parseRange('ZZZ1'), rangeAddress({ top: 1, left: 2, bottom: 1048576, right: 4 }))
// → { kind: 'cells', top: 2, left: 2, bottom: 9, right: 4 } null B:D
console.log(cellAddress(12, 2), parseLines('B:D', 'columns'), linesKey(9, 5, 'rows')) // → B12 { first: 2, last: 4 } 5:9
console.log(quoteSheetIfNeeded('Sales'), quoteSheetIfNeeded('Q1 data'), quoteSheetIfNeeded('R2D2'), quoteSheetIfNeeded('ИСТИНА', ruRU), sameName('straße', 'STRASSE'))
// → Sales 'Q1 data' 'R2D2' 'ИСТИНА' true
const bad = validateSheetName('Итоги/2025', { locale: ruRU })
console.log(bad.code, bad.at, bad.message) // → character 5 Имя листа не может содержать «/» (Excel запрещает : \ / ? * [ ])
console.log(errorText('DIV0', ruRU), entryText(0.125, { numberFormat: '0.0%' }), entryText(45688.75, { numberFormat: 14, locale: ruRU }), entryText('007'), needsApostrophe('1.5', { locale: ruRU }))
// → #ДЕЛ/0! 12.5% 31.01.2025  18:00:00 '007 true
console.log(scanReferences('=СУММ(B2:B9;\'Q1 data\'!C1;"D4")', { locale: ruRU }).map((r) => `${r.at}-${r.end} ${r.text} ${r.sheet ?? ''}`))
// → [ '6-11 B2:B9 ', "12-24 'Q1 data'!C1 Q1 data" ]
console.log(scanReferences('=A1+)B2*"C3').map((r) => r.text)) // → [ 'A1', 'B2' ]
console.log(refAt('=SUM(A1,').action, refAt('=A1').action, refAt('="a').action, refAt('=SUM(A1;', undefined, { locale: ruRU })) // → insert commit none { action: 'insert', at: 8 }
console.log(sameValue(0.1 + 0.2, 0.3), sameValue(1, true)) // → true false
```

Rows and columns are numbered from 1, as A1 numbers them. `parseCell` and `parseRange` take `$` and letters in either case and give null past `XFD1048576` or for anything else; a range given corner to corner the other way round is put in order, as Excel orders a typed reference, and its `kind` says whether it named cells, whole columns (`B:D`) or whole rows (`2:9`). `cellAddress`, `colLetters`, `colNumber` and `rangeAddress` write them back, `rangeAddress` as Excel writes a range: `B2` for one cell, `B:D` for a range over every row. `parseLines` and `linesKey` read and write the keys of `SheetInput.columns` and `rows` (`C`, `B:D`, `5`, `5:9`), null for a key a book does not take, and `patchLines` patches such a map as `setLayout` does ([Layout](#layout)); `rectsMeet` and `rectContains` compare rectangles; `MAX_ROWS` and `MAX_COLS` are the sheet's size.

`refAt(text, caret?, { locale })` says what an arrow does in formula text typed over a cell, as Excel's Point mode decides it: `insert`, a reference to the cell the arrow points at goes in at `at` (after an `=` of its own with `equals`, at the start of the text); `commit`, the text is entered and the selection moves, as after an operand; `none`, nothing happens (inside a string, after a character the locale's formulas do not have). Moving or extending a reference already pointed is the host's ([Point mode](docs/entry.md#point-mode)).

`nameKey` and `sameName` are the rule by which the engine compares names of sheets, tables and defined names: in any case, by Unicode's upper case in every locale (`straße` is `STRASSE`), but the dotless `ı` and the long `ſ`, which Excel for Windows keeps apart from `I` and `S`. `quoteSheetIfNeeded` writes a sheet's name in a reference as Excel writes it, measured in Excel for Mac and for Windows: in quotes (an apostrophe inside doubled) when the name is anything but letters, digits, `_` and `.` after a letter or `_`, when it is a cell of the sheet (`A1`; not `XFE1`), when it begins with an R1C1 reference with a number (`R2D2`, `C3PO`; not `Region`, nor `C1.x`, where a `.` follows), or when it is `R`, `C`, `RC` or one of the locale's booleans (`TRUE`, `'ИСТИНА'` in ru-RU's formula text). Stored formulas keep every sheet in quotes, as their canonical text; the text shown to a person (`toLocal`, a window's `formulaLocal` and `editText`, the text of `explain`) quotes a sheet by this rule. `validateSheetName` checks a sheet's name as a book's sheets are checked (31 characters at most, none of `: \ / ? * [ ]`, no apostrophe at either end, not `History`): `{ ok: true }`, or a refusal with the code `empty`, `too_long`, `character`, `apostrophe` or `reserved`, a message in `opts.locale`'s language and, for a character, its place.

`errorText` writes an error as a cell shows it in a locale; any other value without a book is `displayValue(value, { locale })`. `entryText` is the text a cell opens with for editing, as Excel's formula bar shows its value: a number whole (15 digits, positional up to 21 characters), a percent cell's number as a percent, a date or a time by the locale's short date and long time, booleans and errors in the locale's words, and text with an apostrophe where typing it would read as something else, which `needsApostrophe` tells by `parseEntry`. `sameValue` compares two values as the values hash does (numbers to 15 significant digits, `-0` as 0, errors by code), and `canonValue` is a value's text in the hash. `scanReferences` finds the references in formula text as it is typed, stored or in a locale, with their places for colouring them: cells, whole columns and rows, sheets, references across sheets and spills, and structured references with their table, columns and specifiers. It reads the text with the engine's tokenizer: a token that does not read (a character a formula does not take, an error literal half typed, a sheet with nothing after it) is skipped and the text read on past it, and an open string, quoted sheet name or bracket takes the rest of the text, so a formula half typed or broken in the middle gives the references around the place (the package's own rule, not one measured in Excel). How the rules were measured is in [docs/references.md](docs/references.md#sheet-names-in-references) and [docs/entry.md](docs/entry.md#the-text-a-cell-opens-with).

### Errors

Cell errors come out as `{ error: code }` with codes `DIV0`, `NA`, `NAME`, `NUM`, `REF`, `VALUE`, `CYCLE`, `SPILL`, `CALC` and `ERROR`. Every cell with an error is listed in the result's `diagnostics` with a reason (`kind: 'cell'`), which names the cell an error came from (`division by zero, from 'S'!A2`); a session's `getCellIssue` gives one cell's, with the cell the error began in ([A window for a grid](#a-window-for-a-grid)). A circular reference stops the cells in the cycle and everything that reads them: they hold their values ([Circular references](#circular-references)), each listed with `error: 'CYCLE'`; the rest of the book is computed. A book with [iterative calculation](#iterative-calculation) computes its cycles instead. `#CYCLE!` itself is a name that uses itself, through other names or not.

Input of the wrong shape is a programming error and throws `InputError` (a `TypeError` with `path`), for example `book.sheets[0].cells["b2"]: invalid cell address "b2": expected A1 style such as "B12" (column letters must be upper-case: "B2")`. This covers wrong types, malformed or out-of-grid addresses, duplicate sheet names (compared case-insensitively, as in Excel), unknown fields, bad dates and clocks, and unknown sheets in session edits. A known field given as undefined is no field: `{ border: { top: undefined } }` is a border without a top, in a cell's style as in a conditional format's. A call never ignores what it does not take: an options object or a callback past its arguments, or an option it does not know, is an `InputError` that names the call taking that option, `getCells: takes 2 arguments, got an object as argument 3 (formulaLocal is an option of getViewport)`; an asynchronous call rejects with it. An array, a number or undefined past the arguments passes, so `['A1', 'B2'].map(parseCell)` works, as `map` gives a callback its index and array too.

### Named ranges

```js
import { compute } from 'xlcalc'

const result = compute({
  sheets: [
    { name: 'Data', cells: { A1: { value: 'tea' }, B1: { value: 3 }, A2: { value: 'coffee' }, B2: { value: 5 } } },
    { name: 'Report', cells: { A1: { formula: 'SUMIF(items,"tea",amounts)' } } },
  ],
  names: [
    { name: 'items', ref: "'Data'!$A:$A" },
    { name: 'amounts', ref: "'Data'!$B:$B" },
  ],
})
console.log(result.sheets[1].cells.A1.value) // → 3
```

A name follows Excel's rules for names, in letters of any script (`Ставка`, `НДС_2025`, `Став?ка`), and is compared in any case: `ставка` is the name `Ставка`, and a formula's text shows it as it is defined (`ставка*2` is `Ставка*2`). A name with `scope: 'Sheet'` is visible on that sheet only and shadows a global name. Invalid names are reported in the result's `diagnostics` (`kind: 'name'`), `dropped` telling a name left out from one kept with an error.

A name can hold any formula: a constant (`0.2`), a calculation (`'Data'!$B$1*2`), a reference from `INDEX`, a union. A reference without `$` moves with the cell that uses the name. [docs/names.md](docs/names.md) has the rules.

### Tables

A sheet's `tables` are Excel's tables (list objects): a header row of column names, data rows and a totals row. Formulas refer to their parts by name, and the references follow a table as rows and columns are inserted and deleted in it.

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

A formula stores its references as an xlsx file does (`Orders[[#This Row],[Qty]]` for `[@Qty]`), and `toLocal` shows them as Excel does, in Russian too (`#Данные`). Structure operations grow, shrink and move tables as Excel does, an edit of a header cell renames its column, and `editStructure` adds, renames, deletes and converts tables. [docs/tables.md](docs/tables.md) has the rules.

A formula typed in a cell of a table refers to it without its name, `[@Qty]`. Without the book, the table of the cell is an option: `toLocal(formula, locale, { table })` drops the name as Excel shows such a formula, `toInvariant(text, locale, { table })` puts it back, and `validateFormula` and `validateLocalFormula` with `ctx.table` check the formula as in the table and give the text the book stores. `escapeColumn` writes a column's name in a reference, `'` before the characters that mean something there (`[`, `]`, `#`, `'`, a leading `@`).

```js
import { escapeColumn, toInvariant, validateFormula } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

console.log(toInvariant('=[@Qty]*1,5', ruRU, { table: 'Orders' }).formula) // → Orders[@Qty]*1.5
console.log(validateFormula('[@Qty]*[@Price]', { table: 'Orders' }).canonical) // → Orders[[#This Row],[Qty]]*Orders[[#This Row],[Price]]
console.log(`Orders[${escapeColumn('Price #2')}]`) // → Orders[Price '#2]
```

### External functions

Host functions that are not part of the book (a balance from a database, say) can appear in formulas. `validateFormula(text, { externalFunctions: ['BALANCE'] })` returns the canonical text of each call in `externals`; the host resolves the values in advance and passes them to `compute`:

```js
import { compute } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { formula: 'BALANCE("2026-01")*2' } } }] }
const result = compute(book, { externals: { functions: ['BALANCE'], values: { 'BALANCE("2026-01")': 21 } } })
console.log(result.sheets[0].cells.A1.value) // → 42
```

Arguments of an external function must be literals; a call without a value gives `#N/A`.

### Localization

Books store formulas the way xlsx files do: English function names, `,` between arguments, `.` as the decimal separator. A locale changes how formulas are typed and shown, as Excel shows them in another language, and how a book computes, as Excel computes with those regional settings. en-US is built in. ru-RU (Excel with the Russian interface) is a module of its own, which a bundle carries only when it is imported: `registerLocale` loads it, once, for books in ru-RU and for its tag.

```js
import { LOCALES, registerLocale, toInvariant, toLocal, validateLocalFormula } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

console.log(toLocal('IF(B1>0.5,SUM({1,2;3,4}),#N/A)', ruRU).formula) // → ЕСЛИ(B1>0,5;СУММ({1;2:3;4});#Н/Д)
registerLocale(ruRU)
console.log(LOCALES.get('ru-RU') === ruRU, [...LOCALES.keys()]) // → true [ 'en-US', 'ru-RU' ]
console.log(toInvariant('=СУММ(A1:A3;1,5)', 'ru-RU').formula) // → SUM(A1:A3,1.5)
const bad = validateLocalFormula('=СУММ(A1;1.5)', 'ru-RU')
console.log(bad.ok, bad.code, bad.message, bad.at) // → false syntax Здесь не может стоять «.» 10
console.log(validateLocalFormula('=ОКРУГЛ(Данные!B2*1,2;0)', ruRU, { sheets: ['Данные'] }).canonical) // → ROUND('Данные'!B2*1.2,0)
```

`LOCALES` lists the loaded locales by tag. Every function that takes a locale takes the tag of a loaded one or the locale object, and an object needs no loading where no book is needed: `toLocal`, `toInvariant`, `validateLocalFormula`, `parseEntry`, `displayValue`, `entryWidth`, `builtinFormatCode`, `numberFormatCategory`, `numberFormatChoices`, `entryText`, `needsApostrophe`, `errorText`, `scanReferences`, `validateSheetName`, `validateName` and `validateNames`. A book's `locale` is a tag, and a book in a locale that is not loaded is an `InputError` that says what to import, wherever the book goes (`compute`, sessions, `editStructure`, `fitColumns`, `writeXlsx`, a pool); `readXlsx` keeps a file's locale as it reads it. A pool of threads loads in its threads the locales loaded where it runs, and those loaded later too. A locale is a plain object (the `Locale` type), so formula text in another language is another object of the same shape; one without rules of computation reads and shows values as en-US does. Adding a locale to the package is a module ([docs/locale.md](docs/locale.md#adding-a-locale)).

```js
import { compute } from 'xlcalc'

try {
  compute({ locale: 'ru-RU', sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] })
} catch (e) {
  console.log(e.message) // → book.locale: the locale ru-RU is not loaded: import { ruRU } from 'xlcalc/locales/ru-RU' and call registerLocale(ruRU) (loaded: en-US)
}
```

`toLocal` and `toInvariant` return `{ ok: true, formula }` or `{ ok: false, code, message, at }` with the position in the given text. They translate function names, TRUE and FALSE, error literals, the decimal separator and the separators of arguments and of array constants (`{1;2:3;4}` in ru-RU); strings, defined names, references and whitespace stay as they are. `toLocal` writes a sheet's name in quotes only where Excel puts them, by `quoteSheetIfNeeded` with the locale's booleans ([Addresses and helpers](#addresses-and-helpers)), as every text shown to a person does (a window's `formulaLocal` and `editText`, the text of `explain`): the stored `'Data'!B2*2` is `Data!B2*2` in every locale, en-US's too, `'Q1 data'!A1` keeps its quotes, and a sheet named `ИСТИНА` is `'ИСТИНА'!A1` in ru-RU. `toInvariant` reads a name with quotes and without and keeps it as typed; a stored formula quotes every sheet, as its canonical text. `toInvariant` drops a leading `=` and also reads English function names, TRUE/FALSE and error texts; a function the locale does not name (an external one, say) passes through. Text that would read differently once translated (`1.5` typed in ru-RU, where `.` is not a decimal separator) is a `syntax` error at its place. `validateLocalFormula(text, locale, ctx?)` is `validateFormula` for typed text: the same result with the canonical stored text, and positions in `text`. The messages of all three are in the locale's language.

The ru-RU data come from Excel itself. `bench/locale/snapshot.ts` opens books with 6,381 formulas (the coverage measurement's cases, the probes, a call of every function on Microsoft's list that the measurement lacks, literals) in Excel 16.113 for Mac with the Russian interface and reads each formula back in English and in Russian; `bench/locale/generate.ts ru-RU` writes the formula text of ru-RU's locale file, `src/locales/ru-RU.ts`, from that text (`test/excel/locale/ru-RU.json`). `test/locale-excel.test.ts` checks that both translations reproduce Excel's text of all 6,381 formulas, and `test/locale-roundtrip.test.ts` that the 15,266 formulas of the measurement, the case sets and the probes come back unchanged from ru-RU. Excel 16.113 for Mac does not know `ENCODEURL`, `FILTERXML`, `WEBSERVICE`, `REGEXEXTRACT`, `REGEXREPLACE` and `REGEXTEST`, and deletes `CALL` and `REGISTER.ID` from the books it loads, so these keep their English names; so do the error literals newer than `#GETTING_DATA` (`#SPILL!`, `#CALC!` and the rest), which it does not take as literals in a formula.

A book is computed in its own locale, `BookInput.locale` (the tag of a loaded locale: `'en-US'` by default, or `'ru-RU'` once loaded), as Excel computes with those regional settings, whoever opens the book; a viewer still sees the formulas in their own language through `toLocal`. The locale changes how text reads as a number, a date or a boolean, how numbers and booleans become text, `TEXT`'s format codes and what cells show, `DOLLAR` and `FIXED`, error texts, the words and numbers in criteria and the order of text:

```js
import { compute, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const cells = { A1: { formula: 'VALUE("1,5")+VALUE("1 234,5")' }, A2: { formula: '"a"&1.5&TRUE' }, A3: { formula: 'TEXT(45688.75,"ДД ММММ ГГГГ ч:мм")' }, A4: { formula: 'DOLLAR(-1234.567)' }, A5: { formula: 'VALUE("1.5")' } }
const r = compute({ locale: 'ru-RU', sheets: [{ name: 'S', cells }] }, { clock: { now: '2026-10-05T12:00:00Z', seed: 1 } })
const v = r.sheets[0].cells
console.log(v.A1.value, v.A2.value, v.A3.value, v.A4.value, v.A5.value) // → 1236 a1,5ИСТИНА 31 Январь 2025 18:00 -1 234,57 ₽ 46143
console.log(r.volatile.map((c) => c.cell).join(',')) // → A5
```

In ru-RU, as Excel 16.113 for Mac with the Russian interface and locale computes, and as Excel for Windows where the two part: numbers in text take a decimal comma and a space between groups (`1 234,5`) and the ruble after the amount (`1 234,5 ₽`); dates are day-month-year (`31.01.2025`, `31/1/25`, `2025-01-31`, `31 янв 2025` with a month named by three letters or more of its nominative name; English month names are not read); `ИСТИНА` and `ЛОЖЬ` are the booleans in text (`"TRUE"` is no boolean there); numbers become text with a decimal comma (`a1,5`), complex numbers too; `TEXT` reads the Russian codes (`ДД.ММ.ГГГГ`, `ч:мм:сс`, `# ##0,00`, `Основной`, `[Красный]`, `[Цвет10]`) and treats English letters as literals, while cells read their formats as xlsx files store them (`dd/mm/yyyy` shows `31.01.2025`); `DOLLAR` is `1 234,57 ₽`; month names have a capital (`Январь`) and short ones three letters (`фев`); text sorts Latin before Cyrillic, as Excel for Windows sorts (Excel for Mac puts Cyrillic first). Date text without a year (`1.5` is 1 May, `15 янв`) takes the year of the book's clock, as Excel takes the current year, so a cell that reads one depends on the clock as `TODAY` does: it is in `volatile` and a session recomputes it on `setClock`. The rules were measured case by case in Excel ([docs/locale.md](docs/locale.md)).

Messages for people come in the book's locale too: why a typed formula does not read (`parseEntry`, and `toLocal`, `toInvariant` and `validateLocalFormula` in their locale), why a structure operation is refused, and the text of `explain`, whose numbers, formulas, booleans and errors are written as the locale writes them, and its references as Excel writes them (`План!B1`). Codes, positions and the fields of `explain`'s tree stay as they are, so a program reads them the same in every locale.

```js
import { createSession, parseEntry, registerLocale, editStructure } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const book = { locale: 'ru-RU', sheets: [{ name: 'План', cells: { A1: { value: 400 }, B1: { formula: 'ROUND(A1*1.2,0)' } } }] }
console.log(createSession(book).explain('План', 'B1').text)
// → План!B1 = 480  ← ОКРУГЛ(A1*1,2;0)
//     План!A1 = 400
console.log(parseEntry('=ОКРУГЛ(A1;2;3)', { locale: 'ru-RU' }).message) // → Функция ОКРУГЛ принимает 2 аргумента, а указано 3
console.log(editStructure(book, { op: 'deleteSheet', sheet: 'План' }).message) // → В книге должен остаться хотя бы один лист
```

`validateFormula`, which has no locale, answers in English, as do the reasons of cell errors and the checks of names. A locale's messages are part of its module ([docs/locale.md](docs/locale.md#messages)), and a locale object without messages of its own speaks as the loaded locale of its tag, else in English.

### Showing values

A result gives each cell, beside its `value`, the `text` Excel shows in it (VBA's `Range.Text`) and the `color` of its format's section. The format is the cell's own `numberFormat`, else its row's, else its column's: an Excel format code as an xlsx file stores it (`#,##0.00`, `d mmmm yyyy`, `[Blue]0;[Red]-0`), shown in the book's locale. The width is the column's (`columns: { A: { width: 16 } }`, as xlsx stores it; Excel's default without one), or a merged area's columns: General drops decimals or turns scientific to fit, a number that still does not fit shows `#` across the column, text is never cut (it runs over the next cells). The value stays the number whatever the text shows, and the text is not part of the values hash.

```js
import { compute, displayValue, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const book = {
  sheets: [{
    name: 'S',
    cells: {
      A1: { value: 1234567.891, numberFormat: '#,##0.00' },
      A2: { value: -42, numberFormat: '[Blue]0;[Red]-0' },
      A3: { value: { date: '2025-01-31' }, numberFormat: 'd mmmm yyyy' },
      A4: { value: 1 / 3 },
      B1: { value: 123456789 },
    },
    columns: { A: { width: 16 }, B: { width: 6 } },
  }],
}
const cells = compute(book).sheets[0].cells
console.log(cells.A1, cells.A2.text, cells.A2.color, cells.A3.text, cells.A4.text, cells.B1.text)
// → { value: 1234567.891, text: '1,234,567.89', numberFormat: '#,##0.00' } -42 #FF0000 31 January 2025 0.333333333 1E+08
const ru = compute({ ...book, locale: 'ru-RU' }).sheets[0].cells
console.log(ru.A1.text, ru.A3.text, ru.A4.text) // → 1 234 567,89 31 Январь 2025 0,333333333
console.log(compute(book, { display: false }).sheets[0].cells.A1) // → { value: 1234567.891, numberFormat: '#,##0.00' }
console.log(displayValue(0.4567, { numberFormat: '0.0%', locale: 'ru-RU' }), displayValue(1234.5, { width: 3 })) // → { text: '45,7%' } { text: '##' }
console.log(displayValue(46096, { numberFormat: 'dd.mm.yyyy' }), displayValue(46096, { numberFormat: 'dd.mm.yyyy', width: 8.83203125 })) // → { text: '15.03.2026' } { text: '#########' }
```

`compute(book, { display: false })` and `createSession(book, { display: false })` give values alone and spend nothing on formatting. A session's changes carry the text of the cells they list, and with the display on they also list the cells whose text or colour changed without a new value: a new `numberFormat` through `setCells` or `patchCells`, a width, a row's or a column's format or a merged area ([Layout](#layout)), a column moved by `editStructure`. `displayValue(value, { numberFormat, locale, width })` shows a value without a book, for a cell being edited, say, in a loaded locale's tag or a locale object; the number itself is always the cell's `value`. Without `width` the width is not limited: General keeps its digits and no number turns into `#`, as no column narrows it. With the column's width it shows what the cell shows (Excel's default column is `width: 8.83203125`, as xlsx stores it).

A format can also be the number of one of Excel's built-in formats, from 0 to 49, as an xlsx file stores a cell's format without a code. Excel shows some of them by the locale: `numberFormat: 14` is the short date, `1/31/2025` in en-US and `31.01.2025` in ru-RU, as Excel for Windows shows it. `builtinFormatCode(id, locale?)` gives the code a number shows by (`locale` a loaded locale's tag or a locale object), for a format menu or an editor, and undefined for a number that is no built-in format, so `builtinFormatCode(id) !== undefined` tells whether it is one. The table is in [docs/formats.md](docs/formats.md).

```js
import { builtinFormatCode } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

console.log(builtinFormatCode(14), builtinFormatCode(14, ruRU), builtinFormatCode(7, ruRU)) // → m/d/yyyy dd.mm.yyyy #,##0.00 "₽";-#,##0.00 "₽"
console.log(builtinFormatCode(3), builtinFormatCode(50), builtinFormatCode(164) !== undefined) // → #,##0 undefined false
```

The Number Format box on Excel's Home tab names the category of the active cell's format, and its list applies a format of each category. `numberFormatCategory(format, locale?)` gives that category for a code or a built-in format's number, as a stable key and the name the box shows (`Number`, `Числовой`); `numberFormatChoices(locale?)` gives the list's items in order, each with its name, the format it applies and that format's code in the locale. Both come without the engine, in `xlcalc/text` too. The box names the category of Excel's Format Cells window whose formats hold the code as Excel keeps it, and Custom when none does: `0.000` is Number, `0.00;-0.00` and `#,##0%` are Custom, and so is `$#,##0.00`, as Excel keeps a dollar written bare as a literal; the Currency item writes `"$"#,##0.00`. In ru-RU the names are Russian and the rules Excel for Windows's: `0.00` is Числовой as in en-US, money is the ruble's (`#,##0.00\ "₽"` Денежный, while the dollar's `"$"#,##0.00` is `(все форматы)`), and the short date is `dd/mm/yyyy`, which Excel shows as `ДД.ММ.ГГГГ`. Excel for Mac names nearly every code `(все форматы)` there, as it compares the code with its formats written in Russian syntax. The rules and how they were measured are in [docs/formats.md](docs/formats.md#the-ribbons-number-format-box).

```js
import { numberFormatCategory, numberFormatChoices } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

console.log(numberFormatCategory('0.000'), numberFormatCategory('$#,##0.00').name, numberFormatCategory('"$"#,##0.00').name, numberFormatCategory(14).name)
// → { category: 'number', name: 'Number' } Custom Currency Date
console.log(numberFormatCategory('0.000', ruRU).name, numberFormatCategory('#,##0.00\\ "₽"', ruRU).name, numberFormatCategory('"$"#,##0.00', ruRU).name)
// → Числовой Денежный (все форматы)
console.log(numberFormatChoices().map((c) => c.name).join(', '))
// → General, Number, Currency, Accounting, Short Date, Long Date, Time, Percentage, Fraction, Scientific, Text
const { name, format } = numberFormatChoices(ruRU)[2]
console.log(name, format) // → Денежный #,##0.00\ "₽"
```

`TEXT` uses the same formats, read in the book's locale (`TEXT(x,"# ##0,00")` in ru-RU), without the width: it drops fills, has no colour and answers `#VALUE!` where a cell shows `#`. The rules, as measured in Excel 16.113 for Mac in en-US and ru-RU on 11,649 cells and 771 `TEXT` cases, are in [docs/formats.md](docs/formats.md), with the differences.

Excel widens columns by itself in two places, and the engine gives the same widths. AutoFit (a double click on a column's border) gives a column the width of its widest cell: `fitColumns(book, result, sheet, columns)` computes the widths for a computed book, and the layout edit `{ sheet, columns, fit: true }` applies them in a session ([Layout](#layout)). And a number typed into a column whose width nobody set widens it until the number shows whole: `entryWidth(value, opts)` gives that width for the value a typed text computes, for a host to set with `bestFit: true`. A width Excel fitted (`ColumnInput.bestFit`) widens again as numbers are typed; one set by hand stays.

```js
import { compute, entryWidth, fitColumns, parseEntry, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const book = {
  locale: 'ru-RU',
  sheets: [{ name: 'S', cells: { A1: { value: 'Итого' }, B1: { value: 1234567.5, numberFormat: '#,##0.00' }, C1: { value: 46096, numberFormat: 14 } } }],
}
const result = compute(book)
console.log(result.sheets[0].cells.C1.text, fitColumns(book, result, 'S', 'A:C'))
// → ######### { A: 5.6640625, B: 11.33203125, C: 10.1640625 }
const typed = parseEntry('1234567890', { locale: 'ru-RU' }).input
console.log(entryWidth(typed.value, { locale: 'ru-RU' }), entryWidth(3.14159265, { locale: 'ru-RU' }), entryWidth(typed.value, { locale: 'ru-RU', width: 9 }))
// → 11.1640625 undefined undefined
```

AutoFit counts every cell of the column with something to show, hidden rows too, but not cells merged across columns; a column with nothing to show keeps its width, and a hidden column is shown. Typing widens a column for a number, a date or a formula's number that AutoFit would give more room, but not for text, which runs over the next cells, nor for a number in General that shows with fewer decimals (`3,141593`). Excel widens nothing for text set through its object model, only for text typed. The widths are Excel for Mac's, measured on 994 columns in en-US and ru-RU ([docs/formats.md](docs/formats.md#fitting-columns)); Excel for Windows lays text out in pixels and fits it to other widths.

Rows fit the same way. `fitRows(book, result, sheet, rows)` gives the heights AutoFit gives rows `5` or `2:9` of a computed book, by spans of rows of one height, and `{ sheet, rows, fit: true }` applies them in a session: a row stands its cells on one line at its bottom, one line of each cell's font or the lines of text that wraps (`alignment.wrap`) at its column's width, and gets the most any cell needs above that line and the most any needs below, a line of Calibri 11 15 points alone and 16 wrapped; a cell merged down counts for nothing. A text with a line break typed (`setCellText`, `setCellTexts`, `fillCellText`) wraps its cell and fits its row in the same step, unless the row's height was set by hand; one pasted wraps its cell alone (a quoted field of the clipboard keeps its line breaks). A height Excel fitted is marked `bestFit` and written to xlsx as Excel writes it. The heights are Excel for Mac's, measured on 723 rows in en-US and ru-RU; Excel for Windows fits rows to its pixels ([docs/formats.md](docs/formats.md#fitting-rows)).

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 'Total', style: 0 } } }], styles: [{ font: { size: 16 } }] })
session.setLayout([{ sheet: 'S', rows: '1', fit: true }])
session.setCellText('S', 'B2', 'two\nlines')
console.log(session.getLayout().sheets[0].rows, session.getBook().styles)
// → {
//     '1': { height: 21, bestFit: true },
//     '2': { height: 32, bestFit: true }
//   } [ { font: { size: 16 } }, { alignment: { wrap: true } } ]
```

A sheet's `conditionalFormats` are Excel's conditional formatting rules: values compared with a formula, formulas that hold, text, blanks, dates, top and average rules, duplicates, colour scales, data bars and icon sets. `Session.getCells` and `getViewport` give each cell of a window the `appearance` its rules give it by the current values, as Excel shows it: the rules' styles merged by priority, a colour scale's colour, a bar, an icon, and the text by a rule's number format.

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{
    name: 'S',
    cells: { A1: { value: 3 }, A2: { value: 12 }, A3: { value: 7 } },
    conditionalFormats: [
      { type: 'cellIs', ranges: ['A1:A3'], operator: 'greaterThan', formulas: ['10'], style: { font: { bold: true } } },
      { type: 'colorScale', ranges: ['A1:A3'], points: [{ value: { type: 'min' }, color: '#FFFFFF' }, { value: { type: 'max' }, color: '#63BE7B' }] },
      { type: 'expression', ranges: ['B1:B3'], formula: '$A1<5', style: { fill: { color: '#FFC7CE' } } },
    ],
  }],
})
const w = session.getCells('S', 'A1:B3')
console.log(w.A2.appearance, w.A3.appearance, w.B1.appearance)
// → { style: { font: { bold: true }, fill: { color: '#63BE7B' } } } { style: { fill: { color: '#BAE3C5' } } } { style: { fill: { color: '#FFC7CE' } } }
```

The rules follow Excel's, measured in Excel for Mac and Windows: a formula is written for the top-left cell of the rule's cells and moves to each, statistics leave out text and errors, a colour moves channel by channel. `readXlsx` and `writeXlsx` keep the rules, Excel 2010's among them, and structure operations move them as Excel does. Edits list no appearances: a host reads its window again where the rules' cells in `repaint` meet it. `setLayout` replaces a sheet's rules ([Layout](#layout)). [docs/conditional-formats.md](docs/conditional-formats.md) has the rules and how they were checked.

### Hidden rows and filters

A row or a column is hidden by `hidden: true` in `rows` or `columns`, as an xlsx file stores it, or by a height or width of 0, as Excel takes it. `SUBTOTAL` with codes 101 to 111 and `AGGREGATE` with options 1, 3, 5 and 7 leave out the cells of hidden rows. Hidden columns count, as in Excel.

A sheet whose autofilter has criteria is marked `filtered: true`. Excel takes every hidden row of such a sheet as one a filter hid, and `SUBTOTAL` with codes 1 to 11 leaves those out too.

```js
import { compute } from 'xlcalc'

const cells = { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 4 }, B1: { formula: 'SUBTOTAL(9,A1:A3)' }, B2: { formula: 'SUBTOTAL(109,A1:A3)' }, B3: { formula: 'AGGREGATE(9,5,A1:A3)' } }
const sheet = { name: 'S', cells, rows: { 2: { hidden: true } } }
const v = compute({ sheets: [sheet] }).sheets[0].cells
console.log(v.B1.value, v.B2.value, v.B3.value) // → 7 5 5
console.log(compute({ sheets: [{ ...sheet, filtered: true }] }).sheets[0].cells.B1.value) // → 5
```

A session hides and shows rows and sets the filter mode through `setLayout`, and recomputes these formulas ([Layout](#layout)). Structure operations treat hidden rows as Excel does: on a filtered sheet they delete and copy the visible rows only ([docs/structure.md](docs/structure.md)).

### Data validation and outlines

A sheet's rules of data validation (`validations`) say what its cells take: a list, whole numbers or dates between bounds, a text length, a custom formula, with the message shown while a cell is selected and the alert for a value it does not take. `session.getValidation(sheet, cell)` gives a cell's rule and the items of its list, computed in the book as it is now, each with the `input` Excel writes when it is picked from the dropdown (a cell's value as it is, a literal item read as typed in en-US; no rule checked); `session.validateEntry(sheet, cell, input)` tells whether a value meets the rule, as Excel checks it: `{ ok: true }`, or the rule with its alert's style (`severity`) and message, the rule's `error` or, without one, the text of Excel's alert in the book's locale (`This value doesn't match the data validation restrictions defined for this cell.`); `setCellText` checks the value it writes so.

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{
    name: 'Orders',
    cells: { A1: { value: 'Status' }, A2: { value: 'Open' }, B2: { value: 5 }, B3: { value: 7 } },
    validations: [
      { range: 'A2:A100', type: 'list', formula1: '"Open,Shipped,Closed"', prompt: 'Pick a status' },
      { range: 'B2:B3', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'warning' },
    ],
  }],
})
console.log(session.getValidation('Orders', 'A7').list.map((i) => i.value)) // → [ 'Open', 'Shipped', 'Closed' ]
console.log(session.validateEntry('Orders', 'A7', { value: 'Lost' }).ok, session.validateEntry('Orders', 'B3', { value: -2 }).rule.errorStyle) // → false warning
```

Rules and outlines change through `setLayout` ([Layout](#layout)), move with structure operations as Excel moves them and go both ways in xlsx files. [docs/validation.md](docs/validation.md) has how a value is checked and how it was measured in Excel.

Rows and columns have outline levels (`outlineLevel`, 1 to 7), Excel's groups: a collapsed group is hidden rows with its summary row, the row past it, marked `collapsed`. `outlineGroups(sheet)` (also from `xlcalc/a1`, without the engine) gives a sheet's groups as Excel draws them, each `{ first, last, level, summary, collapsed }`, from a sheet or its layout; Excel's commands on them are structure operations: `group` and `ungroup` of rows or columns, `hideDetail` and `showDetail` at a line (a group's `−` and `+`), `showLevels` (the level buttons), each one step of the undo in a session. The marks are read and written as Excel keeps them: a group whose rows are all hidden, through `setLayout` too, is collapsed, and an inner group collapsed inside an outer one stays collapsed when the outer one opens.

```js
import { createSession, outlineGroups } from 'xlcalc'

const session = createSession({
  sheets: [{ name: 'S', cells: {}, rows: { '3:4': { outlineLevel: 1 }, '5:7': { outlineLevel: 2 }, '8:10': { outlineLevel: 1 } } }],
})
session.editStructure({ op: 'hideDetail', sheet: 'S', row: 8 })
session.editStructure({ op: 'hideDetail', sheet: 'S', row: 11 })
session.editStructure({ op: 'showDetail', sheet: 'S', row: 11 })
console.log(outlineGroups(session.getLayout().sheets[0]).rows.map((g) => [`${g.first}:${g.last}`, g.collapsed])) // → [ [ '3:10', false ], [ '5:7', true ] ]
console.log(session.editStructure({ op: 'showDetail', sheet: 'S', row: 11 }).code) // → detail
```

[docs/outline.md](docs/outline.md) has the rules, measured in Excel for Windows.

### Date systems

A book is in Excel's 1900 date system unless it has `dateSystem: 1904`, as an xlsx file made on an old Mac may have. In the 1904 system serial 0 is 1 January 1904. Dates typed in, date text, `DATE`, `TODAY` and the date functions count from there, and cells and `TEXT` show a negative date or time with a minus.

```js
import { compute } from 'xlcalc'

const cells = { A1: { formula: 'DATE(2025,1,31)' }, A2: { formula: 'YEAR(A1)' }, A3: { value: -0.5, numberFormat: 'h:mm' } }
const c = compute({ dateSystem: 1904, sheets: [{ name: 'S', cells }] }).sheets[0].cells
console.log(c.A1.value, c.A2.value, c.A3.text) // → 44226 2025 -12:00
```

`displayValue` takes `dateSystem` too. A book in the 1904 system has its own values hash.

`dateToSerial(text, { dateSystem })` gives the serial number of a date as `{ date }` takes it, and `serialToDate(serial, { dateSystem })` the date of a serial number in that form, without a time zone: the date alone, or with the time to the millisecond. Serial 60 of the 1900 system is Excel's phantom 1900-02-29. Neither makes a date before the system's day 0 or after 9999-12-31, which Excel does not have: they throw `InputError`. A `Date` object is not taken, as its time zone would move the day.

```js
import { dateToSerial, serialToDate } from 'xlcalc'

console.log(dateToSerial('2026-01-31'), dateToSerial('2026-01-31T18:00'), dateToSerial('2026-01-31', { dateSystem: 1904 })) // → 46053 46053.75 44591
console.log(serialToDate(46053.75), serialToDate(60), serialToDate(0, { dateSystem: 1904 })) // → 2026-01-31T18:00 1900-02-29 1904-01-01
```

### Circular references

A formula in a circular reference, and every formula that reads one, is not computed in a book without iterative calculation, as Excel computes none of them: each holds its value, the cell's `held` or 0 without it, and the rest of the book is computed. `cached`, the values a file stores (`readXlsx`'s `cached`), opens such a book as Excel opens the file: a cell holds the value stored for it, empty text where the file stores none. A session keeps the values as Excel keeps them: an edit that does not break the cycle changes none of them, a formula entered there holds 0, a cell another edit brings into the cycle's way holds what it showed, and a value typed into a cell of the cycle breaks it, and everything computes again. `getBook()` gives the values in `held`, so a full computation of the session's book gives what the session shows, and `writeXlsx` stores them. Each such cell is in `diagnostics` with `error: 'CYCLE'`. The rules and how they were measured: [docs/cycles.md](docs/cycles.md).

```js
import { compute, createSession } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { formula: 'B1/2+1' }, B1: { formula: 'A1' }, C1: { formula: 'A1*10' }, K1: { value: 3 }, E1: { formula: 'K1*2' } } }] }
const cached = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, B1: { value: 2 } } }] }
const r = compute(book, { cached })
console.log(r.sheets[0].cells.A1.value, r.sheets[0].cells.C1.value, r.sheets[0].cells.E1.value) // → 2  6
console.log(r.diagnostics.map((d) => `${d.cell} ${d.error}`), r.diagnostics[0].reason) // → [ 'A1 CYCLE', 'B1 CYCLE', 'C1 CYCLE' ] circular reference: not computed, the value is held

const session = createSession(book, { cached })
session.setCells([{ sheet: 'S', cell: 'B1', input: { value: 4 } }])
console.log(session.getCell('S', 'A1').value, session.getCell('S', 'C1').value) // → 3 30
session.setCells([{ sheet: 'S', cell: 'B1', input: { formula: 'A1' } }])
console.log(session.getCell('S', 'B1').value, session.getBook().sheets[0].cells.C1) // → 0 { formula: 'A1*10', held: 30 }
```

### Iterative calculation

With `iteration`, as Excel's setting "Enable iterative calculation" stores it in a file, the cells of every cycle, and the cells that depend on them, are computed again and again from 0, as Excel computes a book it opens, and `held` has no effect. The rounds stop after `maxIterations` (100 by default) or when no number changes by `maxChange` (0.001 by default) or more.

```js
import { compute } from 'xlcalc'

const cells = { B1: { value: 1000 }, B2: { value: 0.1 }, B3: { formula: 'B2*(B1+B4)/2' }, B4: { formula: 'B1+B3-100' } }
const r = compute({ sheets: [{ name: 'S', cells }], iteration: {} }).sheets[0].cells
console.log(r.B3.value, r.B4.value) // → 99.999984375 999.999984375
console.log(compute({ sheets: [{ name: 'S', cells }] }).sheets[0].cells.B3.value) // → 0
```

Cells are computed sheet by sheet in the order of the sheets' names, row by row. A change of text, a boolean or an error into another of them does not count as a change. Every computation starts from 0, so a session gives what a full computation gives. Excel goes on from the values it holds when it recalculates, which the engine does not do (see [Limitations](#limitations)).

### Links to other workbooks

A formula can read another workbook: `[1]Sheet1!A1`, `'[1]My Sheet'!A1:B9`, `[1]!Rate`. The book holds what Excel stores of that workbook beside the formulas in a file, its `externalBooks`: where it is, its sheets and the values of the cells the formulas read, its names. Element `i` answers `[i+1]`. The engine computes such references from these stored values, as Excel computes a file it opens without updating its links; the other workbook is never read. `readXlsx` reads them from a file's links and `writeXlsx` writes them back.

```js
import { compute, toLocal } from 'xlcalc'

const externalBooks = [{ target: 'Prices.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 5, A2: 7, B1: 'pear' } }] }]
const book = { sheets: [{ name: 'S', cells: { A1: { formula: "[1]Sheet1!A1+1" }, A2: { formula: 'SUM([1]Sheet1!A:A)' }, A3: { formula: "COUNTIF([1]Sheet1!A:A,5)" } } }], externalBooks }
const cells = compute(book).sheets[0].cells
console.log(cells.A1.value, cells.A2.value, cells.A3.value, cells.A1.formula) // → 6 12 { error: 'VALUE' } '[1]Sheet1'!A1+1
console.log(toLocal(cells.A1.formula, 'en-US', { externalBooks })) // → { ok: true, formula: '[Prices.xlsx]Sheet1!A1+1' }
```

A cell the stored values do not hold is blank, as Excel takes it, unless the sheet's last update failed (`refreshError`) or nothing of the sheet is stored: then it is `#REF!`. The criteria functions' ranges (`SUMIF`, `COUNTIFS` and the others), `COUNTBLANK` and the database functions take only an open book's reference and give `#VALUE!` over a link, as in Excel; `INDEX`, `MATCH`, `VLOOKUP`, `XLOOKUP`, `SUMPRODUCT` and the rest read the stored values. A formula with a link entered in a session makes Excel mark that workbook as it cannot read it, and so does the session: its cells the stored values lack are `#REF!` from then on. Formulas are stored with the link's number; `toLocal`, `toInvariant`, `parseEntry`, a window's `formulaLocal` and `editText` show and read them as Excel does, by the file name (`[Prices.xlsx]Sheet1!A1`), and refuse a workbook the book has no link to. [docs/links.md](docs/links.md) has the rules and how they were measured.

### Structure operations

`editStructure(book, op)` inserts or deletes rows and columns, moves (cut and paste) or copies (copy and paste, fill) a range, renames or deletes a sheet, and adjusts every reference in formulas and names on every sheet as Excel adjusts it: ranges grow and shrink, lost references become `#REF!`, absolute references move with inserted rows. It returns the new book and `mapCell`, which says where a cell of the old book is now (null for a cell deleted or overwritten, or of a sheet the old book does not have); the input book is not changed.

```js
import { editStructure } from 'xlcalc'

const book = {
  sheets: [{
    name: 'Sales',
    cells: {
      A1: { value: 120 }, A2: { value: 80 }, A3: { value: 45 },
      C1: { formula: 'SUM(A1:A3)' }, C2: { formula: 'A3*2' },
    },
  }],
}

const ins = editStructure(book, { op: 'insertRows', sheet: 'Sales', row: 2 })
console.log(ins.book.sheets[0].cells.C1.formula, ins.book.sheets[0].cells.C3.formula) // → SUM(A1:A4) A4*2
console.log(ins.mapCell({ sheet: 'Sales', cell: 'A3' })) // → { sheet: 'Sales', cell: 'A4' }

const del = editStructure(book, { op: 'deleteRows', sheet: 'Sales', row: 3 })
console.log(del.book.sheets[0].cells.C1.formula, del.book.sheets[0].cells.C2.formula) // → SUM(A1:A2) #REF!*2
const last = editStructure(book, { op: 'deleteSheet', sheet: 'Sales' })
console.log(last.ok, last.code) // → false last_sheet
```

The operations are `insertRows` and `deleteRows` (`row`, a row number, and `count`), `insertColumns` and `deleteColumns` (`column`, letters, and `count`), `moveRange` (`from: { sheet, range }`, `to: { sheet, cell }`), `copyRange` (`from` and `to` as `{ sheet, range }`; a target that is a whole multiple of the source is tiled, as in fill down; copies have no `meta` unless `withMeta: true`), `renameSheet` (`name`, by Excel's rules for sheet names), `deleteSheet`, `insertSheet` (`name`, `before`), `moveSheet` (`sheet`, `before`; Excel places a sheet by the visible ones), `hideSheet` (`sheet`, `very` for a very hidden sheet) and `unhideSheet` (`sheet`), over tables `addTable` (`sheet`, `table`), `renameTable` (`table`, `name`), `deleteTable` and `convertTable` (`table`; [docs/tables.md](docs/tables.md)), and Excel's outline commands `group` and `ungroup` (`sheet`, `rows` or `columns`), `hideDetail` and `showDetail` (`sheet`, `row` or `column`) and `showLevels` (`sheet`, `rows`, `columns`; [docs/outline.md](docs/outline.md)), and Excel's Merge Cells and Unmerge Cells, `mergeCells` (`sheet`, `range`, `across` for Merge Across, `center` for Merge & Center) and `unmergeCells` (`sheet`, `range`): a merge keeps the top-left cell, or moves the area's first value into it when it is empty, and clears the rest ([docs/structure.md](docs/structure.md#merging-cells)), and Excel's fill handle, `fillRange` (`sheet`, `range`, `to` that holds `range` and reaches past it, `type` `default`, `series` or `copy`): copies of the source with numbers, dates, times, text with a number and names of days and months going on as Excel continues them ([docs/structure.md](docs/structure.md#the-fill-handle)). Only the references that change are rewritten, in place; a formula that does not parse moves with its cell as written. An operation Excel refuses returns `{ ok: false, code, message }`, the message in the book's locale (`Нельзя изменить часть массива S!C1 (C1:C3)` in a ru-RU book), with `off_sheet` (cells would leave the sheet), `sheet_name` (a name Excel would not take, `History` among them, or one another sheet has), `last_sheet` (the last visible sheet deleted or hidden), `very_hidden` (a very hidden sheet deleted or moved), `merged` (a move cutting a merged area, a copy landing on part of one) `filtered` (a copy without the rows a filter hid, pasted over hidden rows of a filtered sheet) `array` (part of an old-kind array formula's area would change, or a merge would cover one), `table` (rows or columns into two tables at once, or deleted from two tables unless both go whole, a header row deleted alone, cells of a table merged), `table_name`, `outline_deep` (Group of lines at the deepest level), `not_grouped` (Ungroup of lines in no group, a group's button where there is none) or `detail` (Show Detail of a group shown, Hide Detail of one hidden). A name whose reference is lost gives `#REF!` where it is used, as in Excel; the parser reads the lost references Excel writes (`'S'!#REF!`, `#REF!A1`) as `#REF!`.

A book keeps styles as xlsx does: `styles`, a table of cell styles (font, solid fill, borders, alignment), `style` in a cell, an index into it, and per sheet `columns` and `rows` (by line `C`, `5` or span `B:D`, `5:9`: a width or height, a style, a number format) and `merges` (`B2:D3`). A cell without its own style or format shows its row's, else its column's. Computation ignores them, but a merged area stops an array spilling over it, as in Excel; the operations change them as Excel does: inserted rows take the format of the row above (its height, style and format, its cells' styles and formats) and columns of the column on the left, merged areas grow, shrink and move like ranges, moved and copied cells keep their look, whole rows and columns take their sizes and styles along. A cell's note and hyperlink go with the cell as its style does, frozen panes keep the rows and columns before their split, rules of data validation move as references do and lines inserted right after a rule take it, and inserted lines take the outline level of the line before them. The rules, and how they were checked against Excel on 317 books and the valid sets, merging on 61 more and the fill handle on 122 in Excel for Windows, are in [docs/structure.md](docs/structure.md).

A session applies the same operations to its book with `session.editStructure(op)` and recomputes, keeping the book's styles, so a host can hold the book in the session alone and save `getBook()` as a version:

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{ name: 'Plan', cells: { A1: { value: 100 }, A2: { value: 250 }, A3: { formula: 'SUM(A1:A2)', style: 0 } } }],
  styles: [{ font: { bold: true } }],
})
const r = session.editStructure({ op: 'deleteRows', sheet: 'Plan', row: 1 })
console.log(r.mapCell({ sheet: 'Plan', cell: 'A3' }), r.moved)
// → { sheet: 'Plan', cell: 'A2' } [ { kind: 'lines', sheet: 'Plan', axis: 'rows', at: 1, count: -1 } ]
console.log(r.changes.map((c) => [c.cell, c.value, c.formula])) // → [ [ 'A2', 250, 'SUM(A1:A1)' ] ]
console.log(session.getBook().sheets[0].cells) // → { A1: { value: 250 }, A2: { formula: 'SUM(A1:A1)', style: 0 } }
```

It answers `mapCell`, as `editStructure` does, but not the book, which `getBook()` gives; `moved`, how the addresses moved (`{ kind: 'lines', sheet, axis, at, count }` for rows and columns inserted or deleted, a range moved, a sheet renamed, deleted, inserted or moved); and `changes`, the cells whose value or formula text differs from those of the cell that was at their old place (`SUM(A1:A2)` above became `SUM(A1:A1)`), so an interface that moves its window by `moved` (or each cell by `mapCell`) and applies `changes` is up to date. It builds the state again, at the cost of a cold computation and a comparison of the cells with those of their old places (170 ms for the 3,400-row synthetic book), but for a sheet hidden or shown (`hideSheet`, `unhideSheet`), which changes no value, the outline commands, which recompute only the `SUBTOTAL` and `AGGREGATE` formulas over the rows they hide and show, and a merge that moves no cell, which writes the cells of its range as an edit does. Each operation is one step of the undo ([Undo and redo](#undo-and-redo)).

### Reading and writing xlsx

`readXlsx(data)` from `xlcalc/xlsx` reads a file's bytes (a `Uint8Array`, a Node `Buffer` or an `ArrayBuffer`: `await readFile(path)`, `await file.arrayBuffer()`) into a book. Beside the book it gives `cached`, the values Excel stored for formulas and spilled cells in the shape of a result, and `issues`, everything the book cannot hold, such as rich text, with the sheet and cells. `compute(read.book, { cached: read.cached })` and `createSession` with the same option open a book with a circular reference as Excel opens the file ([Circular references](#circular-references)); nothing else is taken from `cached`. `writeXlsx({ book, result })` writes a book with the values of its formulas: `result` is `compute`'s, a session's (`session.getSnapshot()` gives the book and its result in one call) or the `cached` of a file read. Formulas are written as dynamic array formulas with their spills, and `mode: 'values'` writes values only.

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

Both functions return promises: the archive is inflated and deflated by the platform's streams. Hidden and very hidden sheets (`hidden: true` and `hidden: 'very'`; the writer makes the first visible sheet the active one), styles, number formats (built-in ones by their number), widths, heights, hidden rows and columns, outline levels, merged areas, notes, hyperlinks, frozen panes, conditional formats (Excel 2010's rules too), rules of data validation, names, tables, a sheet's filter mode, the date system, iterative calculation and the book's locale go both ways. A cell without a style or a format of its own is written with its row's, else its column's, as Excel writes a value typed there, so a date in a column of dates shows as a date in Excel, and it reads back without one. A cell the file stores with a look alone, as Excel stores each cell of a formatted area, is read into its sheet's `looks` ([Looks of many cells](#looks-of-many-cells)), and the cells of `looks` are written back as such cells. A group's collapsed mark is written as Excel keeps and saves it: on the summary row of a group whose rows are all hidden, rows hidden through `setLayout` too, and not on one whose rows show ([docs/outline.md](docs/outline.md#in-xlsx-files)).

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

A thread of comments, as Excel 365 makes them, is read as a note: its first author, and every comment's text with the replies' authors. [docs/xlsx.md](docs/xlsx.md) lists what is read, written and reported, and how it was checked against Excel.

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

### Computing without blocking

`createInlineCalc()` computes in the same thread, in slices that let timers and input run between them; `createCalcPool()` from `xlcalc/worker` computes in a pool of threads, where a session's book lives in its thread. Both have one API: `compute`, `createSession` and the session's methods return promises, and every call takes `timeoutMs`, a soft deadline of computing time, and `signal`, an `AbortSignal`. A call that fails changes nothing. `limits` bound the formulas of a book, the values of one array, the cells of one spill and the cells one edit of a session writes (`writes`: a copy of one cell onto a whole column writes 1,048,576), in the synchronous `compute` and `createSession` too; over them the call throws `CalcError` rather than computing or writing.

```js
import { createInlineCalc } from 'xlcalc'
import { createCalcPool } from 'xlcalc/worker'

const book = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*10' }, A3: { formula: 'SUM(SEQUENCE(1000,1000))' } } }] }

const calc = createInlineCalc({ limits: { cells: 100_000 } })
const e = await calc.compute(book).catch((e) => e)
console.log(e.code, e.limit, e.cell) // → limit cells { sheet: 'S', cell: 'A3' }

const pool = createCalcPool({ threads: 2, hardTimeoutMs: 10_000 })
const session = await pool.createSession(book, { clock: { now: '2026-10-07T12:00:00Z', seed: 1 } })
console.log((await session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }], { timeoutMs: 1_000 })).changes)
// → [
//     { sheet: 'S', cell: 'A1', value: 3, text: '3' },
//     { sheet: 'S', cell: 'A2', value: 30, text: '30' }
//   ]
const stop = new AbortController()
const first = session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 4 } }])
const second = session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 5 } }], { signal: stop.signal }).catch((e) => e)
stop.abort() // the second call waits for the first: it is refused at once
await first
console.log((await second).code, (await session.getResult()).sheets[0].cells.A2.value) // → aborted 40
const bounded = await pool.createSession(book, { limits: { writes: 100_000 } })
const w = await bounded.editStructure({ op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', range: 'B:B' } }).catch((e) => e)
console.log(w.code, w.limit, w.size, bounded.revision) // → limit writes 1048576 0
await pool.close()
```

A task that runs past the pool's `hardTimeoutMs` terminates its thread, and in Node `maxHeapMb` bounds each thread's heap. The pool keeps a snapshot and a journal of each session's book, so a session whose thread was lost comes back on its next call as it was before the call that failed, its history of undo and redo too. With `maxHeapMb` a session's history holds `maxHeapMb / (8 × maxSessionsPerThread)` MB by default, so the histories of a full thread take an eighth of its heap; a session's own `undo.maxMb` takes its place. [docs/async.md](docs/async.md) has the rules, the error codes and what this costs.

## API

From `xlcalc`:

| export | |
|---|---|
| `compute(book, opts?)` | computes the whole book; `opts.clock`, `opts.externals`, `opts.execution` (`'closures'`, the default, or `'eval'`), `opts.display` (`true`, the default, or `false`), `opts.limits` (`formulas`, `cells`, `spill`; over them `CalcError`), `opts.cached` (a file's stored values, which the formulas a circular reference stops hold). Returns `ComputeResult`: `sheets` (values, shown `text` and `color` with the display on, canonical formulas, `numberFormat`, `meta`), `diagnostics` (each cell with an error or stopped by a circular reference, `kind: 'cell'`, then each name with a problem, `kind: 'name'`), `volatile`, `clock`, `hash`, `inputHash`, `configVersion`, `ms` |
| `createSession(book, opts?)` | a `Session`; each call that changes the book answers an `EditResult` (`{ revision, changes, layout?, styles?, moved?, repaint? }`) and takes `{ baseRevision }`: `setCells(edits)` (the fields given, the cell's look kept), `replaceCells(edits)` (cells whole), `clearCells(ranges, 'contents' \| 'formats' \| 'all')`, `patchCells(patches)` (the `style`, by number or by its fields, `numberFormat`, `meta`, `note`, `link` of a cell or a range, whole rows and columns too, without values), `setCellText(sheet, cell, text, { widen, refuseInvalid })` and `previewCellText` (typed text, `{ ok, …, width?, validation? }` or a refusal; with `refuseInvalid` a value its rule of data validation does not take is refused, `code: 'validation'`, and nothing written), `setCellTexts(entries, { widen, refuseInvalid })` and `previewCellTexts` (texts `{ sheet, cell, text }` typed at once, in one step: `{ ok, …, widths?, validations? }` or a refusal with its `sheet` and `cell`), `pasteText(sheet, cell, text)` and `previewPasteText` (text from the clipboard as Excel pastes it: `{ ok, …, range }` or a refusal with its `sheet` and `cell`), `fillCellText(sheet, range, text, { widen, active, refuseInvalid })` and `previewFillCellText` (Control+Return: `{ ok, …, widths?, validation? }` or a refusal with its `sheet` and `cell`), `setNames(names)`, `setClock(clock)`, `editStructure(op)`, `setStyles(styles)`, `setLayout(edits)` (columns, rows, merged areas, the filter mode, frozen panes, conditional formats, rules of data validation, outline settings; columns AutoFit with `{ columns, fit: true }`), `revision`, `getBook()`, `getLayout()` (the book without its cells), `getUsedRange(sheet)` (`{ rows, columns }`, where its cells end, spills too), `getEnd(sheet, cell, 'up' \| 'down' \| 'left' \| 'right')` (the address Control and an arrow key come to, as `Range.End`), `explain(sheet, cell, { depth, maxNodes, expand, errorDepth })` (an error's `origin` marked and the way to it), `getCells(sheet, range or ranges)` (with each cell's `style`, `note`, `link`, an old-kind array formula's `array`, conditional formats' `appearance` and its own format's `numberFormatCode`), `copyText(sheet, range, { trailingNewline, maxCells, maxChars })` (the selection as plain text from full formatted values, skipping hidden lines, without changing the book), `getCell(sheet, cell)` (one cell of `getCells`, or undefined), `getViewport(sheet, range, { formulaLocal, diagnostics, maxCells, rows, columns, editText })` (a window for a grid: cells with their style, by number, and number format resolved, empty ones with a line's style, the styles they show, the lines' numbers and sizes, merged areas, tables, frozen panes, the used range; only some rows and columns, the text each cell opens with for editing), `getCellIssue(sheet, cell)` (an error's reason and `origin`, or null), `getValidation(sheet, cell)` (a cell's rule, its formulas as at the cell, a list's items), `validateEntry(sheet, cell, input)` (`{ ok: true }` or `{ ok: false, code, message, rule, severity }`), `getResult()`, `getSnapshot()` (`{ book, result }` in one call, what `writeXlsx` takes), `getHash()` (the values hash without the result), `undo()` and `redo()` (the call's `EditResult` with `call`, `cells?`, `mapCell?`, or null), `history` (`{ undo, redo }`, the steps each way), `subscribe(listener)` (each change of the book, whoever makes it, with its `EditResult` and `method`; answers the function that unsubscribes: [Following changes](#following-changes)), `close()`, `clock`. Options: `clock`, or `live: true`, `externals`, `execution`, `display`, `limits` (those of `compute` and `writes`, the cells one edit writes), `undo` (`{ depth, maxMb }`, 100 steps and 64 MB by default) and `cached`, as `compute` takes it; edits are compiled in the session's execution mode. |
| `viewportOf(saved, sheet, range, opts?)` | the window `getViewport` gives, over a book and its result as a host saved them (`{ book, result }`), without a session and without computing; no `revision`, and conditional `appearance` only with `appearance: true`, computed over the result's values ([A window for a grid](#a-window-for-a-grid)) |
| `displayValue(value, opts?)` | a value as a cell shows it, `{ text, color? }`; `opts`: `numberFormat`, `locale` (a loaded locale's tag or a locale object, as every function's `locale` below), `width` (no limit without it), `dateSystem` ([Showing values](#showing-values)) |
| `builtinFormatCode(id, locale?)` | the code a built-in number format (0 to 49) shows by in the locale, undefined for a number that is none ([Showing values](#showing-values)) |
| `numberFormatCategory(format, locale?)` | the category of a number format (a code or a built-in format's number) as the ribbon's Number Format box shows it, `{ category, name }` ([Showing values](#showing-values)) |
| `numberFormatChoices(locale?)` | the items of the ribbon's Number Format list in order, `{ key, name, format, code }` ([Showing values](#showing-values)) |
| `dateToSerial(text, opts?)`, `serialToDate(serial, opts?)` | a date as `{ date }` takes it and its serial number, in `opts.dateSystem` ([Date systems](#date-systems)) |
| `validateFormula(text, ctx?)` | `{ ok: true, canonical, functions, sheets, names, externals }` or `{ ok: false, code, message, at }`; `ctx`: `sheets`, `names`, `externalFunctions` (without them sheets and names are not checked), `tables`, and `table`, the table of the formula's cell ([Tables](#tables)), `externalBooks`, the links a reference to another workbook is checked and numbered by ([Links to other workbooks](#links-to-other-workbooks)) |
| `validateName(name, opts?)` | `{ ok: true }` or `{ ok: false, code, message }` (`chars`, `reference`, `boolean`) by the rules of a book of `opts.locale` |
| `validateNames(names, ctx)` | the name issues of a book with `ctx.sheets`, `ctx.tables`, `ctx.externalFunctions`, `ctx.externalBooks` and `ctx.locale`, as `compute` gives them, each with `dropped` |
| `parseEntry(text, opts?)` | text typed into a cell as Excel takes it: `{ ok: true, input }` (the cell's new `CellInput`, or `null`) or `{ ok: false, code, message, at }`; `opts`: `locale`, `cell`, `now`, `dateSystem`, `externalBooks` (a typed reference to a link numbered, one to a book the links lack refused) ([Typed entry](#typed-entry)) |
| `entryText(value, opts?)`, `needsApostrophe(text, opts?)` | the text a cell opens with for editing, as the formula bar shows its value; whether typed text would read as something else than itself; `opts`: `numberFormat`, `locale`, `dateSystem`, `now` ([Addresses and helpers](#addresses-and-helpers)) |
| `validateSheetName(name, opts?)` | `{ ok: true }` or `{ ok: false, code, message, at? }` by Excel's rules for sheet names; `opts.locale` gives the message's language |
| `errorText(code, locale?)` | an error as a cell shows it (`#DIV/0!`, `#ДЕЛ/0!` in ru-RU) |
| `scanReferences(formula, opts?)` | the references of formula text as it is typed, with their places, rectangles, sheets and tables, a token that does not read skipped; `opts.locale` for text typed in a locale; also from `xlcalc/text` |
| `escapeColumn(name)` | a column's name in a structured reference, `'` before `[`, `]`, `#`, `'` and a leading `@`; also from `xlcalc/text` ([Tables](#tables)) |
| `columnWidthToPoints(width?)`, `columnWidthFromPoints(points)` | stored xlsx widths and whole points in the library's Mac geometry, for a grid and a dragged column; null or missing width is 53 points, the inverse rounds points down; also from `xlcalc/text` ([Column geometry](docs/formats.md#column-geometry-for-a-grid)) |
| `formatClipboard(rows, opts?)` | a rectangle of texts as Excel copies plain text, with quoted control characters, CRLF rows and a terminal CRLF unless `trailingNewline: false`; also from `xlcalc/text` |
| `parseClipboard(text)` | text from the clipboard as the rectangle Excel pastes: rows at line breaks, fields at tabs, a quoted field to its closing quote, each row as wide as the longest; also from `xlcalc/text` ([Typed entry](#typed-entry)) |
| `internStyle(styles, style)` | a style's number in a style table, found as `patchCells` finds it: `{ index, styles }`, the same array when an equal style is there, else a new one with the style added ([Editing a grid](#editing-a-grid)) |
| `sameValue(a, b)`, `canonValue(value)` | values compared as the values hash compares them, and a value's text in the hash |
| `parseCell`, `cellAddress`, `colLetters`, `colNumber`, `parseRange`, `rangeAddress`, `parseLines`, `linesKey`, `patchLines`, `outlineGroups`, `rectsMeet`, `rectContains`, `MAX_ROWS`, `MAX_COLS`, `quoteSheetIfNeeded`, `nameKey`, `sameName` | addresses and names without the engine, also from `xlcalc/a1` ([Addresses and helpers](#addresses-and-helpers)) |
| `fitColumns(book, result, sheet, columns)` | the widths AutoFit gives columns `C` or `B:D` of a computed book, by letter (null for the default width; a column with nothing to show left out) ([Showing values](#showing-values)) |
| `entryWidth(value, opts?)` | the width a column takes as a typed value widens it, or undefined; `opts`: `numberFormat`, `locale`, `dateSystem`, the column's `width` and `bestFit` ([Showing values](#showing-values)) |
| `fitRows(book, result, sheet, rows)` | the heights AutoFit gives rows `5` or `2:9` of a computed book, by spans of rows of one height (null for the default height) ([Showing values](#showing-values)) |
| `toLocal(formula, locale, opts?)`, `toInvariant(text, locale, opts?)` | a stored formula as the locale shows it, typed text as a stored formula: `{ ok: true, formula }` or `{ ok: false, code, message, at }` ([Localization](#localization)); `opts.table`, the table of the formula's cell, shows its references as Excel does inside it (`[@Qty]`) and names it again (`Sales[@Qty]`, [docs/tables.md](docs/tables.md)); `opts.externalBooks` shows a link's reference by its file (`[Prices.xlsx]Sheet1!A1`) and reads one back, a book they lack `external_ref` ([Links to other workbooks](#links-to-other-workbooks)) |
| `validateLocalFormula(text, locale, ctx?)` | `validateFormula` for text typed in the locale, positions in `text` |
| `registerLocale(locale)` | loads a locale module of the package (`ruRU` of `xlcalc/locales/ru-RU`) for books in it and for its tag ([Localization](#localization)) |
| `LOCALES` | the loaded locales by tag, en-US and those registered, each a `Locale` (`LOCALES.get('ru-RU')`): the tags `BookInput.locale` takes |
| `enUS` | en-US's `Locale`, built in |
| `editStructure(book, op)` | one structure operation (`StructureOp`): `{ ok: true, book, mapCell }` or `{ ok: false, code, message }`; [docs/structure.md](docs/structure.md) |
| `copyFormula(text, from, to, book?)` | the text of a formula copied from cell `from` to `to` (`{ sheet, cell }`), as `copyRange` pastes it; with the book, a formula leaving its table names it ([docs/structure.md](docs/structure.md)) |
| `FUNCTION_NAMES`, `CONFIG_VERSION` | the whitelist; the configuration version that goes into every hash |
| `createInlineCalc(opts?)` | a `Calc` computing in the same thread in portions: `compute`, `createSession` (an `AsyncSession`) and `close`, as promises; `opts`: `sliceMs`, `timeoutMs`, `limits` ([docs/async.md](docs/async.md)) |
| `InputError` | thrown on input of the wrong shape |
| `CalcError` | an asynchronous call that failed, or a book or an edit over its limits; `code` says why ([docs/async.md](docs/async.md)) |
| `debugGeneratedCode(book, opts?)` | the JavaScript that the `eval` mode generates, in evaluation order, for debugging; not part of the contract |

`xlcalc/contract` exports the types and two constants (`MAX_FORMULA_LENGTH`, `MAX_FORMULA_DEPTH`) without any engine code, for code that only builds books or reads results.

`xlcalc/a1` exports the addresses and names of [Addresses and helpers](#addresses-and-helpers) (`parseCell`, `parseRange`, `rangeAddress`, `parseLines`, `patchLines`, `quoteSheetIfNeeded`, `sameName` and the rest) and a sheet's outline groups (`outlineGroups`), the same functions as the root's, without any engine code.

`xlcalc/text` exports `scanReferences`, `refAt`, `quoteSheetIfNeeded` and `escapeColumn`, the root's functions of formula text, for a browser that colours a formula as it is typed and points at cells without the engine, and the column geometry pair `columnWidthToPoints` and `columnWidthFromPoints`, and `formatClipboard`, the plain text a copy writes, and `parseClipboard`, the rectangle a paste fills, for a grid to show it before the paste: 45 KB (17 KB gzip) bundled and minified, the engine's tokenizer, its parser and the translation of a locale's formula text, without any computation. Of the locales' data it carries en-US's messages alone, which the tokenizer's errors take: text typed in another locale takes the locale as an object (`ruRU` of `xlcalc/locales/ru-RU`, or its formula text alone), and a tag that is not loaded is refused with that advice; the root's advice, which names the module to import and `registerLocale`, comes with the root. Neither `xlcalc/text` nor `xlcalc/a1` holds the string `xlcalc`, so a host that looks for it in its bundle as the sign of the engine finds none in them. `entryText` and `needsApostrophe` stay in the root: they read text as `parseEntry` does, with the parser and number formats.

`xlcalc/worker` exports `createCalcPool(opts?)`, a `Calc` over a pool of threads (`opts`: `threads`, `maxHeapMb`, `hardTimeoutMs`, `timeoutMs`, `maxQueue`, `maxSessions`, `maxSessionsPerThread`, `limits`, `snapshotEvery`, `sliceMs`, `spawn`), and `CalcError`.

`xlcalc/xlsx` exports `readXlsx(data)`, which resolves to `{ book, cached, issues }`, and `writeXlsx({ book, result? }, opts?)`, which resolves to `{ data, issues }` (`opts`: `values` `'require'`, `'omit'` or `'preserve'`, `mode` `'formulas'` or `'values'`), with their types `XlsxRead`, `XlsxCached`, `XlsxCachedCell`, `XlsxSnapshot`, `XlsxWritten`, `WriteXlsxOptions`, `XlsxIssue` and `XlsxIssueCode`, and `fromLegacy(formula, { names?, sheet? })`, a file's formula of the old kind as the dynamic formula Excel 365 shows for it, canonical text ([docs/legacy.md](docs/legacy.md)).

`xlcalc/locales/ru-RU` exports `ruRU`, ru-RU's `Locale`, for `registerLocale` and for the functions that take a locale; importing it does nothing else.

`xlcalc/unstable/internal` (the parser, the criteria index, engine state) and `xlcalc/unstable/synth` (a seeded synthetic book for tests and benchmarks) exist for benchmark harnesses and adapters of other engines. They are outside semver and change in any release.

## Functions

The whitelist, by category (`FUNCTION_NAMES` has it in code):

- Math and trigonometry: `ABS`, `ACOS`, `ACOSH`, `ACOT`, `ACOTH`, `AGGREGATE`, `ARABIC`, `ASIN`, `ASINH`, `ATAN`, `ATAN2`, `ATANH`, `BASE`, `CEILING`, `CEILING.MATH`, `CEILING.PRECISE`, `COMBIN`, `COMBINA`, `COS`, `COSH`, `COT`, `COTH`, `CSC`, `CSCH`, `DECIMAL`, `DEGREES`, `EVEN`, `EXP`, `FACT`, `FACTDOUBLE`, `FLOOR`, `FLOOR.MATH`, `FLOOR.PRECISE`, `GCD`, `INT`, `ISO.CEILING`, `LCM`, `LN`, `LOG`, `LOG10`, `MDETERM`, `MINVERSE`, `MMULT`, `MOD`, `MROUND`, `MULTINOMIAL`, `MUNIT`, `ODD`, `PERCENTOF`, `PI`, `POWER`, `PRODUCT`, `QUOTIENT`, `RADIANS`, `RAND`, `RANDARRAY`, `RANDBETWEEN`, `ROMAN`, `ROUND`, `ROUNDDOWN`, `ROUNDUP`, `SEC`, `SECH`, `SEQUENCE`, `SERIESSUM`, `SIGN`, `SIN`, `SINH`, `SQRT`, `SQRTPI`, `SUBTOTAL`, `SUM`, `SUMIF`, `SUMIFS`, `SUMPRODUCT`, `SUMSQ`, `SUMX2MY2`, `SUMX2PY2`, `SUMXMY2`, `TAN`, `TANH`, `TRUNC`

- Statistical and compatibility: `AVEDEV`, `AVERAGE`, `AVERAGEA`, `AVERAGEIF`, `AVERAGEIFS`, `BETA.DIST`, `BETA.INV`, `BETADIST`, `BETAINV`, `BINOM.DIST`, `BINOM.DIST.RANGE`, `BINOM.INV`, `BINOMDIST`, `CHIDIST`, `CHIINV`, `CHISQ.DIST`, `CHISQ.DIST.RT`, `CHISQ.INV`, `CHISQ.INV.RT`, `CHISQ.TEST`, `CHITEST`, `CONFIDENCE`, `CONFIDENCE.NORM`, `CONFIDENCE.T`, `CORREL`, `COUNT`, `COUNTA`, `COUNTBLANK`, `COUNTIF`, `COUNTIFS`, `COVAR`, `COVARIANCE.P`, `COVARIANCE.S`, `CRITBINOM`, `DEVSQ`, `EXPON.DIST`, `EXPONDIST`, `F.DIST`, `F.DIST.RT`, `F.INV`, `F.INV.RT`, `F.TEST`, `FDIST`, `FINV`, `FISHER`, `FISHERINV`, `FORECAST`, `FORECAST.LINEAR`, `FREQUENCY`, `FTEST`, `GAMMA`, `GAMMA.DIST`, `GAMMA.INV`, `GAMMADIST`, `GAMMAINV`, `GAMMALN`, `GAMMALN.PRECISE`, `GAUSS`, `GEOMEAN`, `GROWTH`, `HARMEAN`, `HYPGEOM.DIST`, `HYPGEOMDIST`, `INTERCEPT`, `KURT`, `LARGE`, `LINEST`, `LOGEST`, `LOGINV`, `LOGNORM.DIST`, `LOGNORM.INV`, `LOGNORMDIST`, `MAX`, `MAXA`, `MAXIFS`, `MEDIAN`, `MIN`, `MINA`, `MINIFS`, `MODE`, `MODE.MULT`, `MODE.SNGL`, `NEGBINOM.DIST`, `NEGBINOMDIST`, `NORM.DIST`, `NORM.INV`, `NORM.S.DIST`, `NORM.S.INV`, `NORMDIST`, `NORMINV`, `NORMSDIST`, `NORMSINV`, `PEARSON`, `PERCENTILE`, `PERCENTILE.EXC`, `PERCENTILE.INC`, `PERCENTRANK`, `PERCENTRANK.EXC`, `PERCENTRANK.INC`, `PERMUT`, `PERMUTATIONA`, `PHI`, `POISSON`, `POISSON.DIST`, `PROB`, `QUARTILE`, `QUARTILE.EXC`, `QUARTILE.INC`, `RANK`, `RANK.AVG`, `RANK.EQ`, `RSQ`, `SKEW`, `SKEW.P`, `SLOPE`, `SMALL`, `STANDARDIZE`, `STDEV`, `STDEV.P`, `STDEV.S`, `STDEVA`, `STDEVP`, `STDEVPA`, `STEYX`, `T.DIST`, `T.DIST.2T`, `T.DIST.RT`, `T.INV`, `T.INV.2T`, `T.TEST`, `TDIST`, `TINV`, `TREND`, `TRIMMEAN`, `TTEST`, `VAR`, `VAR.P`, `VAR.S`, `VARA`, `VARP`, `VARPA`, `WEIBULL`, `WEIBULL.DIST`, `Z.TEST`, `ZTEST`

- Logical: `AND`, `FALSE`, `IF`, `IFERROR`, `IFNA`, `IFS`, `LET`, `NOT`, `OR`, `SWITCH`, `TRUE`, `XOR`

- Information: `ERROR.TYPE`, `ISBLANK`, `ISERR`, `ISERROR`, `ISEVEN`, `ISLOGICAL`, `ISNA`, `ISNONTEXT`, `ISNUMBER`, `ISODD`, `ISREF`, `ISTEXT`, `N`, `NA`, `TYPE`

- Text: `ARRAYTOTEXT`, `BAHTTEXT`, `CHAR`, `CLEAN`, `CODE`, `CONCAT`, `CONCATENATE`, `DOLLAR`, `EXACT`, `FIND`, `FINDB`, `FIXED`, `LEFT`, `LEFTB`, `LEN`, `LENB`, `LOWER`, `MID`, `MIDB`, `NUMBERVALUE`, `PROPER`, `REPLACE`, `REPLACEB`, `REPT`, `RIGHT`, `RIGHTB`, `SEARCH`, `SEARCHB`, `SUBSTITUTE`, `T`, `TEXT`, `TEXTAFTER`, `TEXTBEFORE`, `TEXTJOIN`, `TEXTSPLIT`, `TRIM`, `UNICHAR`, `UNICODE`, `UPPER`, `VALUE`, `VALUETOTEXT`

- Date and time: `DATE`, `DATEDIF`, `DATEVALUE`, `DAY`, `DAYS`, `DAYS360`, `EDATE`, `EOMONTH`, `HOUR`, `ISOWEEKNUM`, `MINUTE`, `MONTH`, `NETWORKDAYS`, `NETWORKDAYS.INTL`, `NOW`, `SECOND`, `TIME`, `TIMEVALUE`, `TODAY`, `WEEKDAY`, `WEEKNUM`, `WORKDAY`, `WORKDAY.INTL`, `YEAR`, `YEARFRAC`

- Financial: `ACCRINT`, `ACCRINTM`, `AMORDEGRC`, `AMORLINC`, `COUPDAYBS`, `COUPDAYSNC`, `COUPNCD`, `COUPNUM`, `COUPPCD`, `CUMIPMT`, `CUMPRINC`, `DB`, `DDB`, `DISC`, `DOLLARDE`, `DOLLARFR`, `DURATION`, `EFFECT`, `FV`, `FVSCHEDULE`, `INTRATE`, `IPMT`, `IRR`, `ISPMT`, `MDURATION`, `MIRR`, `NOMINAL`, `NPER`, `NPV`, `ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`, `PDURATION`, `PMT`, `PPMT`, `PRICE`, `PRICEDISC`, `PRICEMAT`, `PV`, `RATE`, `RECEIVED`, `RRI`, `SLN`, `SYD`, `TBILLEQ`, `TBILLPRICE`, `TBILLYIELD`, `XIRR`, `XNPV`, `YIELD`, `YIELDDISC`, `YIELDMAT`

- Lookup and reference: `ADDRESS`, `AREAS`, `CHOOSE`, `CHOOSECOLS`, `CHOOSEROWS`, `COLUMN`, `COLUMNS`, `DROP`, `EXPAND`, `FILTER`, `HLOOKUP`, `HSTACK`, `HYPERLINK`, `INDEX`, `LOOKUP`, `MATCH`, `ROW`, `ROWS`, `SORT`, `SORTBY`, `TAKE`, `TOCOL`, `TOROW`, `TRANSPOSE`, `UNIQUE`, `VLOOKUP`, `VSTACK`, `WRAPCOLS`, `WRAPROWS`, `XLOOKUP`, `XMATCH`

- Database: `DAVERAGE`, `DCOUNT`, `DCOUNTA`, `DGET`, `DMAX`, `DMIN`, `DPRODUCT`, `DSTDEV`, `DSTDEVP`, `DSUM`, `DVAR`, `DVARP`

- Engineering: `BESSELI`, `BESSELJ`, `BESSELK`, `BESSELY`, `BIN2DEC`, `BIN2HEX`, `BIN2OCT`, `BITAND`, `BITLSHIFT`, `BITOR`, `BITRSHIFT`, `BITXOR`, `COMPLEX`, `CONVERT`, `DEC2BIN`, `DEC2HEX`, `DEC2OCT`, `DELTA`, `ERF`, `ERF.PRECISE`, `ERFC`, `ERFC.PRECISE`, `GESTEP`, `HEX2BIN`, `HEX2DEC`, `HEX2OCT`, `IMABS`, `IMAGINARY`, `IMARGUMENT`, `IMCONJUGATE`, `IMCOS`, `IMCOSH`, `IMCOT`, `IMCSC`, `IMCSCH`, `IMDIV`, `IMEXP`, `IMLN`, `IMLOG10`, `IMLOG2`, `IMPOWER`, `IMPRODUCT`, `IMREAL`, `IMSEC`, `IMSECH`, `IMSIN`, `IMSINH`, `IMSQRT`, `IMSUB`, `IMSUM`, `IMTAN`, `OCT2BIN`, `OCT2DEC`, `OCT2HEX`

Status: 42,764 cases are checked against Excel, with Excel's answers stored in this repository. In en-US they are the coverage measurement (5,275), the case sets (24,884), the probes (959) and the random check (4,000). In ru-RU they are the coverage measurement (5,275) and the case sets (2,371): the locale sets and six en-US sets run again in ru-RU. Spilling is checked on 104 more books (`docs/spill.md`), what cells show on 12,949 cells, 9,819 in en-US and 3,130 in ru-RU (`docs/formats.md`), structure operations on 317 books (`docs/structure.md`), tables on 176 books in en-US and ru-RU (`docs/tables.md`), hidden rows, the 1904 date system and iterative calculation on 98 books and circular references without it on 27 books in Excel for Windows, 6 of them predicted before they ran (`bench/coverage/workbook/`, `docs/cycles.md`), data validation and outlines on 74 books in both Excels (`docs/validation.md`), and hidden sheets on 50 books in both Excels (`docs/xlsx.md`).

Every function in the list matches Excel en-US on all of its cases in the coverage measurement (`docs/coverage-formulajs.md`) and in the Excel case sets (`test/excel/cases/`), which also cover the functions formula.js lacks (`XLOOKUP`, `XMATCH`, `TEXTBEFORE`, `ADDRESS` and others), with two kinds of exceptions, both listed in `test/excel-cases.ts`: `XIRR` matches within 1e-6 relative (`SOFT`: Excel stops its iteration within 1e-8 of the root, the engine finds the root), and so do the t and F quantiles (`T.INV`, `T.INV.2T`, `TINV`, `F.INV`, `F.INV.RT`, `FINV`), where Excel loses up to 7 digits at 10^9 degrees of freedom and more, and a few cases differ knowingly (`KNOWN`: the engine's answer is pinned by the tests and the reason is in [Known differences](#limitations)); a difference the random check finds outside the functions (operators, the array layer) is pinned as open in `test/excel-open.ts` until it is fixed (none is open now). `TODAY`, `NOW`, `RAND` and `RANDBETWEEN` (volatile) have their own semantics and clock tests instead; `RANDARRAY` has Excel cases on what does not depend on the draw (sizes, bounds, whole numbers) and clock tests.

How the acceptance works: the measurement generated 5,275 cases for 402 Excel functions, typical calls (mostly the examples of Excel's function reference) and edge cases, and computed them in Excel 16.113 for Mac with the en-US locale and English UI, as dynamic array formulas the way Excel 365 saves them (no implicit intersection, an array answer spills); the answers are stored in `test/excel/coverage/excel-en.json`. `test/excel.test.ts` computes the cases of every whitelisted function with this engine, in one book with the measurement's data sheet, and compares: numbers within |a−b| ≤ 1e-9·max(1,|a|), text, booleans and error codes exactly. The test needs neither Excel nor LibreOffice. Rules that are not one function's (numbers as text, long number literals, criteria over whole columns) are checked the same way on Excel's answers to 959 probes in `test/excel/probes.json` (`test/excel-probes.test.ts`). Every case set and the probes are computed as dynamic array formulas too. The `core` set pins the rules of operators, arrays and references the random check pointed at, and `near-integer` how Excel reads a number near a whole number, on exact values from one unit in the last place to 10^-6 away. The random check (`fuzz`) generates formulas over the whitelist from a seed, nested calls, operators, ranges and arrays, number and date text and errors among the arguments, and compares them with Excel's answers the same way; the rules it found are pinned by `fuzz-rules`, and the differences it found outside the functions are listed in `test/excel-open.ts`. Arrays are checked by the `arrays` set: operators and constants, array results, aggregates and lookups over arrays, every whitelisted function's one-value arguments given a range and an array, and the functions with an array answer. The older cases of `test/semantics-cases.ts` (expected answers from Excel's documentation, cross-checked with LibreOffice) still run too.

Notes on semantics:

- Arguments are coerced by kind, the same way for every function of a group: numeric arguments read text as `VALUE` does (`"$1,000"`, `"50%"`, `"1 1/2"`, date and time text: `ABS("2025-01-31")` is 45688) and booleans, so do operators and `TEXT`, a range in `SUM`, `AVERAGE`, `MAX` and the like contributes only its numbers while a direct argument is coerced, date arguments read date text as Excel does in the book's locale (en-US: `2025-01-31`, `1/31/2025`, `31-Jan-2025`; `1/31` and `Jan 31` in the clock's year), date and time text after a space is no date or time there (`DAY(" 1/31/2025")` is `#VALUE!`, while `VALUE`, `DATEVALUE`, `TIMEVALUE` and `DAYS` drop the spaces), and `"TRUE"`/`"FALSE"` text (`"ИСТИНА"`/`"ЛОЖЬ"` in ru-RU) is a boolean where a boolean is expected; `AND`, `OR` and `XOR` take it given directly too, and skip other text.
- Where several arguments are wrong, the error is the one Excel finds first: arguments given as values are read before references and arrays (`LARGE(A1:A9,"x")` is `#VALUE!` even with an error in A1:A9; `AND`, `CONCAT` and `TEXTJOIN` answer an error given as a value before one inside a range), except `GCD`, `LCM` and `MULTINOMIAL`, which read every argument in order, values inside ranges and arrays too, and check signs last; `XIRR` finds a boolean or other text among its dates before a blank date; the former Analysis ToolPak functions check as Excel reads: the discount and most bond functions read the basis first, the coupon functions check the dates, the frequency and the basis in turn, the radix functions read their places first, `DELTA` and `GESTEP` their second number, `COMPLEX` its suffix; functions over two arrays (the paired statistics, `SUMX2MY2` and its kin, `CHISQ.TEST`, `T.TEST`) take a reference to one cell first, argument by argument, its error the answer and a blank cell `#VALUE!` (an array of one is an array), then compare sizes (`#N/A`), then look for errors inside; `SUMPRODUCT` takes each array's error, and a blank cell given alone, before it compares their shapes (`#VALUE!`).
- `LINEST`, `LOGEST`, `TREND` and `GROWTH` leave out a variable that depends linearly on the others or on the constant, with coefficient and standard error 0, as Excel does (Excel takes variables by their largest norm left; one y value takes a row of x as its variables); sizes are checked before values (`#REF!` before `#VALUE!`), and an error among `TREND`'s or `GROWTH`'s new x is `#VALUE!`. With a constant they take y and each variable about their means before the decomposition, so data far from 0 against their spread keep their digits (`LINEST` of 1E12, 1E12 + 0.01, 1E12 + 0.02 on 0, 1, 2 gives the slope of the stored values, 0.010009765625, as `SLOPE` does); `TREND` and `GROWTH` evaluate b + m·x from the coefficients, as Excel does.
- The complex-number functions read each part of complex text as `VALUE` reads a number in the book's locale, the first one with its own sign (`IMREAL("50%+i")` is 0.5, `IMREAL("-(1)+i")` 1, `IMREAL("1 000+2i")` `#NUM!` in en-US), and a real number from text as `VALUE` does when no sign follows its first character (`IMABS("50%")` is 0.5, `IMREAL("1/2/2025")` is 45659, `IMREAL("2025-01-31")` is `#NUM!`), as Excel does.
- Criteria (`SUMIF`, `SUMIFS`, `COUNTIFS`) compare text case-insensitively but accent-sensitively, support `*` and `?` (where `~` takes the next character as it is; text without `*` or `?` is compared as it is, `~` included), read numbers, dates and times inside criterion text as `VALUE` does (`">=1/1/2026"`, `"$1,000"`, `"50%"`), and a number criterion's `=` also matches text that reads as that number (`2` matches `"2"`, `0.5` matches `"50%"`) while its `<>` matches all text, as in Excel. An error, or its text in any case (`"#N/A"`), matches the cells holding that error; `<` and `>` order errors as `ERROR.TYPE` does. Their ranges must have the same referenced size (else `#VALUE!`); whole columns of sheets with different numbers of rows fit, and blank cells of a whole column count where every criterion matches a blank (`COUNTIFS(A:A,"")`).
- Numbers become text (`&`, `CONCAT`, text arguments) as in Excel: rounded to 15 significant digits, in positional notation when that takes at most 20 characters, the sign aside (`0.0000000745`, `10000000000000000000`), else in scientific notation with an exponent of two digits at least (`7.4500104300146E-06`, `1E+20`). `TEXT` with the `General` format has Excel's other rule, at most 11 characters (`0.333333333`, `1.23457E-05`).
- Number literals and numeric text keep 15 significant digits and drop the rest, as Excel reads them (`123456789012345678` is 123456789012345000); a result smaller than the smallest normal number (2.2250738585072E-308) is 0.
- A formula's last operation, outside parentheses, when it is `+` or `-`, is 0 where its answer is below eight units in the last place of its left operand, 2^(e-49) for a left operand between 2^e and 2^(e+1), as Excel takes such an answer for the error of decimal operands in binary: `=0.1+0.2-0.3`, `=0.5-0.4-0.1` and `=1E+15+0.3-1E+15` are 0 (the last is 0.25 in doubles), while `=(0.1+0.2-0.3)`, `=1*(0.1+0.2-0.3)`, `=0.1+0.2-0.3+0`, `=-(0.1+0.2-0.3)`, a function's argument and `LET`'s calculation keep 5.55E-17; over an array each element alike. `SUM` takes the rule at its last addition (`SUM(0.1,0.2,-0.3)` is 0, `SUM(0.1,0.2,-0.3,1E-20)` is 5.55E-17 + 1E-20), and with it `SUBTOTAL` and `AGGREGATE` and `AVERAGE`, whose mean is 0 where that addition is made 0 (a sum that plain addition leaves at 0, as `1E16`, `1` and `-1E16` give, keeps its stable mean: not measured); `SUMIF`, `SUMIFS`, `SUMPRODUCT`, `MMULT` and `NPV` keep the double. Excel for Windows answers alike. A defined name's formula, conditional formats and data validation were not measured: the engine takes the rule in a rule's formula as in a cell's, and not in a name's.
- A negative number to a fraction (`^`, `POWER`) is its real root where 1/y, to 15 significant digits, is an odd whole number below 2^31 - 1, computed as Excel computes it, -exp(y·ln(-x)): `(-32)^0.2` is -2, `(-8)^(1/3)` -2 to 15 digits and -1.9999999999999998 inside, as in Excel, `(-8)^0.333333333333333` a root too; `(-8)^(2/3)`, `(-8)^0.33333333333333` and `(-2)^(1/2147483647)` are `#NUM!`. Excel for Windows answers alike, but for the last digit of some roots.
- An operator converts its left operand before it looks at the right one: `"abc"-NA()` is `#VALUE!`, `NA()-"abc"` is `#N/A`; over arrays each element the same way, and a position outside a shorter array is `#N/A` whatever the other operand holds.
- Numbers compare as in Excel: `=`, `<` and the other comparison operators and criteria (`COUNTIF`, `SUMIFS` and the rest) take numbers equal to 15 significant digits as equal (`0.1+0.2=0.3` is TRUE), while lookups, `MATCH`, `XMATCH`, `UNIQUE`, `RANK`, `DELTA` and sorting compare exactly; `INT`, `TRUNC`, `ROUNDDOWN` and `ROUNDUP` read their number to 15 significant digits (`INT(1-2^-52)` is 1, and so is `INT(FTEST(A1:A8,A1:A8))`).
- A whole-number argument less than 2^-22 (about 2.4·10^-7) short of an integer is that integer (`LEFT("abc",0.9999999)` is `"a"`) in `LEFT`, `RIGHT`, `FIND`, `SEARCH`, `REPLACE`, `SUBSTITUTE`, `TEXTBEFORE`, `TEXTAFTER` (the instance), `FIXED`, `DOLLAR`, `ROUND` (the digits), `INDEX`, `ADDRESS`, `DATE`, `WEEKDAY` (the type), `FACT`, the column index of `VLOOKUP` and `HLOOKUP`, the field of the database functions, the function of `SUBTOTAL` and `AGGREGATE`, the tails and type of `T.TEST` and the tails of `TDIST`, the mode of `CEILING.MATH` and `FLOOR.MATH`, the month of `DB` and the format of `ARRAYTOTEXT` and `VALUETOTEXT`; the other whole-number arguments are truncated, as Excel truncates them (`MID`, `REPT`, `CHOOSE`, the digits of `ROUNDDOWN`, `ROUNDUP` and `TRUNC`, the counts of the statistical functions, `EDATE`'s months and the rest the `near-integer` set measured). `ISEVEN` and `ISODD` take a number less than 10^-10 short of an integer as that integer, `DATE` rounds a day below zero down, and `DAY`, `MONTH`, `YEAR`, `WEEKDAY`, `HOUR`, `MINUTE`, `SECOND` and a format's whole seconds read a date and time to the second as Excel does: half a second is added and the sum split into days and seconds in floating point, so 0.4 s before midnight is the next day, an exact half second at a date of today stays below (`SECOND(45688+1/256)`, 37.5 s, is 37) and one under a day goes up (`SECOND(10.5/86400)` is 11), and a moment that rounds past 9999-12-31 is `#NUM!`.
- Argument counts are those Excel loads: `SORTBY` takes each further by-array with its order (2, 3, 5… arguments; Excel refuses a book with `SORTBY` of 4), `SWITCH` at most 126 pairs and a default, a function newer than Excel 2007 (`CONCAT`, `TEXTJOIN`, `STDEV.S`, `XOR`, `HSTACK`, `AGGREGATE` and the rest stored with the `_xlfn.` prefix) at most 254 arguments (253 for `SORTBY`, `MAXIFS` and `MINIFS`), the others 255; any other count is the check code `syntax`.
- `IF`, `IFERROR`, `IFNA` and `SWITCH` do not compute the branch not taken, unless their condition (value, expression or a value to match) is an array: then every argument is computed and the function works element by element.
- An argument left out between commas (`IF(A1,,1)`, `SUM(1,,2)`) is taken as Excel 16.113 takes it, measured on every whitelisted function's arguments (the `omit` sets): as a blank, read by the argument's kind as 0, `""` or FALSE (`ROUND(1.5,)` is 2, `LEFT("abc",)` is `""`, `VLOOKUP(4,A1:C6,2,)` an exact match, `IF(TRUE,)&"x"` is `"x"`, `COUNT(1,,2)` is 3); as `#VALUE!` where the function takes an array or a reference (`CORREL(,A1:A8)`, `INDEX(,1)`); as not given, with the function's default, where the function treats it so (the newer functions' optional arguments: `SEQUENCE(3,,5)`, `XLOOKUP(4,A1:A6,B1:B6,)` is `#N/A`; `FIXED`'s decimals, `ADDRESS`'s style, the significance of `CEILING.MATH` and its kin, `PRODUCT`'s values; `TEXTJOIN`'s left-out ignore_empty is TRUE); in the former Analysis ToolPak's functions as not given, a required one as `#N/A` (`EDATE(,1)`, `ERF(1,)`). A formula that leaves out a reference Excel requires (a criteria range, `RANK`'s list, `SUBTOTAL`'s references) is the check code `syntax`, as Excel does not load it. The table of the functions that differ from the blank is `src/functions/omitted.ts`.
- `LET` binds names to values, each seen from the end of its value to the end of its `LET` and in any case, an inner `LET` shadowing an outer one and both shadowing a defined name of the same word; a name used in its own value or past its `LET` is a defined name (`LET(x,x+1,x)` is `#NAME?` without one). A value is computed once, on its first use (`LET(x,RAND(),x-x)` is 0; an unused value's error does not count), and a reference stays a reference (`LET(r,A1:A9,SUM(r))` skips text, `SUMIF(LET(r,A1:A9,r),">0")` works); a left-out value is a blank. Excel loads a `LET` that ends with a value (an even number of arguments: `LET(x,1,y,2)` is 2) and at most 254 arguments, not a name declared twice or a word that is no name (`LET(A1,1,A1)`: the check code `syntax`). The canonical text spells every use as its name is declared (`LET(Abc,1,ABC+1)` is `LET(Abc,1,Abc+1)`), as Excel shows it; in ru-RU `LET` keeps its name.
- Arrays follow Excel 365's dynamic arrays. An operator, or a function argument that takes one value, given an array or a range of several cells works element by element: the result has the larger number of rows and of columns, a single row or column is repeated, a position outside a shorter array is `#N/A` (`{1,2,3}+{1,2}` is `{2,4,#N/A}`). The functions of the former Analysis ToolPak (`EOMONTH`, `BESSELJ`, the radix and complex-number functions, the bond functions and others) and the field of the database functions do not take a range there: `#VALUE!`, while an array (`EOMONTH(A1:A2+0,0)`) works; `N` and `T` take the top-left cell of a range. Functions that take ranges take arrays the same way (`SUM` adds an array's numbers), except that `MAXA`, `STDEVA` and their kin take only an array's numbers and `AVERAGEA` counts its text as 0. An array of one value stays an array under an operator or a lifted call (`TYPE({1}+0)` is 64, so `SUM({TRUE}=TRUE)` skips the boolean and is 0), while a one-cell reference is a value (`TYPE(A1:A1+0)` is 1). A whole column is an array of 1,048,576 rows (`SUM(B:B+1)` adds 1,048,576), computed on its occupied rows.
- `INDEX`, `TAKE`, `DROP`, `XLOOKUP` (its `if_not_found` too), `CHOOSE`, `IF`, `IFS` and `SWITCH` answer references when given references, one cell too, so `SUMIF(INDEX(A1:C9,0,1),">2")` works and `SUM` skips text there (`SUM(IF(TRUE,A2))` is 0 when A2 holds the text "2"); `IFERROR`, `IFNA`, `FILTER`'s `if_empty` and unary `+` give values. `HSTACK` and `VSTACK` answer an error given as a value (`VSTACK(1,1/0)` is `#DIV/0!`) but stack a reference to a cell holding an error; the other functions with an array answer give arrays, where a blank cell is an empty value (`COUNTA` counts it, `ISBLANK` is TRUE). `TYPE` of an array, or of a range of several cells, is 64. `IFS` answers the `#VALUE!` of a condition that is text as an array of that one error, so `ROWS(IFS("x",1,TRUE,2))` is 1, as in Excel, while an error given as a condition is the answer itself (`ROWS(IFS(1/0,1))` is `#DIV/0!`).
- Sorting (`SORT`, `SORTBY`) is stable and orders numbers, text, booleans, errors, then blanks (blanks last in a descending sort too); text compares as `<` does. `SORTBY` with an array or a range of several cells as a sort order answers as Excel does, an array as large as the orders, lifted as an operator lifts them (an order of one row or column repeats, past a shorter one `#N/A`): at each position, in rows, the error of the call with the orders there (an error given as an argument, `#VALUE!` for an order that is not 1 or −1 or has no value, or for a by-array of another size), the first value of the first call that sorts, and after it blanks, or `#VALUE!` when it has more than one by-array (`SORTBY(A1:A6,B1:B6,{1;-1})` is the first value and a blank, `SORTBY(A1:A6,B1:B6,{2;-1})` `#VALUE!` and the first value of the descending sort). A by-array of one value counts as one row and one column: the first by-array decides whether rows or columns are sorted and every later one must fit it, so one value fits an array of one column or one row (`SORTBY(A1:A6,5)` keeps the order) and beside a by-array of the array's rows is `#VALUE!` (`SORTBY(A1:A6,B1:B6,1,5,1)`).
- `SUBTOTAL` and `AGGREGATE` (options 0–3) skip the cells whose formula calls either of them anywhere (`=SUBTOTAL(9,A1:A2)*10` too, a plain reference to such a cell not), so nothing is counted twice. `SUBTOTAL` 101–111 and `AGGREGATE` options 1, 3, 5 and 7 skip hidden rows, and `SUBTOTAL` 1–11 skips the hidden rows of a filtered sheet ([Hidden rows and filters](#hidden-rows-and-filters)). `AGGREGATE` 14–19 take an array as well as a reference (`AGGREGATE(14,6,10/A1:A9,1)`).
- Lookups compare within a type (numbers, text, booleans), text in any case, and never match blanks or errors in the data. An exact match in `VLOOKUP`, `HLOOKUP` and `MATCH` (and `XLOOKUP`, `XMATCH` in match mode 2) takes `*`, `?` and `~`, where `~` escapes any character. An approximate match is Excel's binary search over the referenced cells, so it gives Excel's answer on unsorted, mixed and whole-column data too (checked on about 200 such cases); `MATCH` with type −1 scans. A blank lookup value is 0 in `VLOOKUP`, `HLOOKUP`, `MATCH` and `LOOKUP` and finds a blank cell in `XLOOKUP` and `XMATCH`. `LOOKUP` reads on past the end of a result reference shorter than its lookup array, as Excel does, whatever answers the reference (a spill, `IF`, `INDEX`, `TAKE`): a column down, a row or a single cell to the right (`LOOKUP(4.19,L1:L6,B2:B3)` and `LOOKUP(4.19,L1:L6,IF(TRUE,B2:B3))` are B4 and `LOOKUP(4.19,L1:L6,B2)` D2 when 4.19 falls between L3 and L4), a blank past the sheet's end; the cells it may read count among the formula's references, a column's below it and a row's or a cell's to its right as far as the lookup array is long (both ways for a part of a rectangle `INDEX` or `TAKE` may answer), past a spill in its array's direction. An error in a column index or match type taken from a cell is `#REF!`, as in Excel. An array or a range of several cells as the column (row) index of `VLOOKUP` (`HLOOKUP`) answers the found row (column) at each number, shaped as the index; the first number in order that is not valid answers for all: `#VALUE!` below 1, `#REF!` past the table, for an error or for text that is not a number. A call lifted by an array key or match mode reads the index's top-left number in each call (`INDEX(VLOOKUP({5;8},D!L1:N6,{2;4},FALSE),2)` is `"eight"`). A range as `MATCH`'s type gives its top-left cell. A table of `VLOOKUP`, `HLOOKUP` or `MATCH` given as one value is read as a number and finds nothing (`MATCH(5,5,0)` is `#N/A`, `MATCH(5,"x",0)` `#VALUE!`); one value as `XLOOKUP`'s lookup array looks along the return array's one row, else its one column (`XLOOKUP(0,0,A1:A6)` is A1:A6).
- Database functions (`DSUM` and the rest) read records by rows under a header row; a database without records is `#VALUE!`, and so is a field taken from a cell holding an error (an error given as the field is itself the answer). Criteria are written as in `SUMIFS`, except that text without an operator matches the values that begin with it (`App` selects `Apple`); criteria rows are alternatives, conditions in a row all hold. A condition under a header that is not a field name holds for every record when it is `TRUE` (Excel's computed criterion) and for none otherwise.
- Text longer than 32,767 characters: `&` and `CONCATENATE` cut it there, `CONCAT` and `TEXTJOIN` answer `#CALC!`, the other text functions `#VALUE!`, as Excel does. An empty array (`FILTER` that keeps nothing, `TAKE` of zero rows, `DROP` of everything) is `#CALC!` too.
- `VALUE` reads currency, thousands groups, parentheses, `%`, exponents, `1 1/2`, dates and times as Excel en-US does, times past 24 hours too (`25:00`), date text without a year in the clock's year (`1/31`; `2/29` in a year that is not leap is February 2029, as in Excel); time text follows one rule in `VALUE`, operators and date arguments (`6 PM` and `6 P` but not `6PM`, `6:30.5` as minutes and seconds); `NUMBERVALUE` removes whitespace and the group separators, then reads the rest the same way.
- `CHAR` and `CODE` work with the Windows code page of the book's locale, Windows-1252 in en-US and Windows-1251 in ru-RU, so a book gives the same characters as in Excel for Windows, where most books are made; Excel for Mac answers codes 128–255 by Mac Roman and Mac Cyrillic, a platform's own encoding rather than the book's. Codes 1–127, a code outside 1–255 (`#VALUE!`), the code of a character outside the page (95, `_`) and the arguments are as Excel computes them (the `char-code` set, in both locales); the tables are checked against Unicode's published ones. `UNICHAR` and `UNICODE` work with code points.
- `ACCRINT` counts each whole coupon period before the settlement as 1 and a part of one by its days over the period's length; on basis 1 that length is the period ending at the first interest date for every later period, and a period ending on the settlement counts as a part (`ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2025,1,1),0.05,1000,2,1)` is 25·(1 + 184/182)). A coupon period that begins before 1900 leaves no coupons: `ACCRINT` is 0 when the issue's period does, `PRICE` `#NUM!` when the settlement's does, as `COUPNUM` finds none. Serial 0 as `ACCRINT`'s issue or first interest date and any date of 0 in `XNPV` are `#NUM!`, while `XIRR` takes it. `AMORDEGRC` and `AMORLINC` take no more than cost less salvage in the first period, and `AMORDEGRC`'s period that passes the salvage takes its whole amount, the later ones nothing.
- `WORKDAY`, `WORKDAY.INTL`, `NETWORKDAYS` and `NETWORKDAYS.INTL` read date and numeric text among their holidays, in cells and arrays too (`WORKDAY(A1,5,H1:H3)` with `"1/6/2025"` in H1); other text and booleans there are `#VALUE!`.
- Dates follow Excel's 1900 system, phantom 1900-02-29 included: `DATE(1900,2,29)` is 60. A book can be in the 1904 system instead ([Date systems](#date-systems)), where the former Analysis ToolPak's functions read date text given as a value as Excel does there: `EDATE("1/31/2025",1)` is 42793, while `EDATE(A1,1)` with that text in A1 is 44254. `DATE` takes a day outside −32768…32767 as 32767 and answers `#NUM!` for a month of 32767 or more, as Excel does.

## Limitations

- **Spilling.** Of two arrays over one cell that do not depend on each other, the engine always lets the lower one spill; Excel's choice varies between runs (`docs/spill.md`). An array whose size depends on a volatile draw spills by the engine's clock, where Excel gives `#SPILL!` when the size changes between its calculation passes (for most of ten such arrays in one book, for none of one alone). A cycle through a spill range stops its cells as any cycle does ([Circular references](#circular-references)). There are no `LAMBDA` (nor `MAP`, `REDUCE` and the other functions over it).
- **Content Security Policy.** Only `execution: 'eval'` uses `new Function`: in a browser with a CSP that does not allow `'unsafe-eval'`, `compute` and `createSession` in that mode throw an `EvalError` saying so. The default `closures` mode generates no code and computes under such a policy.
- **Known differences from Excel:**
  - `XLOOKUP` and `XMATCH` with binary search (search mode 2 or −2) scan linearly: the same answer on sorted data, where Excel's binary search is defined. Match mode 3 (regular expressions) is `#VALUE!`.
  - Text order for `<`, `>` and sorting uses `Intl.Collator` with the en-US order in both locales (case-insensitive, accent-sensitive, Latin before Cyrillic, as Excel for Windows sorts), which may differ from Excel for some characters; without `Intl` it falls back to comparing lower-cased code units.
  - Excel for Mac puts Cyrillic before Latin in ru-RU (`"a"<"а"` is FALSE there); the engine puts Latin first, as Excel for Windows does in every locale, in comparisons, criteria, lookups and SORT.
  - Excel for Mac writes month names in ru-RU in lower case (`TEXT` with `ММММ`, `МММММ` or `[$-419]mmmm`, and cells: `январь`, `я`); the engine writes them with a capital, as Excel for Windows does (`Январь`, `Я`).
  - Excel for Mac writes the short month names of ru-RU (`МММ`) `февр`, `март`, `июнь`, `июль`, `сент`, `нояб`; the engine takes three letters, as Excel for Windows does (`фев`, `мар`, `июн`, `июл`, `сен`, `ноя`).
  - A second decimal separator among the decimals of a number format (`#,##0,00` in ru-RU) is a literal between digits; Excel shows one digit fewer before it in some numbers (`TEXT(-0.5,"#,##0,00")` is `-,50,00` in Excel and `-,500,00` here).
  - `CHAR` and `CODE` take codes 128–255 from the Windows code page of the book's locale (Windows-1252 in en-US, Windows-1251 in ru-RU), as Excel for Windows does; Excel for Mac takes Mac Roman and Mac Cyrillic (`CHAR(200)` is `È` here and `»` in Excel for Mac en-US, `CODE("Ж")` in ru-RU 198 here and 134 there).
  - `CODE` of a character outside the code page of the book's locale is 63 (`?`), as Excel for Windows answers; Excel for Mac answers 95 (`_`): `CODE("中")`, and `CODE("Ж")` in en-US.
  - `RATE` with a rate of exactly 0 as the answer gives 0 or a value within 1e-10 of it; Excel answers its own rounding noise up to about 2e-9 or, from some guesses, `#NUM!`.
  - `GAMMA.INV`, `CHISQ.INV`, `CHISQ.INV.RT` (and `GAMMAINV`, `CHIINV`) give the quantile where Excel's iteration fails with `#NUM!`: shapes from about 5·10^9 (10^10 degrees of freedom) and probabilities near 1 with shapes under 0.03; Excel's own `GAMMA.DIST` at the engine's answer gives the probability back.
  - `BINOM.INV` where alpha equals the probability of no success, (1 − p)^n, exactly: Excel decides by its own rounding of that power, the engine by exp(n·ln(1 − p)); they part on `BINOM.INV(3,0.25,0.421875)`, 0 in Excel and 1 here.
  - A number that Excel computes by another algorithm can differ from the engine's in the 15th significant digit, within the tests' relative tolerance of 1e-9; turned into text (`&`, `TEXTJOIN`) the last digit then differs (`TEXTJOIN(";",FALSE,BESSELK(10,3))` is `2.72527006358477E-05` in Excel and `2.72527006358476E-05` here).
  - `LINEST` with a constant x variable besides the intercept: both give it coefficient 0 and the same fit, but Excel's statistics count it as a variable in some way the engine does not reproduce (`LINEST({1;2;3;5},{1,1,2;2,1,4;3,1,6;4,1,8},TRUE,TRUE)`: degrees of freedom 1 in Excel, 2 here, and the F statistic and the intercept's standard error differ).
  - `FTEST` of a sample with itself is 1 − 3·2^-53 in Excel and 1 − 2^-52 here: the last bits of a number Excel computes by another algorithm can differ, which shows only when the difference is magnified (`(1-FTEST(A1:A8,A1:A8))*2^52` is 1.5 in Excel, 1 here).
  - A LET name that reads as R1C1's row or column (`R2`, `C5`, `RC5`) or a boolean (`TRUE`): an xlsx file writes LET's names with the prefix `_xlpm.`, so Excel loads `LET(_xlpm.TRUE,1,_xlpm.TRUE)` (1); formula text has no prefix, and there Excel refuses such a word as LET's name, as the engine does (`LET(R2,1,R2)` typed), while it takes a cell's word otherwise (`LET(A1,1,A1)` is 1).
  - A spilled array that an edit brings into a circular reference's way holds its first value here and spills no more, as `held` is one value a cell; Excel for Windows keeps its whole spill as it was (the workbook set `cycles-held`, `docs/cycles.md`). Excel for Mac shows a warning window on opening a book with a circular reference, so cycles without iterative calculation were measured in Excel for Windows alone.
  - Iterative calculation starts from 0 at every computation. Excel starts a book it opens from the values the file stores, empty text where it stores none, and goes on from the values it holds at every recalculation and edit. A cycle that converges differs from Excel's within the maximum change, and one that does not, such as a counter, keeps counting in Excel.
  - Excel's AutoCorrect extends a table when a value is typed right below or right of it and fills a calculated column when a formula is typed into an empty column of a table; `setCells` changes the cells typed alone ([docs/tables.md](docs/tables.md)). Formulas pasted into part of a calculated column fill the whole column in Excel too; `copyRange` pastes the cells alone.
  - A table renamed to a name Excel does not take, or to a defined name: Excel's object model makes the name one it takes (`AB12_`, `_1Ставка`, `_?Продажи`, `Rate_1`, `СТАВКА_1`), as its name box refuses it; `renameTable` refuses with `table_name`.
  - Excel ru-RU takes TRUE and FALSE as defined names, and Excel for Windows ru-RU as a table's name; the engine keeps formulas in en-US text, where they are the booleans, and refuses them in every locale. Excel for Mac ru-RU keeps a table named `ИСТИНА` or `ЛОЖЬ`, which Excel for Windows and the engine refuse, as both Excels refuse such a defined name.
  - A typed formula longer than 255 characters (one over a name of 255 letters) is not measured in Excel for Mac, whose AppleScript cuts what it sets to 255 characters; Excel for Windows computes it as the engine does.
  - Text typed into a cell is read by Excel as a number, a date or a boolean where it can be (`2025-01-31` in ru-RU is a date); `setCells` takes a value as given, so a header cell typed so names its column `2025-01-31` here and `31.01.2025` in Excel ru-RU. `setCellText` types text as Excel reads it, and `parseEntry` gives the value Excel reads without a session.
  - `parseEntry` reads English function names and TRUE in a formula typed in ru-RU as `toInvariant` does, as the functions and the boolean, where Excel ru-RU takes them as unknown names (`=TODAY()` computes `#ИМЯ?` there).
  - `parseEntry` has no book: a typed formula that refers to a cell with a date format (`=C2+1`) gets that format in Excel and none here; a typed `#GETTING_DATA` is an error in Excel and text here; an unknown name typed in another case than before in the book (`=Foo*2` after `=foo`) is written in its first case in Excel, as typed here; typed on the keyboard, `+5%+1` takes a percent format and `+-+5` is the formula `+-5` in Excel ([docs/entry.md](docs/entry.md#known-differences)).
  - A cell's word used as LET's name and again after the LET in one formula: Excel reads the later one as a defined name, so `=LET(IW1,5,IW1)+IW1` is `#NAME?` there and 5 here, where it is the cell.
  - Excel's formula bar keeps the apostrophe a person typed before text that needed none (`'abc`, `'TRUE` in ru-RU); a book holds no such mark, and `entryText` puts the apostrophe only before text that would read as something else. Excel for Mac's bar writes a ru-RU time with a leading zero (`01:30:00`); `entryText` writes it without, by Excel for Windows's regional data ([docs/entry.md](docs/entry.md#known-differences)).
  - `parseClipboard` reads quoted controls, but plain text can give two grids the same encoding; exact roundtrip needs another clipboard format or separately carried rows ([Copying plain text](docs/entry.md#copying-plain-text)).
  - Excel's Fill Right and fill handle move the column of a structured reference (`SUM(Sales[Qty])` filled right is `SUM(Sales[Price])`); `copyRange`, as copy and paste, keeps it.
  - Excel, formatting a row across a formatted column (or a column across a row), gives the cells where they cross both lines' formats merged field by field: a bold row across a filled column makes them bold and filled. A style of the book is one whole, and `setLayout` sets the line's alone, so the book shows such a cell the row's style, bold without the fill ([docs/xlsx.md](docs/xlsx.md#lines-excel-formats)).
  - `VAR`, `STDEV`, `DEVSQ`, `CORREL`, `RSQ`, `SLOPE`, `STEYX` and their kin, the database ones too, are computed here from exact sums; Excel computes them in double precision about a mean it sums plainly, and where the spread is small against the mean, or the residual against the variation, it loses the digits of the answer: of 10,000 values 1E12 + 0, 0.01, 0.02 `VAR.P` is 1.598E-04 in Excel and 6.680E-05 here, as exactly, `CORREL` of them with their steps 0, 1, 2 is 0.647 in Excel and 1 here, and `STEYX` of 1, 2.00000001, 3 on 1, 2, 3 is 0 in Excel and 8.16E-09 here.
  - Where the squares or products of the values leave the range of doubles, Excel answers what is left of them and the engine the exact answer: the squares of 1E-200 and 2E-200 are 0 in Excel, so `STDEV.P` of them is 0 (5E-201 here, within the tests' tolerance) and `CORREL` with 1, 2 and `SLOPE` of 1, 2 on them `#DIV/0!` (1 and 1E+200 here); a sum or a product past 1.8E+308 is `#NUM!` in Excel where the answer fits (`AVERAGE` of 1E+308 three times and its variance 0, the covariance of 1E+308, −1E+308, 1E-200 with 0, 0, 1E+308, 2.2E+107 here).
  - `LINEST` and `TREND` over y far from 0 against its spread give Excel's coefficients, but Excel's intercept and regression sum of squares follow the mean it sums plainly: of 10,000 values 1E12 + 0, 0.01, 0.02 on their steps 0, 1, 2 the slope is 0.0100098 in both (0.010009765625 exactly), the intercept 1E12 − 0.0096 in Excel and 1E12 here, the regression sum of squares 1.598 in Excel and 0.668 here (as exactly, `DEVSQ` of the values), so `TREND` at step 3 is 1E12 + 0.0204 in Excel and 1E12 + 0.0300 here.
  - `GROWTH` far from 0 in x is `#NUM!` in Excel where the fitted value fits a double: of 100·1.1^k to cents on x = 1E10 + k, k = 0…11, the value at 1E10 + 12 is 313.84 here, where b·m^x has `LOGEST`'s b 0 and m^x past the largest double.
  - `LOOKUP` counts among its references every cell it may read on to, as far as its lookup array is long, whatever the value looked up: a formula among them, the `LOOKUP` itself too, comes first, and the `LOOKUP` in that zone is a cycle here, not computed (it holds its value, 0 in a book computed afresh), where Excel, which follows the cells `LOOKUP` reads in the end, computes it (`E10` = `LOOKUP(5,A1:A6,B10:D10)` reads `F10` in Excel and holds 0 here, as `E10` lies in `B10:G10`; so beside a spill that could not spill, and right of a rectangle `INDEX` takes a column from).
  - A cell's shown text is written from its value as the result gives it, to 15 significant digits, where Excel writes a date and a time from the whole number: a moment a few millionths of a second from half a second can show another second or day (`45688.99999421295` shows `31 23:59:59` in Excel with `d h:mm:ss` and `1 0:00:00` here); `TEXT` and `displayValue` of the value write Excel's text.
  - Excel for Mac reads the long s (`ſ`) as `s`: `falſe` is FALSE in a formula, typed and as text (`IF("falſe",1,2)` is 2), `=ſum(1,2)` typed is `SUM`, and `ſ` is the name `S` (a LET name, a defined name, a table, a sheet); the engine follows Excel for Windows, where `ſ` is a letter of its own (`falſe` a name, `#NAME?`, and the text `"falſe"` no boolean).
  - Excel compares sheet names letter by letter, so it takes a sheet named `ß` beside `SS` (`ﬁ` beside `FI`), and reads a typed function's name with the dotted `İ` as `I` (`=İF(TRUE,1,2)` is `IF`; Excel for Windows the dotless `ı` too); the engine compares sheets as defined names and tables, where `ß` is `SS`, and refuses the second sheet, and reads a function's name by its letters, as Excel reads a file's (`#NAME?`).
  - `findCells` by Excel's Find finds nothing for an empty query or `~` alone, where Excel's Find finds empty cells, searches a range of one cell alone, where `Range.Find` searches the whole sheet from after it, and by columns on a filtered sheet gives the first column of the first block once round, where Excel's `FindNext` going round over the whole sheet passes it ([docs/live-books.md](docs/live-books.md#known-differences)).
  - `replaceInCells` keeps a text that reads as a number a text, as Excel keeps one typed after an apostrophe, where Excel's Replace makes a number of one without the prefix character (a file's); the book keeps no mark of an apostrophe. It replaces in a range of one cell alone, where `Range.Replace` replaces in the whole sheet ([docs/live-books.md](docs/live-books.md#known-differences)).
  - Outlines: Excel reads the collapsed marks again only for the groups whose rows change, so a summary row shown by hand over a group it was hidden with stays without a mark, where the engine marks every summary row of a group all hidden, as Excel reads a file; and columns grouped in a span of several columns reaching XFD, which Excel's level buttons leave and its Ungroup refuses, are taken as other columns ([docs/outline.md](docs/outline.md#known-differences)).
  - `WEEKNUM` with an invalid return type is `#NUM!`, as Excel answers it alone; inside one generated book of the random check Excel 16.113 gave a number for `DAY(WEEKNUM(40977,1E+3)<>2.5)` (1, on two runs).
- **Shown text.** Widths are those of Excel for Mac with the default font, Calibri 11; a cell's own font, its indent, wrapping and rotation are not taken into account, characters outside Latin, Cyrillic, digits and common signs count as a digit, and Excel for Windows lays text out at other widths, so near the edge of a column it can show `#` where the engine shows the number or the other way round. AutoFit and the widths typing gives are Excel for Mac's: Excel for Windows fits columns to its pixels, to widths that differ from the Mac's in most columns; the Mac shows its own short date (`3/15/26`) and ru-RU month names (`март`) where the engine follows Windows; text outside ASCII is fitted by the advance widths of Carlito, a font metrically compatible with Calibri, a point off Excel's on some long runs. Locale tags other than `[$-409]`, `[$-419]`, `[$-F800]` and `[$-F400]`, calendars and numeral systems (`[$-x-sysdate]`, `[DBNum1]`, `B2`) are not read. [docs/formats.md](docs/formats.md) has the details.
- **The ribbon's Number Format box.** `numberFormatCategory` and `numberFormatChoices` follow Excel 16.113 for Mac in en-US and Excel 16.0 for Windows in ru-RU. Excel for Mac in ru-RU names nearly every code `(все форматы)`, as it compares the code, kept in en-US syntax, with its formats written in Russian syntax: in 100 of its 368 cells it shows `(все форматы)` where the function, as Excel for Windows, names a category. Excel for Windows names a date or time with `;@` by the lists of its Format Cells window (in ru-RU `[$-419]d mmm;@` is Дата, `d mmm;@` none, `dd/mm/yy h:mm;@` Время); the function names every such code by its letters, as those codes come from such lists. Excel for Mac names the category of a format its list applied by the list's state as well, and its Short Date, Long Date and Time items were not the same in two sessions of one day; `numberFormatChoices` gives the first's (built-in 14, the system's long date and long time). The code `m/d/yyyy` is the short date in en-US here, as built-in 14 shows by it (Excel for Windows's code); Excel for Mac, whose short date is `m/d/yy`, calls it Custom. Excel for Windows in en-US was not measured.
- **Syntax.** Stored formulas use English function names and separators (no `;` or decimal commas; text typed in another language goes through `toInvariant`), no R1C1 references. A reference to another workbook needs its link in `externalBooks` (a new link is not made); a reference across another workbook's sheets (`[1]Jan:Mar!A1`) and a spill reference into one are refused. Functions outside the whitelist (`INDIRECT`, `OFFSET`, `LAMBDA`, …) give `#NAME?`.
- **Row heights.** Heights are Excel for Mac's for Calibri; a cell in another font is fitted as Calibri of its size (Arial 16 gets 21 points here and 20 in Excel), and Excel for Windows fits rows to its pixels (a line of Calibri 11 15 points, wrapped or not). Excel keeps every row whose height nobody set at its fitted height, on opening a file and after any edit; the engine fits a row on AutoFit and as text that wraps is typed into it, not as a file opens or a font changes. A formula typed or pasted over lines keeps its line breaks in Excel's text, not in the engine's canonical text ([docs/formats.md](docs/formats.md#fitting-rows)).
- **Links to other workbooks.** A closed workbook's reference is shown with the path its link holds (`'../data/[Book3.xlsx]Sheet1'!A1`), where Excel shows the absolute path it resolves from the folder of the opened file. Only entering or pasting a formula with a link marks the link (`refreshError`), as measured; whether moving one or undoing does was not. On a marked sheet a whole column or row is read to one row and column past the stored values, so a count over it (`COUNTA`) counts one `#REF!` where Excel counts every cell the values lack. [docs/links.md](docs/links.md#known-differences) has the rest.
- **Structure operations and styles.** Whole rows or columns copied onto part of a merged area are refused, where Excel pastes them. Excel's fill handle fills the hidden rows of a filtered sheet, while `copyRange` fills only the visible ones, as Fill Down does. Styles, notes, hyperlinks, frozen panes, rules of data validation and outlines are kept and moved, not computed: `compute` does not return them, `Session.getCells` and `Session.validation` do. A hyperlink's place is not rewritten when its sheet is renamed or rows move, as Excel leaves it.
- **Output precision.** Values in the result and in the hash are rounded to 15 significant digits, as Excel shows them; inside, computation keeps full double precision.
- **Conditional formats.** What they show is given by `Session.getCells` and `getViewport` alone: edits do not list the cells whose appearance changed, they give the rules' cells as `repaint`, for a host to read its window again where it meets them. A data bar's length and an icon follow Excel's documented rules, which no object model reports, so they were not measured; borders, underline and strike of a rule's style are kept for a host to draw. After rows are deleted, Excel put one rule whose reference the deletion lost after the others, where the engine keeps its place; Excel drops an Excel 2010 data bar's lengths other than 0 and 100 ([docs/conditional-formats.md](docs/conditional-formats.md)).
- **Data validation.** What a dropdown shows and what picking an item enters were not measured, as no object model gives them; a list's items are checked by which values Excel takes. A list item holding a comma, typed in ru-RU, is read split at the comma. [docs/validation.md](docs/validation.md#known-differences) has the rest.
- **xlsx.** A file's pictures, charts and pivot tables are not read, and each is reported in `issues`. A thread of comments is read and written as a note, without its dates and state; split panes that are not frozen are not read. Of a filter the book keeps the filter mode alone: an autofilter's range and criteria are reported, and a filtered sheet is written with criteria that every row meets. Formulas of the old kind are read as the dynamic formulas Excel 365 shows for them, with `@` where it takes one value of a range; the intersection operator is not read. Every formula is written as a dynamic array formula, which Excel 2019 and older show as an array formula. The East Asian date formats 50 to 58 are read as the code Excel for Mac en-US shows for them, `m/d/yy`, where Excel for Windows shows its short date in each locale. A gradient fill is read as no fill, and so are a conditional format's fill that takes the cell's own fill away (`patternType="none"`) and its solid fill without a colour; each is reported. [docs/xlsx.md](docs/xlsx.md) has the full list.

## Development

```sh
npm ci
npm run typecheck
npm test                # build, then all tests against src
npm run test:dist       # build, then all tests against the built dist
npm run test:oracle     # recalculate the LibreOffice answers with soffice and check them
npm run check:package   # build, then publint and attw on the packed tarball
npm run pack:release    # build, then the tarball in release/
FORMULA_BENCH=1 npx vitest run test/budgets.test.ts   # budgets on the full synthetic book
```

The tests cover the whitelist against Excel's answers on the coverage measurement's cases (`test/excel.test.ts`) and on the probes (`test/excel-probes.test.ts`), the argument layer rule by rule, the semantics cases against Excel's documentation and the stored LibreOffice answers (`test/oracle/libreoffice.json`, refreshed by `npm run test:oracle`, which needs LibreOffice's `soffice` on `PATH`; the usual tests need no office suite); determinism (shuffled input, a child process, three time zones, cents over 3,400 amounts); references, errors, volatile functions, parsing and the canonical form; sessions (200 random edits against full computations) and `explain`; structure operations against Excel's answers in `test/excel/struct/`, on the synthetic book and on seeded books; budgets; SHA-256 against test vectors and `node:crypto`; a browser bundle run in an empty JavaScript context and with code generation disabled; the two execution modes against each other on every case book and over 200 session edits, the closures mode with code generation disabled; computation in portions and in threads against the synchronous engine, threads lost and sessions restored, the pool on Web Workers in headless Chromium, Firefox and WebKit (through `playwright-core`; each skipped where its engine is not found, `CHROME_PATH` names a Chromium); the published package (exports, tarball contents, a consumer importing it by name). The code blocks of this README run as a test too.

Excel's answers that the tests compare with are stored in `test/excel/`, so the tests need no Excel: the coverage measurement's cases and Excel columns in `coverage/`, the case sets in `cases/` (`cases/ru-RU/` for the sets computed in ru-RU), the probes in `probes.json`, structure operations in `struct/`, spilling in `spill/`, Excel's ru-RU formula text in `locale/`, the files Excel saved for the xlsx subpath in `xlsx/`. They are recomputed only with Excel for Mac, by a step of `npx tsx bench/coverage/run.ts`: `excel en` and `excel ru` for the measurement, `probes` for the probes, `probes <set>` and `probes <set> ru` for a case set, `struct <set>`, `spill <set>`, `xlsx <step>`, `entry <set>` for typed text, `cond <set>` for conditional formats; `npx tsx bench/locale/snapshot.ts ru` takes Excel's ru-RU formula text. A run keeps the stored answers of unchanged cases and computes the new and changed ones. `bench/coverage/README.md` has the layout, the steps and the rules for driving Excel. Excel for Windows's answers to the same books are in `test/excel/win/`, computed by `npx tsx bench/windows/run.ts` on a Windows machine over SSH (`bench/windows/README.md`).

Functions live in `src/functions/<category>.ts` (`math`, `statistical`, `logical`, `information`, `text`, `date`, `financial`, `lookup`, `engineering`, `database`; old names such as `NORMDIST` sit next to their modern versions), each file with its own registry, a large category assembled from `src/functions/<category>/`; `src/functions.ts` assembles them. A function declares the kinds of its arguments (`args: 'n i?'`, see `src/args.ts`), and the shared argument layer coerces them, so an implementation gets numbers, text, lists or grids and never repeats the coercion rules. Ported code keeps its attribution in the file header. A function joins its registry when `test/excel.test.ts` passes for it; the coverage measurement (`bench/coverage`, `docs/coverage-formulajs.md`) lists the cases and the classes of the remaining formula.js functions.

## License

MIT. Function implementations are ported from [formula.js](https://github.com/formulajs/formulajs) (MIT, Copyright (c) 2014 Sutoiku, Inc.), [jStat](https://github.com/jstat/jstat) (MIT, Copyright (c) 2013 jStat) and [bessel](https://github.com/SheetJS/bessel) (Apache-2.0, Copyright (C) 2013-present SheetJS); [THIRD_PARTY.md](THIRD_PARTY.md) has their license texts and the files they are in.
