# xlcalc

A formula engine for spreadsheet-like workbooks in JavaScript. You give it a book (sheets of cells with values and Excel formulas, plus named ranges); it gives back the computed values, the errors with their reasons, and a deterministic hash of the values. Formulas use Excel syntax: English function names, commas between arguments, a decimal point; [Localization](docs/locale.md) shows them and reads them in another language, as Excel does.

What it does differently:

- **Compiles each formula once.** By default a formula becomes a tree of JavaScript closures, with no code generated, so a strict Content Security Policy does not stand in the way; with `execution: 'eval'` it becomes JavaScript text, and one `new Function` call compiles the book. A formula that exceeds the JS compiler's nesting limit falls back to closures, preserving evaluation order. Both modes give the same values and hashes, and the tests compute every case book both ways. Cells are evaluated over the reference graph in topological order. A session keeps the compiled book and recomputes only the cells that depend on an edit. On a synthetic book with 3,400 data rows and 3,907 formulas, mostly SUMIFS over whole columns, a cold compute takes about 0.3 s with the cells' shown text (0.28 s without) and an edit of a criterion cell under 10 ms (p95) in either mode on an Apple M5 Pro with Node 26; `test/budgets.test.ts` enforces 1.5 s and 100 ms.
- **Deterministic values and hashes.** The same book with the same clock gives the same values and the same `hash` (SHA-256 of a canonical text of all values plus the engine configuration version), regardless of input order, process or time zone. `inputHash` identifies the input.
- **A clock for volatile functions.** `TODAY` and `NOW` come from `clock.now` (UTC), `RAND` and `RANDBETWEEN` from `clock.seed` and the cell address. Without a clock the engine takes the current moment and a random seed, and `result.clock` reproduces the result.
- **A whitelist of functions, checked against Microsoft Excel.** A function is admitted when it gives Excel's answers on its test cases, typical calls and edge cases (text and booleans in numeric arguments, blanks, errors, date text, ranges with mixed content), computed in Excel 16: 43,778 cases, 36,132 of them in the en-US locale and 7,646 in ru-RU, with Excel's answers stored in this repository ([Functions](docs/functions.md) has the list and the counts). Where an admitted function still answers otherwise than Excel, the difference is listed in [Limitations](docs/limitations.md). Inputs outside these cases can differ too. Any other function is reported as `unknown_function` (`#NAME?`) rather than computed with different semantics.
- **Arrays inside a formula.** Operators and functions work element by element over ranges and arrays with Excel 365's rules (`SUM(B:B*C:C)`, `SUMPRODUCT((A1:A9="x")*B1:B9)`, `SUM(IF(A1:A9>0,A1:A9))`), array constants `{1,2;3,4}` are read, and functions with an array answer (`SORT`, `FILTER`, `UNIQUE`, `SEQUENCE`, `MMULT`, …) can be used inside a formula. A formula whose result is an array spills into the cells to its right and below, as in Excel 365 (`#SPILL!` where they are taken), and `E2#` refers to the range it filled (see [Arrays](docs/formulas.md#arrays)).
- **Cells shown as Excel shows them.** Each cell of a result carries, beside its value, the text Excel shows in it: its number format in the book's locale, General narrowed to the column's width, `#` across a column a number does not fit, the colour of the format's section (see [Showing values](docs/formats.md)). The value stays the number; `display: false` leaves the text out. A session also gives what a sheet's conditional formats show in each cell of a window: fills, fonts, colour scales, data bars and icons by the current values.
- **Errors are values.** An invalid formula never throws: its cell gets an Excel error, and the result's `diagnostics` say why, with a check code and the position in the formula.
- **xlsx in and out.** `xlcalc/xlsx` reads an xlsx file into a book, with the values Excel stored and a list of what the book cannot hold, and writes a book with its computed values back to xlsx. It has no dependencies and loads none of the engine's computation (see [Reading and writing xlsx](docs/xlsx.md)).
- **Computing without blocking.** `createInlineCalc` computes in the same thread in portions that give way to the event loop, `createCalcPool` from `xlcalc/worker` in Node worker threads or Web Workers, with soft and hard deadlines, cancellation, a thread's memory limit and limits on the size of a book; a session survives the loss of its thread (see [Computing without blocking](docs/async.md)).
- **No Node APIs.** It runs in Node 20.19 or newer and in browsers (see [Limitations](docs/limitations.md) about Content Security Policy).

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
| objects of a session's cells: `cells.A1.value = 100`, a listener of a cell | `liveWorkbook` from `xlcalc` ([Live objects](docs/live-objects.md)) |
| reading and writing xlsx files | `readXlsx`, `writeXlsx` and `fromLegacy` from `xlcalc/xlsx` |
| computing in threads | `createCalcPool` from `xlcalc/worker`, and `createInlineCalc` from `xlcalc` for portions in the same thread |
| cell addresses in A1 style and a sheet's outline commands, without the engine | `xlcalc/a1` |
| the references of formula text as it is typed, for colouring them, the name being typed for AutoComplete, a number format's category as the ribbon shows it, Find over a host's texts and the sort order of values, without the engine | `xlcalc/text` |
| typed entry, and the text a cell opens with for editing, without the engine | `parseEntry`, `needsApostrophe`, `entryText` from `xlcalc/entry` ([Cell entry](docs/entry.md)) |
| the tips of functions' arguments, and their categories, without the engine's computation | `functionSignature` and `FUNCTION_CATEGORIES` from `xlcalc`; what a function does, `functionDescription` from `xlcalc/function-text`; their Russian words, `ruRUFunctionText` from `xlcalc/function-text/ru-RU` ([Addresses and formula text](docs/addresses.md)) |
| books and formulas in ru-RU | `ruRU` from `xlcalc/locales/ru-RU`, loaded once by `registerLocale(ruRU)` from `xlcalc` ([Localization](docs/locale.md)) |
| types alone | `xlcalc/contract` |

en-US is built in. Bundled by esbuild for the browser and minified (17.10.2026, 0.49.0), `compute` alone takes 284 KB (108 KB gzip), `createSession` 488 KB (184 KB), the whole root 550 KB (204 KB), `xlcalc/xlsx` 204 KB (76 KB), `xlcalc/worker` 25 KB (9 KB) in the page and the thread it starts 598 KB (205 KB), `xlcalc/a1` 16 KB (6 KB), `xlcalc/text` 50 KB (20 KB), `xlcalc/entry` 96 KB (36 KB), the ru-RU module 99 KB (16 KB); `test/bundle-size.test.ts` keeps `compute` under 300 KB.

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

Common tasks, each with the calls it takes; [Documentation](#documentation) has the rules.

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

### Assign to a cell and follow another

```js
import { createSession, liveWorkbook } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*21' } } }] })
const s = liveWorkbook(session).sheet('S')
s.cells.A2.watch(() => console.log('A2 is now', s.cells.A2.value))
s.cells.A1.value = 3
console.log(s.cells.A2.text)
// → A2 is now 63
// → 63
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

## Documentation

- [API reference](docs/api.md): every function, value and type of the package, entry by entry.

Books and formulas:

- [Formulas](docs/formulas.md): arrays inside formulas, checking formula text, errors, external functions, hidden rows, date systems.
- [Functions](docs/functions.md): the whitelist by category.
- [Array operations in formulas](docs/arrays-incell.md) and [Spilling](docs/spill.md).
- [Defined names](docs/names.md), [Tables](docs/tables.md), [References](docs/references.md) (`@`, unions, references across sheets).
- [Circular references](docs/cycles.md), [External references](docs/links.md), [Legacy array formulas and implicit intersection](docs/legacy.md).

Sessions and grids:

- [Sessions](docs/sessions.md): opening one, undo and redo, following changes, the clock.
- [Rendering a grid from a session](docs/grid.md): windows, layout, looks of many cells.
- [Editing a grid](docs/edits.md), [Cell entry](docs/entry.md), [Batch changes, previews and paged queries in a live session](docs/live-books.md).
- [Live sheet and cell objects](docs/live-objects.md): `cells.A1.value = 100`, batches, listeners of a cell or a sheet, a Vue `customRef`.
- [Structure operations](docs/structure.md): rows and columns, moves and copies, the fill handle, merged cells, sheets.
- [Conditional formats](docs/conditional-formats.md), [Data validation](docs/validation.md), [Outlines](docs/outline.md).
- [Addresses and formula text](docs/addresses.md): `xlcalc/a1` and `xlcalc/text`, without the engine.

Values, locales, files and threads:

- [Number formats and displayed text](docs/formats.md), [Locales](docs/locale.md).
- [Reading and writing xlsx](docs/xlsx.md), [Asynchronous and threaded computation](docs/async.md).

The project:

- [Limitations](docs/limitations.md): known differences from Excel and what is not supported.
- [Development](docs/development.md): commands, tests, how the rules were checked against Excel.
- [formula.js against Excel](docs/coverage-formulajs.md), [Range and lookup caches](docs/cache.md), [Numerical stability of statistical moments](docs/statistics.md).

## License

MIT. Function implementations are ported from [formula.js](https://github.com/formulajs/formulajs) (MIT, Copyright (c) 2014 Sutoiku, Inc.), [jStat](https://github.com/jstat/jstat) (MIT, Copyright (c) 2013 jStat) and [bessel](https://github.com/SheetJS/bessel) (Apache-2.0, Copyright (C) 2013-present SheetJS); [THIRD_PARTY.md](THIRD_PARTY.md) has their license texts and the files they are in.
