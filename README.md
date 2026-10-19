# xlcalc

A formula engine for spreadsheet workbooks in JavaScript. You pass it a workbook (sheets of cells with values and Excel formulas, plus defined names), and it returns the calculated values, the errors with their reasons and a deterministic hash of the values. Formulas use Excel syntax with English function names; [Localization](docs/locale.md) displays and parses them in another language, as Excel does.

- **Checked against Excel.** A function is admitted only after it matches Excel 16 on all of its cases: 43,778, with Excel's results stored in this repository ([Functions](docs/functions.md)). Any other function is `#NAME?`. Known differences are in [Limitations](docs/limitations.md); other inputs can differ too.
- **Compiled once, recalculated incrementally.** Closures by default, so it runs under a strict Content Security Policy, or `new Function` with `execution: 'eval'`, with the same values either way. A session recalculates only what an edit affects: under 10 ms (p95) for an edit of a criteria cell on a 3,400-row SUMIFS workbook ([Compiling and recalculating](docs/sessions.md#compiling-and-recalculating)).
- **Deterministic.** The same workbook and clock give the same values and SHA-256 `hash` in any process or time zone; `TODAY`, `NOW` and `RAND` read the clock ([The clock](docs/sessions.md#the-clock-and-volatile-functions)).
- **Excel 365 arrays.** Operations over ranges, array constants, functions that return arrays, and spilling with `#SPILL!` and `E2#` ([Arrays](docs/formulas.md#arrays)).
- **Displayed text.** Each cell carries the text Excel displays for it by its number format, locale and column width; a session also gives a viewport's conditional formats ([Number formats](docs/formats.md)).
- **Errors are values.** An invalid formula never throws: its cell gets an Excel error, and `diagnostics` say why.
- **xlsx in and out**, without dependencies and without loading the engine ([Reading and writing xlsx](docs/xlsx.md)).
- **Without blocking**: in slices on the same thread or in worker threads, with deadlines, cancellation and memory limits ([Asynchronous computation](docs/async.md)).
- **Node 20.19 or newer and browsers**, no Node APIs.

## Install

```sh
npm install xlcalc
```

ESM with TypeScript types, no dependencies. From CommonJS use `await import('xlcalc')`, or `require()` on Node 20.19+ and 22.12+.

### What to import

A bundle includes only the entry points it imports. Sizes are for the browser, bundled with esbuild and minified (19.10.2026, 0.52.1), gzipped in brackets; `test/bundle-size.test.ts` keeps `compute` under 300 KB.

| import | for | size |
|---|---|---:|
| `compute` from `xlcalc` | calculating a workbook | 294 KB (112) |
| `createSession` from `xlcalc` | a session: edits, structure operations, undo, AutoFit, conditional formats, data validation | 499 KB (188) |
| `liveBook` from `xlcalc` | a session's cells as objects: `cells.A1.value = 100`, listeners ([Live objects](docs/live-objects.md)) | |
| all of `xlcalc` | also `createInlineCalc`, `registerLocale`, `functionSignature` and `FUNCTION_CATEGORIES` for tooltips | 562 KB (208) |
| `xlcalc/xlsx` | `readXlsx`, `writeXlsx`, `fromLegacy` | 205 KB (77) |
| `xlcalc/worker` | `createCalcPool`: worker threads | 33 KB (12), its worker 610 KB (209) |
| `xlcalc/a1` | A1 addresses, a sheet's outline commands | 16 KB (6) |
| `xlcalc/text` | references in typed formula text, AutoComplete, number format categories, Find over a host's text, sort order | 51 KB (20) |
| `xlcalc/entry` | typed entries and the text a cell opens with in edit mode ([Cell entry](docs/entry.md)) | 97 KB (37) |
| `xlcalc/function-text` | what a function does, in a line; Russian in `xlcalc/function-text/ru-RU` | |
| `xlcalc/locales/ru-RU` | ru-RU, loaded once by `registerLocale(ruRU)`; en-US is built in | 99 KB (16) |
| `xlcalc/contract` | types only | |

`xlcalc/xlsx`, `xlcalc/a1`, `xlcalc/text`, `xlcalc/entry` and `xlcalc/function-text` do not load the engine.

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

A formula goes in `formula` without the leading `=` (one given is dropped, with the same `inputHash`), apart from `value`, so text from your data never becomes a formula; the same holds for the formulas of names, tables, conditional formats and data validation. A date is `{ date: '2026-01-31' }` (`{ value: { date: '2026-01-31' } }`) and becomes an Excel serial number. `numberFormat` and `meta` pass through to the result, whose `text` is the value in its number format.

## Recipes

More of them are in the [quick start](docs/quick-start.md#recipes), and the details in the [documentation](#documentation).

### Open a session, edit a cell and get the changes

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*21' } } }] })
const { revision, changes } = session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
console.log(revision, changes.map((c) => `${c.cell}=${c.value}`)) // → 1 [ 'A1=3', 'A2=63' ]
```

### Assign to a cell and watch another

```js
import { createSession, liveBook } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*21' } } }] })
const s = liveBook(session).sheet('S')
s.cells.A2.watch(() => console.log('A2 is now', s.cells.A2.value))
s.cells.A1.value = 3
console.log(s.cells.A2.text)
// → A2 is now 63
// → 63
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

### Calculate in a thread pool with a deadline

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

- [Quick start](docs/quick-start.md): the main uses, from a calculator on a server to a grid, with examples.
- [API reference](docs/api.md): every function, constant and type the package exports.

Workbooks and formulas: [Formulas](docs/formulas.md), [Functions](docs/functions.md), [Array operations](docs/arrays-incell.md), [Spilling](docs/spill.md), [Defined names](docs/names.md), [Tables](docs/tables.md), [References](docs/references.md), [Circular references](docs/cycles.md), [External references](docs/links.md), [Legacy array formulas](docs/legacy.md).

Sessions and grids: [Sessions](docs/sessions.md), [Rendering a grid](docs/grid.md), [Editing a grid](docs/edits.md), [Cell entry](docs/entry.md), [Batch changes and queries](docs/live-books.md), [Live objects](docs/live-objects.md), [Structure operations](docs/structure.md), [Conditional formats](docs/conditional-formats.md), [Data validation](docs/validation.md), [Outlines](docs/outline.md), [Addresses and formula text](docs/addresses.md).

Values, locales, files and threads: [Number formats and displayed text](docs/formats.md), [Locales](docs/locale.md), [Reading and writing xlsx](docs/xlsx.md), [Asynchronous and threaded computation](docs/async.md).

The project: [Limitations](docs/limitations.md), [Development](docs/development.md), [formula.js against Excel](docs/coverage-formulajs.md), [Range and lookup caches](docs/cache.md), [Numerical stability of statistical moments](docs/statistics.md).

## License

MIT. Function implementations are ported from [formula.js](https://github.com/formulajs/formulajs) (MIT, Copyright (c) 2014 Sutoiku, Inc.), [jStat](https://github.com/jstat/jstat) (MIT, Copyright (c) 2013 jStat) and [bessel](https://github.com/SheetJS/bessel) (Apache-2.0, Copyright (C) 2013-present SheetJS), with fixes verified against Excel; criteria and database functions, the lookups, `TEXT` and `VALUE`, dates, `IF` and the volatile functions were written for this library. [THIRD_PARTY.md](THIRD_PARTY.md) has the license texts and the files they are in.
