# xlcalc

A formula engine for spreadsheet workbooks in JavaScript. You pass it a workbook (sheets of cells with values and Excel formulas, plus defined names), and it returns the calculated values, the errors with their reasons and a deterministic hash of the values. Formulas use Excel syntax with English function names; [Localization](docs/locale.md) displays and parses them in another language, as Excel does.

Under heavy development. AI-engineered.

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

A bundle includes only the entry points it imports. Sizes are for the browser, bundled with esbuild and minified (28.10.2026, 0.65.0), gzipped in brackets; `test/bundle-size.test.ts` keeps `compute` under 325 KB.

| import | for | size |
|---|---|---:|
| `compute` from `xlcalc` | calculating a workbook | 305 KB (116) |
| `createSession` from `xlcalc` | a session: edits, structure operations, undo, AutoFit, conditional formats, data validation | 543 KB (204) |
| `liveBook` from `xlcalc` | a session's cells as objects: `cells.A1.value = 100`, listeners ([Live objects](docs/live-objects.md)) | |
| all of `xlcalc` | also `createInlineCalc`, `registerLocale`, `functionSignature` and `FUNCTION_CATEGORIES` for tooltips | 612 KB (227) |
| `xlcalc/xlsx` | `readXlsx`, `writeXlsx`, `fromLegacy` | 229 KB (85) |
| `xlcalc/xlsx/keep` | `openKept`, `writeKept`: a workbook written over the file it was read from, keeping its charts, pictures and everything else the workbook does not hold ([Writing over the original file](docs/xlsx-keep.md)) | 301 KB (112) |
| `xlcalc/worker` | `createCalcPool`: worker threads | 25 KB (9), its worker 666 KB (227) |
| `xlcalc/a1` | A1 addresses, a sheet's outline commands | 16 KB (6) |
| `xlcalc/text` | references in typed formula text, AutoComplete, number format categories, Increase and Decrease Decimal, Find over a host's text, sort order | 50 KB (19) |
| `xlcalc/entry` | typed entries and the text a cell opens with in edit mode ([Cell entry](docs/entry.md)) | 101 KB (37) |
| `xlcalc/function-text` | what a function does, in a line; Russian in `xlcalc/function-text/ru-RU` | |
| `xlcalc/locales/ru-RU` | ru-RU, loaded once by `registerLocale(ruRU)`; en-US is built in | 110 KB (17) |
| `xlcalc/contract` | types, and the formula limits `MAX_FORMULA_LENGTH` and `MAX_FORMULA_DEPTH` | |

`xlcalc/xlsx`, `xlcalc/xlsx/keep`, `xlcalc/a1`, `xlcalc/text`, `xlcalc/entry`, `xlcalc/function-text`, `xlcalc/locales` (`loadLocale`, a package locale by its tag) and `xlcalc/contract` do not load the engine.

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

## Compared with other libraries

Each library calculated the 5,275 formulas of xlcalc's coverage run over the same data, and its results were compared with Excel 16's, which this repository stores, by the same tolerance and with no allowance for known differences:

| | xlcalc | Univer 1.0.4 | IronCalc 0.9.0 | HyperFormula 3.5.0 |
|---|---:|---:|---:|---:|
| Formulas with Excel's result | 99.5% | 94.5% | 93.0% | 80.5% |
| Functions with Excel's result in every case (of 402) | 395 | 240 | 241 | 171 |

These are the cases xlcalc's functions were checked against, so they favor it.

Speed and memory on Node 26.7 and an Apple M5 Pro, each case in its own process, with xlcalc's faster mode of execution, HyperFormula's best configuration and IronCalc's WebAssembly build; IronCalc's memory is its heap with the WebAssembly memory it grew to, an upper bound. A dash marks values that differ from xlcalc's full calculation: IronCalc's after a row is inserted in the SUMIFS workbook. The protocol and every case are in `bench/compare` in the repository.

| Milliseconds, median | xlcalc | HyperFormula | IronCalc |
|---|---:|---:|---:|
| Open a 3,400-row workbook of 3,907 SUMIFS formulas | 98.6 | 3,765 | 2,207 |
| Edit a criterion of its SUMIFS | 0.151 | 0.956 | 2,070 |
| Edit an amount it sums | 8.51 | 2,810 | 2,134 |
| Open 50,000 lookups | 46.9 | 1,119 | 13,648 |
| Edit a lookup key | 0.156 | 0.095 | 12,875 |
| Open a chain of 10,000 formulas | 32.8 | 30.2 | 24.9 |
| Edit its first cell | 2.21 | 4.51 | 2.30 |
| Open 50,000 formulas over one cell | 128 | 362 | 111 |
| Edit that cell | 15.2 | 40.6 | 16.1 |
| Set 2,000 inputs at once | 2.58 | 2.92 | 2.50 |
| Insert a row in the SUMIFS workbook | 107 | 2,669 | – |
| Insert a row among 5,000 cumulative SUMs | 24.3 | 6.36 | 187 |

| Retained heap after opening, MB | xlcalc | HyperFormula | IronCalc |
|---|---:|---:|---:|
| The SUMIFS workbook | 13.6 | 17.7 | 22.0 |
| 50,000 lookups | 8.66 | 20.4 | 26.5 |
| A chain of 10,000 formulas | 7.19 | 4.76 | 6.52 |
| 1,000 values 97 rows apart | 1.16 | 0.842 | 3.77 |

| xlsx, 34,000 rows | xlcalc | SheetJS | ExcelJS |
|---|---:|---:|---:|
| Read, ms | 446 | 546 | 596 |
| Write, ms | 664 | 347 (compressed 461) | 1,068 |
| File written, MB | 3.50 | 20.0 (compressed 8.69) | 3.54 |

SheetJS writes without compression by default.

## Documentation

- [Quick start](docs/quick-start.md): the main uses, from a calculator on a server to a grid, with examples.
- [API reference](docs/api.md): every function, constant and type the package exports.

Workbooks and formulas: [Formulas](docs/formulas.md), [Functions](docs/functions.md), [Array operations](docs/arrays-incell.md), [Spilling](docs/spill.md), [Defined names](docs/names.md), [Tables](docs/tables.md), [References](docs/references.md), [Circular references](docs/cycles.md), [External references](docs/links.md), [Legacy array formulas](docs/legacy.md).

Sessions and grids: [Sessions](docs/sessions.md), [Rendering a grid](docs/grid.md), [Editing a grid](docs/edits.md), [Cell entry](docs/entry.md), [Batch changes and queries](docs/live-books.md), [Live objects](docs/live-objects.md), [Structure operations](docs/structure.md), [Conditional formats](docs/conditional-formats.md), [Data validation](docs/validation.md), [Outlines](docs/outline.md), [Addresses and formula text](docs/addresses.md).

Values, locales, files and threads: [Number formats and displayed text](docs/formats.md), [Locales](docs/locale.md), [Reading and writing xlsx](docs/xlsx.md), [Writing over the original file](docs/xlsx-keep.md), [Asynchronous and threaded computation](docs/async.md).

The project: [Limitations](docs/limitations.md), [Development](docs/development.md), [formula.js against Excel](docs/coverage-formulajs.md), [Range and lookup caches](docs/cache.md), [Numerical stability of statistical moments](docs/statistics.md).

## License

MIT. Function implementations are ported from [formula.js](https://github.com/formulajs/formulajs) (MIT, Copyright (c) 2014 Sutoiku, Inc.), [jStat](https://github.com/jstat/jstat) (MIT, Copyright (c) 2013 jStat) and [bessel](https://github.com/SheetJS/bessel) (Apache-2.0, Copyright (C) 2013-present SheetJS), with fixes verified against Excel; criteria and database functions, the lookups, `TEXT` and `VALUE`, dates, `IF` and the volatile functions were written for this library. [THIRD_PARTY.md](THIRD_PARTY.md) has the license texts and the files they are in.
