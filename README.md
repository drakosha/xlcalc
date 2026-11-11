# xlcalc

A spreadsheet formula engine for JavaScript. You give it a workbook as plain data: sheets of cells with values and Excel formulas, plus defined names. It returns the calculated values, the errors with their reasons, and a deterministic hash of the values.

Under heavy development. AI-engineered.

## Quick start

```sh
npm install xlcalc
```

```js
import { compute } from 'xlcalc'

const cells = {
  A1: { value: 'North' }, B1: { value: 120 },
  A2: { value: 'South' }, B2: { value: 80 },
  A3: { value: 'North' }, B3: { value: 45.5 },
  D1: { formula: 'SUMIFS(B:B,A:A,"North")' },
  D2: { formula: 'XLOOKUP("South",A1:A3,B1:B3)' },
  D3: { formula: 'B1/(B2-80)' },
}
const { sheets, diagnostics } = compute({ sheets: [{ name: 'Sales', cells }] })
console.log(sheets[0].cells.D1.value) // → 165.5
console.log(sheets[0].cells.D2.value) // → 80
console.log(sheets[0].cells.D3.value, diagnostics[0].reason) // → { error: 'DIV0' } division by zero
```

A formula goes in `formula`, without the leading `=` (one given is dropped, with the same `inputHash`). Data goes in `value`, so text from your data never becomes a formula. The same holds for the formulas of names, tables, conditional formats and data validation. A date goes in `value` as `{ date: '2026-01-31' }`: the cell is `{ value: { date: '2026-01-31' } }`, and the date becomes an Excel serial number. `numberFormat` and `meta` pass through to the result, whose `text` is the value in its number format.

More examples are in [Recipes](#recipes) below and in the [quick start guide](docs/quick-start.md).

## Checked against real Excel

A function is admitted only after it matches Excel on all of its cases. Any other function is `#NAME?`.

The expected results come from real Excel. A harness drives Excel 16 for Mac, and Excel for Windows on a Windows machine, and the results are stored in the repository ([`test/excel/`](https://github.com/drakosha/xlcalc/tree/master/test/excel)). The tests compare the engine with them, so the tests run without Excel. The functions and formulas have 51,925 cases; displayed text, structure operations, tables and xlsx files are checked on workbooks of their own. Many case sets are predicted first: the expected results are written down before Excel runs them. When a prediction misses, the rule is fixed and tested again on a new set.

How each rule was checked: [Development](docs/development.md#how-the-rules-were-checked). The cases by function: [Functions](docs/functions.md). Known differences: [Limitations](docs/limitations.md); other inputs can differ too.

## Features

- **Compiled once, recalculated incrementally.** Closures by default, so it runs under a strict Content Security Policy, or `new Function` with `execution: 'eval'`, with the same values either way. A session recalculates only what an edit affects: under 10 ms (p95) for an edit of a criteria cell on a 3,400-row SUMIFS workbook ([Compiling and recalculating](docs/sessions.md#compiling-and-recalculating)).
- **Deterministic.** The same workbook and clock give the same values and SHA-256 `hash` in any process or time zone; `TODAY`, `NOW` and `RAND` read the clock ([The clock](docs/sessions.md#the-clock-and-volatile-functions)).
- **Excel 365 arrays and LAMBDA.** Operations over ranges, array constants, functions that return arrays, and spilling with `#SPILL!` and `E2#` ([Arrays](docs/formulas.md#arrays)); functions of a workbook's own with `LAMBDA`, in cells and in names, `MAP` and its kin, `GROUPBY` and `PIVOTBY` ([Functions of your own](docs/formulas.md#functions-of-your-own-lambda)).
- **Displayed text.** Each cell carries the text Excel displays for it by its number format, locale and column width; a session also gives a viewport's conditional formats ([Number formats](docs/formats.md)).
- **Formatting as plain data.** Fonts, fills, borders, alignment and number formats on a cell or a range ([Make a range bold](docs/quick-start.md#make-a-range-bold-and-set-its-number-format), [Formatting](docs/edits.md#formatting)), and rules that color cells by their values ([Conditional formats](docs/conditional-formats.md)); read from xlsx and written back.
- **Errors are values.** An invalid formula never throws: its cell gets an Excel error, and `diagnostics` say why.
- **Other languages.** Formulas use Excel syntax with English function names; [Localization](docs/locale.md) displays and parses them in another language, as Excel does.
- **PivotTables.** Read from xlsx with their slicers and timelines, `GETPIVOTDATA` over them, created and changed in a session, and with `xlcalc/pivot` refreshed and laid out as Excel does: filters, sorting, Show Values As, grouping, calculated fields and items; written back as Excel opens them ([PivotTables](docs/pivot-tables.md)).
- **xlsx in and out**, without dependencies and without loading the engine ([Reading and writing xlsx](docs/xlsx.md)); CSV as Excel opens and saves it ([Reading and writing CSV](docs/csv.md)). Sessions apply and clear filters, reapply them and sort visible rows, with undo ([Filters and sorting](docs/structure.md#applying-a-filter-and-sorting)). Filter criteria, sorts, table slicers and timelines survive round trips; text keeps its font runs ([Rich text](docs/edits.md#rich-text)).
- **Without blocking**: in slices on the same thread or in worker threads, with deadlines, cancellation and memory limits ([Asynchronous computation](docs/async.md)).
- **Node 20.19 or newer and browsers**, no Node APIs.

## What to import

ESM with TypeScript types, no dependencies. From CommonJS use `await import('xlcalc')`, or `require()` on Node 20.19+ and 22.12+.

A bundle includes only the entry points it imports. Sizes are for the browser, bundled with esbuild and minified (11.11.2026, 0.89.0 for the main entry points and xlsx; the worker's as of 0.71.0), gzipped in brackets; `test/bundle-size.test.ts` keeps `compute` under 405 KB.

| import | for | size |
|---|---|---:|
| `compute` from `xlcalc` | calculating a workbook | 401 KB (154) |
| `createSession` from `xlcalc` | a session: edits, structure operations, undo, AutoFit, conditional formats, data validation | 712 KB (269) |
| `liveBook` from `xlcalc` | a session's cells as objects: `cells.A1.value = 100`, listeners ([Live objects](docs/live-objects.md)) | |
| all of `xlcalc` | also `createInlineCalc`, `registerLocale`, `registerRegex`, `functionSignature` and `FUNCTION_CATEGORIES` for tooltips | 799 KB (298) |
| `xlcalc/core` | the same API without five groups of rarely used functions, which `registerFunctions` loads from `xlcalc/functions` ([Groups of functions](docs/function-groups.md)) | `compute` 330 KB (125), all 724 KB (268) |
| `xlcalc/functions` | the groups: `engineering`, `finance`, `lambda`, `statistics`, `xml`, `allFunctions` | 102 KB (41) |
| `xlcalc/xlsx` | `readXlsx`, `writeXlsx`, `fromLegacy` | 345 KB (129) |
| `xlcalc/xlsx/keep` | `openKept`, `writeKept`, `inspectKept`: a workbook written over the file it was read from, keeping its charts, pictures and everything else the workbook does not hold ([Writing over the original file](docs/xlsx-keep.md)) | 438 KB (163) |
| `xlcalc/csv` | `readCsv`, `writeCsv`: a CSV file as Excel opens it, a sheet as Excel saves it ([Reading and writing CSV](docs/csv.md)) | 194 KB (72) |
| `xlcalc/worker` | `createCalcPool`: worker threads | 35 KB (13), its worker 960 KB (330), 630 KB (238) loaded first where a bundler splits off the locales, regex levels, groups of functions and the PivotTable engine it loads |
| `xlcalc/a1` | A1 addresses, a sheet's outline commands | 16 KB (7) |
| `xlcalc/text` | references in typed formula text, AutoComplete, number format categories, Increase and Decrease Decimal, Find over a host's text, sort order, what a PivotTable's cells are, a value field's name and the field list's place for a field, the Accounting Number Format menu | 72 KB (28) |
| `xlcalc/entry` | typed entries and the text a cell opens with in edit mode ([Cell entry](docs/entry.md)) | 112 KB (42) |
| `xlcalc/function-text` | what a function does, in a line; Russian in `xlcalc/function-text/ru-RU` | |
| `xlcalc/locales/ru-RU` | ru-RU, loaded by importing it; en-US is built in | 153 KB (24) |
| `xlcalc/locales` | `loadLocale(tag)`, a package's locale by its tag | 8 KB (3), and the locale's module |
| `xlcalc/pivot` | `pivotEngine`, which `registerPivot` loads: PivotTables refreshed and laid out ([PivotTables](docs/pivot-tables.md)) | 74 KB (27) over the root |
| `xlcalc/regex` | `pcreSyntax` and `pcreEngine`, Excel's syntax and match limit for match mode 3 ([Regular expressions](docs/regex.md)) | 36 KB (14) |
| `xlcalc/contract` | types, and the formula limits `MAX_FORMULA_LENGTH` and `MAX_FORMULA_DEPTH` | |

`xlcalc/xlsx`, `xlcalc/xlsx/keep`, `xlcalc/csv`, `xlcalc/a1`, `xlcalc/text`, `xlcalc/entry`, `xlcalc/function-text`, `xlcalc/locales` (`loadLocale`, a package locale by its tag) and `xlcalc/contract` do not load the engine.

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

| | Formulas with Excel's result | Functions with Excel's result in every case (of 402) |
|---|---:|---:|
| xlcalc 0.67.0 | 99.5% | 395 |
| Univer 1.3.0 | 95.3% | 249 |
| Univer 1.0.4 | 94.5% | 240 |
| IronCalc 0.9.1 | 92.9% | 237 |
| IronCalc 0.9.0 | 93.0% | 241 |
| HyperFormula 3.5.0 | 80.5% | 171 |

These are the cases xlcalc's functions were checked against, so they favor it.

Speed and memory, measured on 29.10.2026 with xlcalc 0.67.0, HyperFormula 3.5.0, IronCalc 0.9.0, SheetJS 0.18.5 and ExcelJS 4.4.0 on Node 26.7 and an Apple M5 Pro, each case in its own process, with xlcalc's faster mode of execution, HyperFormula's best configuration and IronCalc's WebAssembly build; IronCalc's memory is its heap with the WebAssembly memory it grew to, an upper bound. A dash marks values that differ from xlcalc's full calculation: IronCalc's after a row is inserted in the SUMIFS workbook. The protocol and every case are in `bench/compare` in the repository.

| Milliseconds, median | xlcalc | HyperFormula | IronCalc |
|---|---:|---:|---:|
| Open a 3,400-row workbook of 3,907 SUMIFS formulas | 104 | 3,765 | 2,208 |
| Edit a criterion of its SUMIFS | 0.163 | 0.964 | 2,046 |
| Edit an amount it sums | 6.46 | 2,800 | 2,065 |
| Open 50,000 lookups | 44.5 | 1,272 | 13,478 |
| Edit a lookup key | 0.146 | 0.100 | 13,599 |
| Open a chain of 10,000 formulas | 34.8 | 31.0 | 25.0 |
| Edit its first cell | 2.30 | 4.51 | 1.96 |
| Open 50,000 formulas over one cell | 131 | 366 | 113 |
| Edit that cell | 16.0 | 42.3 | 18.1 |
| Set 2,000 inputs at once | 2.85 | 3.31 | 2.62 |
| Insert a row in the SUMIFS workbook | 91.0 | 2,805 | – |
| Insert a row among 5,000 cumulative SUMs | 25.2 | 6.36 | 185 |

| Retained heap after opening, MiB | xlcalc | HyperFormula | IronCalc |
|---|---:|---:|---:|
| The SUMIFS workbook | 14.2 | 17.7 | 22.0 |
| 50,000 lookups | 7.42 | 20.4 | 26.5 |
| A chain of 10,000 formulas | 7.41 | 4.76 | 6.52 |
| 1,000 values 97 rows apart | 1.36 | 0.848 | 3.77 |

| xlsx, 34,000 rows | xlcalc | SheetJS | ExcelJS |
|---|---:|---:|---:|
| Read, ms | 408 | 527 | 579 |
| Write, ms | 518 | 368 (compressed 473) | 1,068 |
| File written, MB | 3.50 | 20.0 (compressed 8.69) | 3.54 |

SheetJS writes without compression by default.

## Documentation

- [Quick start](docs/quick-start.md): the main uses, from a calculator on a server to a grid, with examples.
- [API reference](docs/api.md): every function, constant and type the package exports.

Workbooks and formulas: [Formulas](docs/formulas.md), [Functions](docs/functions.md), [Forecasting with exponential smoothing](docs/forecast.md), [Array operations](docs/arrays-incell.md), [Spilling](docs/spill.md), [Defined names](docs/names.md), [Tables](docs/tables.md), [References](docs/references.md), [Circular references](docs/cycles.md), [External references](docs/links.md), [Legacy array formulas](docs/legacy.md).

Sessions and grids: [Sessions](docs/sessions.md), [Rendering a grid](docs/grid.md), [Editing a grid](docs/edits.md), [Cell entry](docs/entry.md), [Batch changes and queries](docs/live-books.md), [Live objects](docs/live-objects.md), [Structure operations](docs/structure.md), [Conditional formats](docs/conditional-formats.md), [Data validation](docs/validation.md), [Outlines](docs/outline.md), [PivotTables](docs/pivot-tables.md), [Addresses and formula text](docs/addresses.md).

Values, locales, files and threads: [Number formats and displayed text](docs/formats.md), [Locales](docs/locale.md), [Reading and writing xlsx](docs/xlsx.md), [Writing over the original file](docs/xlsx-keep.md), [Reading and writing CSV](docs/csv.md), [Asynchronous and threaded computation](docs/async.md), [Regular expressions in XLOOKUP and XMATCH](docs/regex.md), [Groups of functions](docs/function-groups.md).

The project: [Limitations](docs/limitations.md), [Development](docs/development.md), [formula.js against Excel](docs/coverage-formulajs.md), [Range and lookup caches](docs/cache.md), [Numerical stability of statistical moments](docs/statistics.md).

## License

MIT. Function implementations are ported from [formula.js](https://github.com/formulajs/formulajs) (MIT, Copyright (c) 2014 Sutoiku, Inc.), [jStat](https://github.com/jstat/jstat) (MIT, Copyright (c) 2013 jStat) and [bessel](https://github.com/SheetJS/bessel) (Apache-2.0, Copyright (C) 2013-present SheetJS), with fixes verified against Excel; criteria and database functions, the lookups, `TEXT` and `VALUE`, dates, `IF` and the volatile functions were written for this library. [THIRD_PARTY.md](THIRD_PARTY.md) has the license texts and the files they are in.
