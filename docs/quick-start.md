# Quick start

Common uses of the engine, each with an example that runs as written and prints what its `// →` comments show. The README's [Recipes](../README.md#recipes) cover single calls; each section here links to the full documentation.

## A calculator on a server

A model is a workbook whose formulas turn input cells into results: a loan, a price, a tariff. The host stores it as data, fills in the inputs of each request and calculates. Each cell of the result has its `value` and the `text` it displays with its number format, as Excel displays it in a column of that width (a column that is too narrow shows `#`, so the model widens column B). An input that makes a formula fail gives the cell an error value and adds a reason to `diagnostics`; `compute` does not throw.

```js
import { compute } from 'xlcalc'

// the model, stored as data: inputs in B1:B3, results in B5:B6
const loan = {
  name: 'Loan',
  columns: { B: { width: 12 } },
  cells: {
    A1: { value: 'Amount' }, A2: { value: 'Rate' }, A3: { value: 'Years' },
    A5: { value: 'Payment' }, B5: { formula: '-PMT(B2/12,B3*12,B1)', numberFormat: '#,##0.00' },
    A6: { value: 'Interest' }, B6: { formula: 'B5*B3*12-B1', numberFormat: '#,##0.00' },
  },
}

function quote(amount, rate, years) {
  const cells = { ...loan.cells, B1: { value: amount }, B2: { value: rate }, B3: { value: years } }
  const r = compute({ sheets: [{ ...loan, cells }] })
  const { B5, B6 } = r.sheets[0].cells
  return { payment: B5.value, shown: [B5.text, B6.text], issues: r.diagnostics.map((d) => `${d.cell}: ${d.reason}`) }
}

console.log(quote(200_000, 0.06, 20))
// → {
//     payment: 1432.86211695635,
//     shown: [ '1,432.86', '143,886.91' ],
//     issues: []
//   }
const bad = quote(200_000, 0.06, 0)
console.log(bad.payment, bad.issues)
// → { error: 'NUM' } [
//     'B5: number out of range',
//     "B6: number out of range, from 'Loan'!B5"
//   ]
```

`compute` compiles the workbook on every call. A model that serves many requests in a row can stay open in a session, which compiles it once and recalculates only what new inputs affect (next section). A result's `inputHash` identifies its input, so for a workbook without volatile functions (the result's `volatile` is empty) it can serve as a cache key. [Formulas](formulas.md) describes what a formula can contain, and [Errors](formulas.md#errors) the error values and their reasons.

## What-if: change an input and see the effect

A session holds a calculated workbook. `setCells` changes inputs and returns the cells whose value or displayed text changed, which are the cells a form or a chart needs to redraw. `previewChanges` calculates the effect of a change without applying it, `undo` reverts a change, and `explain` shows where a number comes from.

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{
    name: 'Loan',
    columns: { B: { width: 12 } },
    cells: {
      B1: { value: 200_000 }, B2: { value: 0.06 }, B3: { value: 20 },
      B5: { formula: '-PMT(B2/12,B3*12,B1)', numberFormat: '#,##0.00' },
      B6: { formula: 'B5*B3*12-B1', numberFormat: '#,##0.00' },
    },
  }],
})

const lower = session.previewChanges([{ op: 'setCells', edits: [{ sheet: 'Loan', cell: 'B2', input: { value: 0.05 } }] }])
console.log(lower.impacts.map((i) => `${i.cell} ${Math.round(i.after * 100) / 100}`), session.revision) // → [ 'B2 0.05', 'B5 1319.91', 'B6 116778.75' ] 0

const { changes } = session.setCells([{ sheet: 'Loan', cell: 'B3', input: { value: 15 } }])
console.log(changes.map((c) => `${c.cell} ${c.text}`)) // → [ 'B3 15', 'B5 1,687.71', 'B6 103,788.46' ]
console.log(session.explain('Loan', 'B5').text)
// → Loan!B5 = 1687.7136560969248  ← -PMT(B2/12,B3*12,B1)
//     Loan!B2 = 0.06
//     Loan!B3 = 15
//     Loan!B1 = 200000
session.undo()
console.log(session.getCell('Loan', 'B5').text, session.history) // → 1,432.86 { undo: 0, redo: 1 }
```

For a form bound to cells, `liveWorkbook(session)` returns objects where an assignment is an edit (`cells.B2.value = 0.05`) and a cell accepts listeners, a basis for a Vue `customRef` or a signal ([Live sheet and cell objects](live-objects.md)). [Sessions](sessions.md) documents the session's methods, and [Batch changes, previews and paged queries](live-books.md) the preview.

## A spreadsheet grid

A grid passes what the user types to `setCellText`, which parses it as Excel does: a date, a number with separators, a percentage and a formula, each with the number format Excel applies to it. `getViewport` reads the range the grid displays, with each cell's displayed text and the text its editor opens with. Text that does not parse as a formula is rejected with the error position, and the workbook stays unchanged.

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'Sheet1', cells: {} }] })
for (const [cell, text] of [['A1', '1/15/2026'], ['A2', '1,250.5'], ['A3', '12%'], ['A4', '=A2*A3']]) session.setCellText('Sheet1', cell, text)
const { cells } = session.getViewport('Sheet1', 'A1:A4', { editText: true })
console.log(Object.entries(cells).map(([a, c]) => `${a} ${c.value} ${c.text} ${c.editText}`))
// → [
//     'A1 46037 1/15/2026 1/15/2026',
//     'A2 1250.5 1,250.50 1250.5',
//     'A3 0.12 12% 12%',
//     'A4 150.06 150.06 =A2*A3'
//   ]
const typo = session.setCellText('Sheet1', 'A5', '=A2*')
console.log(typo.ok, typo.code, typo.at, typo.message) // → false syntax 4 unexpected end of formula
console.log(session.undo().changes.map((c) => `${c.cell} ${c.value}`), session.history) // → [ 'A4 null' ] { undo: 3, redo: 1 }
```

[Rendering a grid from a session](grid.md) covers viewports, layout and formatting; [Editing a grid](edits.md) covers styles, clearing and copying; [Cell entry](entry.md) typed text and the clipboard; [Structure operations](structure.md) rows, columns and sheets.

## Excel files

`readXlsx` returns a workbook and `cached`, the values Excel stored for its formulas. Comparing them with the engine's result shows where the engine differs from Excel on this file before a host relies on it: here it is `INDIRECT`, a function outside the whitelist, which the engine evaluates to `#NAME?`. The workbook then opens in a session, and `writeXlsx` saves it with the new values.

```js
import { compute, createSession, sameValue } from 'xlcalc'
import { readXlsx, writeXlsx } from 'xlcalc/xlsx'

// a file saved by Excel, as `await readFile('report.xlsx')` returns it; here, one written for the example
const saved = { sheets: [{ name: 'Report', cells: { B1: { value: 120 }, B2: { value: 80 }, B3: { formula: 'SUM(B1:B2)' }, C1: { formula: 'INDIRECT("B1")' } } }] }
const file = (await writeXlsx({ book: saved, result: { sheets: [{ name: 'Report', cells: { B3: { value: 200 }, C1: { value: 120 } } }] } })).data

const read = await readXlsx(file)
const result = compute(read.book, { cached: read.cached })
const differ = read.cached.sheets.flatMap((sheet, i) =>
  Object.entries(sheet.cells)
    .filter(([cell, c]) => !sameValue(c.value, result.sheets[i].cells[cell]?.value))
    .map(([cell]) => `${sheet.name}!${cell}`))
console.log(differ, read.issues) // → [ 'Report!C1' ] []

const session = createSession(read.book)
session.setCells([{ sheet: 'Report', cell: 'B1', input: { value: 150 } }])
const { data, issues } = await writeXlsx(session.getSnapshot()) // data: bytes for writeFile or a download
console.log((await readXlsx(data)).cached.sheets[0].cells.B3, issues) // → { value: 230 } []
```

`read.issues` lists what the workbook could not keep from the file (pictures, charts, rich text), with the sheet and cells. [Reading and writing xlsx](xlsx.md) describes what is read, written and reported.

## Formulas people type

Before a formula from a form, an API or an import is stored, `validateFormula` checks it against the workbook's sheets and names. It returns the canonical text and the functions, sheets and names the formula uses, or an error code, a message and the position to highlight.

```js
import { validateFormula } from 'xlcalc'

const book = { sheets: ['Sales', 'Costs'], names: ['tax_rate'] }
for (const text of ['sum(Sales!b:b)*tax_rate', 'SUM(Sales!B:B', 'Sales!B1*vat', 'Budget!A1', 'INDIRECT("A1")']) {
  const v = validateFormula(text, book)
  console.log(v.ok ? `ok ${v.canonical}` : `${v.code} at ${v.at}: ${v.message}`)
}
// → ok SUM('Sales'!B:B)*tax_rate
// → syntax at 13: expected ), found end of formula
// → unknown_name at 9: unknown name vat
// → unknown_sheet at 0: unknown sheet Budget
// → unknown_function at 0: unknown function INDIRECT
```

A formula stored without validation still calculates: an unknown function or name evaluates to `#NAME?` and an unknown sheet to `#REF!`, with the reason in `diagnostics`. Typed text goes through `setCellText` instead, which closes an unclosed parenthesis as Excel does. [Checking formulas](formulas.md#checking-formulas) lists the error codes, and `validateLocalFormula` validates text in a locale.

## Values from your own systems

A formula can call a function the host defines, for example a balance from a database. `validateFormula` lists the calls with their arguments, the host looks up each one, and `compute` accepts the values. Arguments must be literals, and a call without a value evaluates to `#N/A`.

```js
import { compute, validateFormula } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { formula: 'BALANCE("40702")+BALANCE("40817")' }, A2: { formula: 'A1*0.2' } } }] }
console.log(validateFormula(book.sheets[0].cells.A1.formula, { externalFunctions: ['BALANCE'] }).externals) // → [ 'BALANCE("40702")', 'BALANCE("40817")' ]
const values = { 'BALANCE("40702")': 1500, 'BALANCE("40817")': 250 } // the host looks up each call
const r = compute(book, { externals: { functions: ['BALANCE'], values } })
console.log(r.sheets[0].cells.A1.value, r.sheets[0].cells.A2.value) // → 1750 350
```

A session accepts the same `externals` when it is created ([External functions](formulas.md#external-functions)).

## Reproducible results

`TODAY`, `NOW`, `RAND` and `RANDBETWEEN` read a clock: a moment in UTC and a seed. The same workbook with the same clock produces the same values and the same `hash` in any process and time zone, so a stored calculation can be verified later, and a test can pin a workbook's values by its hash. Without a clock the engine uses the current time and a random seed, and the result's `clock` reproduces that result.

```js
import { compute } from 'xlcalc'

const book = { sheets: [{ name: 'S', columns: { A: { width: 12 } }, cells: { A1: { formula: 'TODAY()', numberFormat: 'yyyy-mm-dd' }, A2: { formula: 'RANDBETWEEN(1,100)' } } }] }
const clock = { now: '2026-10-18T09:00:00Z', seed: 7 }
const a = compute(book, { clock })
const b = compute(book, { clock })
console.log(a.sheets[0].cells.A1.text, a.sheets[0].cells.A2.value, a.hash === b.hash) // → 2026-10-18 95 true

const now = compute(book) // the current time and a random seed
console.log(compute(book, { clock: now.clock }).hash === now.hash, now.inputHash === a.inputHash, now.volatile.length) // → true true 2
```

The hash includes `configVersion`, which changes with any engine change that can affect values, so hashes are comparable only between results of the same version. A session keeps its clock until `setClock` is called, or takes a fresh one on every recalculation with `live: true` ([Sessions](sessions.md)).

## Large workbooks without blocking

`createInlineCalc` calculates on the same thread in slices that yield to the event loop, so timers and input keep running while a workbook calculates; `createCalcPool` from `xlcalc/worker` calculates in worker threads. Both accept a deadline and a cancellation signal for each call and produce the same values and hashes as `compute`.

```js
import { createInlineCalc } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: Object.fromEntries(Array.from({ length: 30_000 }, (_, i) => [`A${i + 1}`, { formula: `SUM(SEQUENCE(${(i % 100) + 1}))` }])) }] }
const calc = createInlineCalc()
let ticks = 0
const timer = setInterval(() => ticks++, 1) // a timer keeps running while the workbook calculates
const r = await calc.compute(book, { timeoutMs: 10_000 })
clearInterval(timer)
const late = await calc.compute(book, { timeoutMs: 5 }).catch((e) => e)
console.log(r.sheets[0].cells.A100.value, ticks > 0, late.code) // → 5050 true soft_timeout
await calc.close()
```

The README shows the same with a thread pool; [Asynchronous and threaded computation](async.md) covers deadlines, limits and what happens when a thread is lost.

## Locales

A locale is Excel's language and regional settings, and it has two roles. A workbook calculates and displays values according to its `locale`, as Excel does on a machine with those settings: which text is recognized as a number or a date, how numbers and booleans are converted to text, and how number formats display values. A locale is also the language in which formulas are entered and displayed (`=ОКРУГЛ(A1*1,2;0)` in ru-RU). Whatever its locale, a workbook stores formulas as xlsx files do, with English function names, `,` between arguments and `.` as the decimal point. The package includes en-US, which is built in, and ru-RU, a separate module that `registerLocale` loads once, so a bundle without ru-RU workbooks does not include it.

```js
import { compute, createSession, registerLocale, toInvariant, toLocal } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU) // once, before any ru-RU workbook is calculated

// the same cells under each locale's rules
const sheet = {
  name: 'S',
  columns: { A: { width: 12 } },
  cells: {
    A1: { value: 1234.5, numberFormat: '#,##0.00' },
    A2: { formula: 'DATEVALUE("3/4/2026")', numberFormat: 'yyyy-mm-dd' },
    A3: { formula: '1.5&" "&TRUE' },
  },
}
for (const locale of ['en-US', 'ru-RU']) {
  const { A1, A2, A3 } = compute({ locale, sheets: [sheet] }).sheets[0].cells
  console.log(locale, A1.text, A2.text, A3.text)
}
// → en-US 1,234.50 2026-03-04 1.5 TRUE
// → ru-RU 1 234,50 2026-04-03 1,5 ИСТИНА

// typed text is parsed in the workbook's locale, and the formula is stored in en-US syntax
const session = createSession({ locale: 'ru-RU', sheets: [{ name: 'Лист1', cells: { A1: { value: 1234.5 } } }] })
session.setCellText('Лист1', 'A2', '=ОКРУГЛ(A1*1,2;0)')
const { A2 } = session.getViewport('Лист1', 'A2', { editText: true }).cells
console.log(A2.value, A2.formula, A2.editText) // → 1481 ROUND(A1*1.2,0) =ОКРУГЛ(A1*1,2;0)
console.log(session.setCellText('Лист1', 'A3', '=ОКРУГЛ(A1;2;3)').message) // → Функция ОКРУГЛ принимает 2 аргумента, а указано 3

// any workbook's formula in the reader's language, and what they type converted back
console.log(toLocal('IF(B1>0.5,SUM(B2:B9),#N/A)', ruRU).formula) // → ЕСЛИ(B1>0,5;СУММ(B2:B9);#Н/Д)
console.log(toInvariant('=СУММ(B2:B9;1,5)', ruRU).formula) // → SUM(B2:B9,1.5)
```

Calculating a workbook or opening a session on one whose locale is not loaded throws an `InputError` that names the import to add. An xlsx file does not record which locale it is calculated in, so the same file gives different values in Excel on a Russian machine and on an English one (`3/4/2026` above); a workbook that carries its `locale` gives the same values everywhere. A session in ru-RU parses typed text as Russian Excel does (`1 234,5`, `15.03.2026`) and explains in Russian why input is rejected and where a value comes from (`explain`); the reasons in `diagnostics` stay in English. `toLocal`, `toInvariant` and `validateLocalFormula` accept the locale object without loading it, so users can read and enter formulas in Russian in a workbook that calculates in en-US. [Locales](locale.md) covers what changes in calculation, which messages are translated and how to add a locale.
