# Quick start

The jobs the engine is put to, each with an example that runs as shown and prints what its `// →` comments say. The README's [Recipes](../README.md#recipes) give single calls; each section here links to the full rules.

## A calculator on a server

A model is a book whose formulas turn input cells into results: a loan, a price, a tariff. The host keeps it as data, fills the inputs of each request and computes. Each cell of the result has its `value` and the `text` it shows by its number format, as Excel shows it in a column of that width (`#` fills a column too narrow, so the model widens B). An input that makes a formula fail gives the cell an error value and `diagnostics` a reason; `compute` does not throw on it.

```js
import { compute } from 'xlcalc'

// the model, kept as data: inputs in B1:B3, results in B5:B6
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

`compute` compiles the book on every call. A model that serves many requests one after another can stay open in a session, which compiles it once and recomputes only what new inputs touch (next section). `inputHash` of a result names its input, so for a book without volatile functions (`volatile` of the result is empty) it can key a cache of results. [Formulas](formulas.md) has what a formula can hold, [Errors](formulas.md#errors) the error values and reasons.

## What-if: change an input and see what follows

A session holds a computed book. `setCells` changes inputs and answers the cells whose value or shown text changed, the cells a form or a chart draws again. `previewChanges` computes what a change would do without making it, `undo` takes a change back, and `explain` shows where a number comes from.

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

For a form bound to cells, `liveWorkbook(session)` gives objects whose assignment is an edit (`cells.B2.value = 0.05`) and listeners of a cell, a base for a Vue `customRef` or a signal ([Live sheet and cell objects](live-objects.md)). [Sessions](sessions.md) has the session's calls and [Batch changes, previews and paged queries](live-books.md) the preview's.

## A spreadsheet grid

A grid passes what a person types to `setCellText`, which reads it as Excel does: a date, a number with separators, a percent and a formula, each with the format Excel gives it. `getViewport` reads the window the grid shows, with each cell's shown text and the text its editor opens with. Text that does not read as a formula is refused with its position, and the book stays as it was.

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

[Rendering a grid from a session](grid.md) covers windows, layout and looks, [Editing a grid](edits.md) styles, clearing and copying, [Cell entry](entry.md) typed text and the clipboard, [Structure operations](structure.md) rows, columns and sheets.

## Excel files

`readXlsx` gives a book and `cached`, the values Excel stored for its formulas. Compared with the engine's result, they show where the engine answers otherwise than Excel on this file before a host relies on it: here `INDIRECT`, a function outside the whitelist, which the engine computes as `#NAME?`. The book then opens in a session, and `writeXlsx` saves it with the new values.

```js
import { compute, createSession, sameValue } from 'xlcalc'
import { readXlsx, writeXlsx } from 'xlcalc/xlsx'

// a file Excel saved, as `await readFile('report.xlsx')` gives it; here one written for the example
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

`read.issues` lists what the book could not keep from the file (pictures, charts, rich text), with the sheet and cells. [Reading and writing xlsx](xlsx.md) has what is read, written and reported.

## Formulas people type

Before a formula from a form, an API or an import is stored, `validateFormula` checks it against the book's sheets and names. It answers the canonical text and the functions, sheets and names the formula uses, or a check code, a message and the position to point at.

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

A formula stored without the check still computes: an unknown function or name gives `#NAME?`, an unknown sheet `#REF!`, with the reason in `diagnostics`. Typed text goes through `setCellText` instead, which closes an open parenthesis as Excel does. [Checking formulas](formulas.md#checking-formulas) has the check codes, and `validateLocalFormula` checks text in a locale.

## Values from your own systems

A formula can call a function the host defines, a balance from a database, say. `validateFormula` lists the calls with their arguments, the host looks each one up, and `compute` takes the values. Arguments must be literals, and a call without a value gives `#N/A`.

```js
import { compute, validateFormula } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { formula: 'BALANCE("40702")+BALANCE("40817")' }, A2: { formula: 'A1*0.2' } } }] }
console.log(validateFormula(book.sheets[0].cells.A1.formula, { externalFunctions: ['BALANCE'] }).externals) // → [ 'BALANCE("40702")', 'BALANCE("40817")' ]
const values = { 'BALANCE("40702")': 1500, 'BALANCE("40817")': 250 } // the host looks each call up
const r = compute(book, { externals: { functions: ['BALANCE'], values } })
console.log(r.sheets[0].cells.A1.value, r.sheets[0].cells.A2.value) // → 1750 350
```

A session takes the same `externals` when it opens ([External functions](formulas.md#external-functions)).

## Results that reproduce

`TODAY`, `NOW`, `RAND` and `RANDBETWEEN` take a clock: a moment in UTC and a seed. The same book with the same clock gives the same values and the same `hash` in any process and time zone, so a stored computation can be checked later and a test can pin a book's values by its hash. Without a clock the engine takes the current moment and a random seed, and the result's `clock` computes that result again.

```js
import { compute } from 'xlcalc'

const book = { sheets: [{ name: 'S', columns: { A: { width: 12 } }, cells: { A1: { formula: 'TODAY()', numberFormat: 'yyyy-mm-dd' }, A2: { formula: 'RANDBETWEEN(1,100)' } } }] }
const clock = { now: '2026-10-18T09:00:00Z', seed: 7 }
const a = compute(book, { clock })
const b = compute(book, { clock })
console.log(a.sheets[0].cells.A1.text, a.sheets[0].cells.A2.value, a.hash === b.hash) // → 2026-10-18 95 true

const now = compute(book) // the current moment and a random seed
console.log(compute(book, { clock: now.clock }).hash === now.hash, now.inputHash === a.inputHash, now.volatile.length) // → true true 2
```

The hash includes `configVersion`, which changes with any change of the engine that can shift values, so hashes compare between results of one version. A session keeps its clock until `setClock`, or takes a fresh one on every recomputation with `live: true` ([Sessions](sessions.md)).

## Large books without blocking

`createInlineCalc` computes in the same thread in slices that give way to the event loop, so timers and input go on while a book computes; `createCalcPool` from `xlcalc/worker` computes in worker threads. Both take a deadline and a cancellation signal for each call and give the same values and hashes as `compute`.

```js
import { createInlineCalc } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: Object.fromEntries(Array.from({ length: 30_000 }, (_, i) => [`A${i + 1}`, { formula: `SUM(SEQUENCE(${(i % 100) + 1}))` }])) }] }
const calc = createInlineCalc()
let ticks = 0
const timer = setInterval(() => ticks++, 1) // a timer goes on while the book computes
const r = await calc.compute(book, { timeoutMs: 10_000 })
clearInterval(timer)
const late = await calc.compute(book, { timeoutMs: 5 }).catch((e) => e)
console.log(r.sheets[0].cells.A100.value, ticks > 0, late.code) // → 5050 true soft_timeout
await calc.close()
```

The README has the same with a pool of threads; [Asynchronous and threaded computation](async.md) has the deadlines, limits and what happens when a thread is lost.

## Books in Russian

`registerLocale(ruRU)` loads ru-RU once. A book with `locale: 'ru-RU'` reads typed text as Russian Excel does (`1 234,5`, `15.03.2026`, `=ОКРУГЛ(A1*1,2;0)`), shows values by its rules and writes in Russian why typed text is refused and where a value comes from. The book keeps formulas in en-US syntax, as xlsx files do, and the editor opens with the Russian text; the reasons in `diagnostics` stay in English.

```js
import { createSession, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const session = createSession({ locale: 'ru-RU', sheets: [{ name: 'Лист1', cells: {} }] })
session.setCellText('Лист1', 'A1', '1 234,5')
session.setCellText('Лист1', 'A2', '=ОКРУГЛ(A1*1,2;0)')
const { cells } = session.getViewport('Лист1', 'A1:A2', { editText: true })
console.log(cells.A1.text, cells.A2.value, cells.A2.formula, cells.A2.editText) // → 1 234,50 1481 ROUND(A1*1.2,0) =ОКРУГЛ(A1*1,2;0)
console.log(session.setCellText('Лист1', 'A3', '=ОКРУГЛ(A1;2;3)').message) // → Функция ОКРУГЛ принимает 2 аргумента, а указано 3
console.log(session.explain('Лист1', 'A2').text)
// → Лист1!A2 = 1481  ← ОКРУГЛ(A1*1,2;0)
//     Лист1!A1 = 1234,5
```

[Locales](locale.md) has formula text in a locale (`toLocal`, `toInvariant`), what changes in computing and how to add a locale.
