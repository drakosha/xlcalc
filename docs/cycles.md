# Circular references

A circular reference is a formula that reads itself, directly or through other cells. Without iterative calculation, Excel calculates neither the cells of a cycle nor the cells that depend on one: it keeps the values they display and calculates the rest of the workbook. The engine does the same. With iterative calculation (`iteration`, [Iterative calculation](#iterative-calculation-in-short)), the cycles are calculated instead, and nothing below applies.

The rules below were verified in Excel 16 for Windows (en-US). Excel for Mac shows a warning dialog when it opens such a workbook, so it was not tested.

## In short

In a workbook without iterative calculation, a formula in a circular reference, and every formula that depends on one, is not calculated, just as Excel calculates none of them: each keeps its value, the cell's `held`, or 0 without it, and the rest of the workbook is calculated. `cached`, the values a file stores (`readXlsx`'s `cached`), opens such a workbook as Excel opens the file: a cell keeps the value stored for it, or empty text where the file stores none. A session keeps the values as Excel keeps them: an edit that does not break the cycle changes none of them, a formula entered there keeps 0, a cell that another edit pulls into the cycle keeps what it displayed, and a value typed into a cell of the cycle breaks it, and everything is calculated again. `getBook()` returns the values in `held`, so a full calculation of the session's workbook produces what the session displays, and `writeXlsx` stores them. Each such cell appears in `diagnostics` with `error: 'CYCLE'`. The rules and how they were verified: [docs/cycles.md](cycles.md).

```js
import { compute, createSession } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { formula: 'B1/2+1' }, B1: { formula: 'A1' }, C1: { formula: 'A1*10' }, K1: { value: 3 }, E1: { formula: 'K1*2' } } }] }
const cached = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, B1: { value: 2 } } }] }
const r = compute(book, { cached })
console.log(r.sheets[0].cells.A1.value, r.sheets[0].cells.C1.value, r.sheets[0].cells.E1.value) // → 2  6
console.log(r.diagnostics.map((d) => `${d.cell} ${d.error}`), r.diagnostics[0].reason) // → [ 'A1 CYCLE', 'B1 CYCLE', 'C1 CYCLE' ] circular reference: not calculated, the value is held

const session = createSession(book, { cached })
session.setCells([{ sheet: 'S', cell: 'B1', input: { value: 4 } }])
console.log(session.getCell('S', 'A1').value, session.getCell('S', 'C1').value) // → 3 30
session.setCells([{ sheet: 'S', cell: 'B1', input: { formula: 'A1' } }])
console.log(session.getCell('S', 'B1').value, session.getBook().sheets[0].cells.C1) // → 0 { formula: 'A1*10', held: 30 }
```

## Iterative calculation in short

With `iteration`, which corresponds to Excel's "Enable iterative calculation" setting as a file stores it, the cells of every cycle and the cells that depend on them are calculated repeatedly, starting from 0, as Excel calculates a workbook it opens, and `held` has no effect. The iterations stop after `maxIterations` (100 by default) or when no number changes by `maxChange` (0.001 by default) or more.

```js
import { compute } from 'xlcalc'

const cells = { B1: { value: 1000 }, B2: { value: 0.1 }, B3: { formula: 'B2*(B1+B4)/2' }, B4: { formula: 'B1+B3-100' } }
const r = compute({ sheets: [{ name: 'S', cells }], iteration: {} }).sheets[0].cells
console.log(r.B3.value, r.B4.value) // → 99.999984375 999.999984375
console.log(compute({ sheets: [{ name: 'S', cells }] }).sheets[0].cells.B3.value) // → 0
```

Cells are calculated sheet by sheet in the order of the sheet names, row by row. A change of text, a boolean or an error into another one does not count as a change. Every calculation starts from 0, so a session produces what a full calculation produces. An array in a cycle can spill into cells that a formula outside the cycles reads and give them back by the last iteration (a `FILTER` over its own spill that a value further down stops); the cycles are then calculated again for that formula, and once that ends where the time before ended, the calculation stops with those values. Excel continues from the values it holds when it recalculates, which the engine does not do (see [Limitations](limitations.md)).

## Which cells a cycle stops

A cycle stops each formula in it and each formula that reads a cell of it, directly or through other formulas, wherever they are: above the cycle, below it or on another sheet. A cycle runs through anything a formula refers to: a cell, a range, a defined name, a spilled array's range.

A formula that does not read the cycle is calculated as usual, even when it reads a precedent of the cycle or sits next to it. A workbook without a cycle is calculated when it opens, whatever values its file stores.

## What a stopped cell holds

A stopped cell holds a value, the cell's `held`, or 0 without it. It is not calculated, so changing a precedent changes nothing, and neither does a full calculation.

```js
import { compute, createSession } from 'xlcalc'

const cells = { A1: { formula: 'B1/2+K1' }, B1: { formula: 'A1' }, C1: { formula: 'A1*10', held: 20 }, K1: { value: 1 } }
const r = compute({ sheets: [{ name: 'S', cells }] }).sheets[0].cells
console.log(r.A1.value, r.B1.value, r.C1.value)
// → 0 0 20
// none is computed: K1 changed to 5, they stay
const session = createSession({ sheets: [{ name: 'S', cells }] })
session.setCells([{ sheet: 'S', cell: 'K1', input: { value: 5 } }])
console.log(session.getCell('S', 'A1').value, session.getCell('S', 'C1').value)
// → 0 20
```

A file stores the values Excel displayed when it saved the file, and `readXlsx` returns them as `cached`. `compute(read.book, { cached: read.cached })` and `createSession(read.book, { cached: read.cached })` open the workbook as Excel opens the file: a stopped cell without a `held` of its own keeps the value stored for it, or empty text where the file stores none. Nothing else is read from `cached`.

| the file stores | A1 = B1/2+1, B1 = A1, C1 = A1*10 display |
|---|---|
| 2, 2 and 0 | 2, 2 and 0 |
| 50, 7 and 0 | 50, 7 and 0 |
| 2 for A1 only | 2, empty text and empty text |
| nothing | empty text in each |

## Edits in a session

A session keeps the stopped cells as Excel keeps them, and writes what each one holds into its `held`, so `getBook()` returns it and a full calculation of the session's workbook produces what the session displays.

- An edit that does not break the cycle changes none of them: a changed precedent (C1 from 1 to 5 under A1 = B1/2+C1), or a formula elsewhere.
- A formula entered into a stopped cell, or one entered elsewhere that reads the cycle, holds 0, even when it is the formula the cell already had (B1 = A1 typed again). Pasting a formula enters it: a copy holds 0.
- A cell that another edit pulls into a cycle keeps what it displayed. With A1 = 2, B1 = A1 and C1 = B1*10, typing A1 = B1/2+1 closes the cycle: A1 holds 0, as a newly entered formula, B1 holds 2 and C1 20.
- A value typed into a cell of the cycle, or clearing the cell, breaks it: each cell it stopped is calculated again and holds nothing (B1 = 4 under A1 = B1/2+1 gives A1 3 and C1 30). A cell that also depends on another cycle stays stopped.
- Closing the cycle again stops the cells once more: the entered formula holds 0 and the others what they display (B1 = A1 typed back: A1 3, B1 0, C1 30).
- A formula changed so that it no longer reads a cycle is calculated (C1 = K1*2 gives 6).
- Formatting (a style, a number format, a note) keeps the value. Inserting or deleting a row or a column, or moving a range, takes the held values along with their cells; a cell that an operation or new names pull into a cycle keeps what it displayed.
- Undo and redo restore the held values, and so does a preview that is discarded: the workbook after an undo holds what it held before the edit.

## What the result reports

A stopped cell's value is the value it holds, not an error. The result's `diagnostics`, `getCellIssue`, `readCells` with `issue` and a viewport with `diagnostics` list it with `error: 'CYCLE'` and the reason `circular reference: not calculated, the value is held`; its `origin` is the first cell of a cycle its references reach. `explain` reports the same in the workbook's language; for the workbook above, `explain('S', 'C1', { depth: 1 })` returns:

```
S!C1 = 20  ← A1*10  (held by a circular reference through S!A1)
  S!A1 = 0  ← B1/2+K1  (held by a circular reference)
    S!B1 = 0  ← A1  (held by a circular reference)
```

`#CYCLE!` itself is reserved for a defined name that refers to itself, directly or through other names ([names.md](names.md)).

## Saving

`writeXlsx` stores a stopped cell's value as Excel does: the result's value, or the cell's `held` where the result has none, even with `values: 'omit'`. Excel opens the file with that value, and so does `compute` with the `cached` of the file read back.

## Arrays

An array formula that a cycle stops spills nothing and holds a single value. A cycle through a spill range (`A1 = SEQUENCE(3)*B5` with `B5 = SUM(A1:A3)`) keeps the stored values, as Excel does: A1 keeps its stored 0 and A2:A3 are empty.

An array that has spilled and that an edit pulls into a cycle keeps only its first value here and no longer spills, since `held` is a single value per cell. Excel keeps its entire spill as it was: with `A1 = SEQUENCE(3)+D1` and D1 = A3*0 typed, Excel keeps 1, 2 and 3 in A1:A3, while the engine keeps 1 in A1 and nothing below.

## Checked against Excel

Tested in Excel 16 for Windows, en-US, on 27 workbooks, 6 of them predicted before Excel ran them. The engine matches all of them except the cycles through `INDIRECT` and `OFFSET`, which are not in the whitelist, and iteration from stored values ([Limitations](limitations.md)).
