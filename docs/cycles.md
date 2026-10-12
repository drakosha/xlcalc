# Circular references

A circular reference is a formula that reads itself, directly or through other cells. Without iterative calculation, Excel computes neither the cells of a cycle nor the cells that read one: it keeps the values they show and computes the rest of the book. The engine does the same. With iterative calculation (`iteration`, [Iterative calculation](../README.md#iterative-calculation)) the cycles are computed instead, and nothing below applies.

The rules below were measured in Excel 16 for Windows (en-US). Excel for Mac shows a warning window when it opens such a book, so it was not measured.

## Which cells a cycle stops

A cycle stops each formula in it and each formula that reads a cell of it, directly or through other formulas, wherever they are: above the cycle, below it or on another sheet. A cycle runs through anything a formula refers to: a cell, a range, a defined name, a spilled array's range.

A formula that does not read the cycle is computed as usual, even when it reads a precedent of the cycle or sits next to it. A book without a cycle is computed when it opens, whatever values its file stores.

## What a stopped cell holds

A stopped cell holds a value, the cell's `held`, or 0 without it. It is not computed, so a change of a precedent changes nothing, and neither does a full calculation.

```js
const cells = { A1: { formula: 'B1/2+K1' }, B1: { formula: 'A1' }, C1: { formula: 'A1*10', held: 20 }, K1: { value: 1 } }
compute({ sheets: [{ name: 'S', cells }] })
// A1 0, B1 0, C1 20: none is computed; K1 changed to 5, they stay
```

A file stores the values Excel showed when it saved it, and `readXlsx` gives them as `cached`. `compute(read.book, { cached: read.cached })` and `createSession(read.book, { cached: read.cached })` open the book as Excel opens the file: a stopped cell without a `held` of its own holds the value stored for it, and empty text where the file stores none. Nothing else is read from `cached`.

| the file stores | A1 = B1/2+1, B1 = A1, C1 = A1*10 show |
|---|---|
| 2, 2 and 0 | 2, 2 and 0 |
| 50, 7 and 0 | 50, 7 and 0 |
| 2 for A1 alone | 2, empty text and empty text |
| nothing | empty text in each |

## Edits in a session

A session keeps the stopped cells as Excel keeps them, and writes what each holds into its `held`, so `getBook()` gives it and a full computation of the session's book gives what the session shows.

- An edit that does not break the cycle changes none of them: a precedent changed (C1 from 1 to 5 under A1 = B1/2+C1), a formula elsewhere.
- A formula entered into a stopped cell, or one entered elsewhere that reads the cycle, holds 0, also when it is the formula the cell had (B1 = A1 typed again). Pasting a formula enters it: a copy holds 0.
- A cell that another edit brings into a cycle's way holds what it showed. With A1 = 2, B1 = A1 and C1 = B1*10, typing A1 = B1/2+1 closes the cycle: A1 holds 0, as a formula entered, B1 holds 2 and C1 20.
- A value typed into a cell of the cycle, or the cell cleared, breaks it: each cell it stopped is computed again and holds nothing (B1 = 4 under A1 = B1/2+1 gives A1 3 and C1 30). A cell that also reads another cycle stays stopped.
- Closing the cycle again stops the cells once more: the formula entered holds 0, the others what they show (B1 = A1 typed back: A1 3, B1 0, C1 30).
- A formula changed so that it reads no cycle is computed (C1 = K1*2 gives 6).
- A look (a style, a number format, a note) keeps the value. A row or a column inserted or deleted, or a range moved, takes the held values with their cells; a cell an operation or new names bring into a cycle's way holds what it showed.
- Undo and redo give the held values back, and so does a preview taken back: the book after an undo holds what it held before the edit.

## What the result says

A stopped cell's value is the one it holds, not an error. The result's `diagnostics`, `getCellIssue`, `readCells` with `issue` and a window with `diagnostics` list it with `error: 'CYCLE'` and the reason `circular reference: not computed, the value is held`; its `origin` is the first cell of a cycle its references reach. `explain` says the same in the book's language; of the book above, `explain('S', 'C1', { depth: 1 })` gives:

```
S!C1 = 20  ← A1*10  (held: a circular reference through S!A1)
  S!A1 = 0  ← B1/2+K1  (held: a circular reference)
    S!B1 = 0  ← A1  (held: a circular reference)
```

`#CYCLE!` itself is left for a defined name that uses itself, directly or through other names ([names.md](names.md)).

## Saving

`writeXlsx` stores a stopped cell's value as Excel does: the result's value, and the cell's `held` where the result gives none, also with `values: 'omit'`. Excel opens the file holding it, and so does `compute` with the `cached` of the file read back.

## Arrays

An array formula a cycle stops spills nothing and holds one value. A cycle through a spill range (`A1 = SEQUENCE(3)*B5` with `B5 = SUM(A1:A3)`) holds the stored values, as Excel does: A1 holds its stored 0 and A2:A3 are empty.

An array that spilled and that an edit brings into a cycle's way holds its first value here and spills no more, as `held` is one value a cell. Excel keeps its whole spill as it was: with `A1 = SEQUENCE(3)+D1` and D1 = A3*0 typed, Excel keeps 1, 2 and 3 in A1:A3, the engine 1 in A1 and nothing below.

## How it was checked

The workbook step of the Windows harness (`bench/windows`, `cycle` cases) opens each book through COM, takes the case's steps (a value or a formula typed, a full calculation, a recalculation, a save) and reads the cells after each; the file of a case stores the values it lists, as Excel stores what it computed (`test/excel/win/workbook/`).

| set | books | what | the engine |
|---|---:|---|---|
| `cycles` | 21 | a cycle typed into a book, the values a file stores (converged ones, others, none), a precedent changed, a formula typed again, the link broken and typed back, a cycle of one cell, through INDIRECT, OFFSET, a name and a spill, cells beside a cycle (above it, on another sheet, typed after opening, a chain closed), a file with `calcPr iterate` | all but the cycles through INDIRECT and OFFSET (not in the whitelist) and the iteration from stored values (`iterate-file`, [Limitations](../README.md#limitations)) |
| `cycles-held` | 6 | predicted by the engine before Excel ran them: an array brought into a cycle's way, a formula typed again as it was, a cell cleared, a dependent typed out of the way, two cycles of which one is broken, a file storing values for part of a cycle | 125 of the 129 cells read as predicted; the four others are the spill an array a cycle comes to keeps in Excel |

`test/workbook-win.test.ts` compares the engine with the stored answers without Excel. `test/session-cycles.test.ts` takes the measured scenarios through each kind of session (the same thread synchronously and in portions, a thread of the pool) in both execution modes, with undo, redo, previews, typed text, a look, a row inserted and a copy, and after each step checks the session's hash against a full computation of its book. The seeded session tests do the same after each of their edits.
