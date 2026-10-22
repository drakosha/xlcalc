# Spilling

When a formula returns several values, they fill the cells to the right of and below the formula's cell, as in Excel 365. The formula's cell holds the top-left value. The other cells hold values too, but have no input of their own.

The rules below were verified in Excel 16.113 for Mac (en-US) with dynamic array formulas.

```js
import { compute, createSession } from 'xlcalc'

const book = {
  sheets: [{
    name: 'S',
    cells: { A1: { value: 3 }, B1: { formula: 'SEQUENCE(A1)*10' }, C1: { formula: 'SUM(B1#)' }, D1: { formula: 'ROWS(B1#)' } },
  }],
}
const { cells } = compute(book).sheets[0]
console.log([cells.B1.value, cells.B2.value, cells.B3.value], cells.B1.spill, cells.B3.spillOf)
// → [ 10, 20, 30 ] B1:B3 B1
console.log(cells.C1.value, cells.D1.value)
// → 60 3

// a value in the way: #SPILL!, and its readers follow; cleared, the array spills again
const session = createSession(book)
session.setCells([{ sheet: 'S', cell: 'B3', input: { value: 'x' } }])
console.log(session.getCell('S', 'B1').value, session.getCell('S', 'C1').value)
// → { error: 'SPILL' } { error: 'SPILL' }
session.setCells([{ sheet: 'S', cell: 'B3', input: null }])
console.log(session.getCell('S', 'C1').value)
// → 60
```

## Where the values go

An array takes as many rows and columns from the formula's cell as it has. A reference to several cells spills their values: `=A1:A4` copies A1:A4.

An empty element displays 0. If A2 is blank, `=A1:A4` puts 0 in the second cell, `ISBLANK` of that cell is FALSE, and `COUNTBLANK` does not count it.

An array with one value is just that value: `={5}` is 5.

## What blocks a spill

The formula returns `#SPILL!` and fills nothing when a cell of its range, other than its own, is in the way. A cell is in the way when it:

- holds a value, including an empty string stored as a constant;
- holds a formula, whatever it returns: `""`, an error, or a reference back to the array's own cell (with `A1 = SEQUENCE(3)` and `A3 = A1+1`, both cells are `#SPILL!`);
- belongs to a merged area inside the range, cut by the range, or containing the formula's own cell (`A1:B1` merged blocks `A1 = SEQUENCE(3)`); a merged area next to the range does not block;
- lies beyond the edge of the sheet (`SEQUENCE(3)` in A1048575 does not fit, in A1048574 it does);
- belongs to the range of a legacy array formula ([Legacy array formulas and implicit intersection](legacy.md)).

A number format or a fill alone does not block. Typing a value into the range blocks the array, and clearing it lets the array spill again.

## Two arrays over the same cells

If one formula reads the other, the one being read spills and the other returns `#SPILL!`.

If neither reads the other, Excel's choice changes from one calculation to the next: calculated four times, 11 of 15 such pairs changed their winner at least once, and the lower array won 45 times out of 63. The engine has to be deterministic, so it lets the lower array win (by sheet, then row, then column).

## Spill references

`E2#` is the range that E2's array fills.

| E2 holds | `E2#` is |
|---|---|
| a formula that spilled | its spill range |
| a formula with one value, or `#SPILL!` | E2 alone (`ROWS(E2#)` is 1, and the value is E2's, `#SPILL!` included) |
| a value, nothing, or a cell of another array | `#REF!` |

A spill reference works in functions (`SUM(E2#)`, `ROWS`, `COLUMNS`, `INDEX(E2#,2)`), with operators (`E2#*2` spills itself), from another sheet (`S!E2#`) and in a defined name (`'S'!$E$2#`). In an xlsx file it is stored as `_xlfn.ANCHORARRAY(E2)`.

## Dependents and sizes

A formula that reads a spilled cell gets its value. `COUNT(B:B)` over `B1 = A:A*2` is 1,048,576.

`SEQUENCE(B1)`, `FILTER`, `UNIQUE` and `SORT` spill their current size. A larger size that runs into a value returns `#SPILL!`. A `FILTER` that keeps nothing returns `#CALC!` and spills nothing.

`A:A*2` in B1 spills to row 1,048,576, with 0 below the data. From B2 it does not fit, and any value below B1 blocks it. `1:1*2` in A2 spills to column XFD.

`RANDARRAY(3)` spills three draws. When the size itself is random (`SEQUENCE(RANDBETWEEN(1,3))`), Excel may return `#SPILL!` because the size changes between its calculation passes. A single such formula spilled in five runs out of five, but ten of them in one workbook returned `#SPILL!` 47 times out of 50. The engine draws once from its clock and spills.

## Edits and structure operations

In a session, a value typed into a spill range makes the array return `#SPILL!`, and the formulas that read it (`SUM(A1#)`, `A1#*2`) follow. Clearing the value lets the array spill again. Merged areas beside the range, below it or containing the formula's cell do not change this.

Spilled cells are not cells of the workbook, so copying them without their formula pastes nothing. Copying the formula pastes it, and it spills at the destination. Moving spilled cells moves the references to them (`E3+1` follows `E3`), and moving the formula moves `E2#`.

`E2#` changes like a reference to E2: inserted and deleted rows and columns move it, copies shift it unless it is absolute, and renaming the sheet renames it. Deleting the formula's cell leaves `#REF!#` (`S!#REF!#` from another sheet, and in a name too), and deleting its sheet leaves `#REF!E2#`. The parser reads these forms back.

## Results and sessions

In `ComputeResult`, spilled cells appear in `sheets[].cells` with `spillOf`, the address of their formula. The formula's cell has `spill`, its range. Spilled values are part of the values hash. A workbook without spills hashes as it did before spilling was implemented.

An error inside a spilled array is listed in the result's `diagnostics`, for example "division by zero, in the array spilled from 'S'!B1". Session changes carry the same fields and include a formula whose range changed even when its own value did not. `setLayout` recalculates the arrays whose range overlaps a merged area that was added or removed. `explain` of a spilled cell leads to its formula.

## Large arrays

An array of up to 65,536 values is stored in full. A larger one is stored up to the last row and column whose values differ from its bottom-right value; the rest of the range holds that one value and is not stored. Such a formula has `spillFill` with the value, and the hash has one line for the unstored part.

So `A:A*2` over 1,000 rows of data stores 1,000 cells and calculates in about 16 ms, and `SUM(B:B)` over it reads the million values once, in about 50 ms. An array without such a tail is stored in full: `ROW(A:A)` lists 1,048,576 cells and takes about 3 s and 0.5 GB, just as Excel also holds a million cells; a thread pool with a heap under 1 GB rejects such a spill by its default `limits.spill` ([async.md](async.md)).

## How the engine does it

`src/spill/place.ts` places an array: it checks the edge of the sheet, the merged areas, the input cells in the way (an index of cells with a value or a formula, `src/spill/occupied.ts`) and the ranges of other arrays. An array that cannot spill returns `#SPILL!`, records which cell is in the way, and is recalculated when that cell changes.

A formula that reads a cell of a spilled range has to run after the array's formula. The graph (`src/graph.ts`) adds an edge from each spilled formula to the formulas that read its range, and changes these edges only when a range changes. Evaluation (`src/evaluate.ts`) runs in rounds until nothing changes; a workbook without spills takes one round.

A formula that reads a range which its own array's formula also reads closes a cycle, and neither is calculated: both keep their values ([cycles.md](cycles.md)), and the array no longer spills. Where two arrays want the same cells, one whose formula reads, itself or through the formulas it reads, a cell below or to the right of it on its sheet (a cell its spill may cover) gives way to one whose formula reads none, as Excel decides for a workbook opened from its file where such an array closes a cycle through its spill; between two that both read such a cell, or neither, the order above decides. Both tests look at the formulas and their references (a spill reference as far as its array spills), not at which of the two arrays spilled first, so the cells and edits of a workbook coming in another order do not change which of them spills. A session recalculates the whole workbook from scratch when a cycle runs through a spill range, so it shows what a full calculation of its workbook shows. This costs as much as a cold calculation and happens only in such workbooks. Without iterative calculation a cycle holds the values its cells had when it formed ([cycles.md](cycles.md)): where the size of an array in the cycle follows those values, the outcome can still depend on the order the cells were entered in.

On the synthetic workbook, which has no spills, cold calculation, edits and the hash are unchanged within noise.

## Checked against Excel

Tested in Excel 365 on 104 workbooks; the last 17 cases were predicted before Excel calculated them, and all came out as predicted.

## Known differences

- Two independent arrays over the same cells: Excel's choice varies between runs, while the engine always lets the lower one spill. Two cases of `base` that Excel ran once show the other choice.
- An array whose size depends on a random draw spills according to the engine's clock. Excel returns `#SPILL!` when the size changes between its passes (three cases of `more`).
- Arrays in a cycle through their own spill beside another array over the same cells: Excel, entering the formulas one by one, keeps the outcome their order gave it, through a full calculation too; the engine decides by the formulas, as Excel does for the workbook opened from its file (the workbook sets `spill-cycles-held`).
- A cycle through a spill range stops its cells as any cycle does ([cycles.md](cycles.md)). An array that a cycle reaches after an edit no longer spills here and keeps its first value; Excel for Windows keeps its entire spill.
