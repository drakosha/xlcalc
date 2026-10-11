# Spilling

When a formula returns several values, they fill the cells to the right of and below the formula's cell, as in Excel 365. The formula's cell holds the top-left value. The other cells hold values too, but no input cell stands behind them.

The rules below were measured in Excel 16.113 for Mac (en-US) with dynamic array formulas.

## Where the values go

An array takes as many rows and columns from the formula's cell as it has. A reference to several cells spills their values: `=A1:A4` copies A1:A4.

An empty element shows 0. If A2 is blank, `=A1:A4` puts 0 in the second cell, `ISBLANK` of that cell is FALSE, and `COUNTBLANK` does not count it.

An array of one value is just that value: `={5}` is 5.

## What blocks a spill

The formula gives `#SPILL!` and fills nothing when a cell of its range, other than its own, is in the way. A cell is in the way when it:

- holds a value, including an empty string stored as a constant;
- holds a formula, whatever it returns: `""`, an error, or a reference back to the array's own cell (with `A1 = SEQUENCE(3)` and `A3 = A1+1`, both cells are `#SPILL!`);
- belongs to a merged area inside the range, cut by the range, or holding the formula's own cell (`A1:B1` merged blocks `A1 = SEQUENCE(3)`); a merged area next to the range does not block;
- lies past the edge of the sheet (`SEQUENCE(3)` in A1048575 does not fit, in A1048574 it does);
- belongs to the area of an array formula of the old kind ([Formulas of the old kind](legacy.md)).

A number format or a fill alone does not block. Typing a value into the range blocks the array, and clearing it lets the array spill again.

## Two arrays over the same cells

If one formula reads the other, the one being read spills and the other gives `#SPILL!`.

If neither reads the other, Excel's choice changes from one calculation to the next: computed four times, 11 of 15 such pairs changed their winner at least once, and the lower array won 45 times out of 63. The engine has to be deterministic, so it lets the lower array win (by sheet, then row, then column).

## Spill references

`E2#` is the range that E2's array fills.

| E2 holds | `E2#` is |
|---|---|
| a formula that spilled | its spill range |
| a formula with one value, or `#SPILL!` | E2 alone (`ROWS(E2#)` is 1, the value is E2's, `#SPILL!` included) |
| a value, nothing, or a cell of another array | `#REF!` |

A spill reference works in functions (`SUM(E2#)`, `ROWS`, `COLUMNS`, `INDEX(E2#,2)`), under operators (`E2#*2` spills itself), from another sheet (`S!E2#`) and in a defined name (`'S'!$E$2#`). In an xlsx file it is stored as `_xlfn.ANCHORARRAY(E2)`.

## Readers and sizes

A formula that reads a spilled cell gets its value. `COUNT(B:B)` over `B1 = A:A*2` is 1,048,576.

`SEQUENCE(B1)`, `FILTER`, `UNIQUE` and `SORT` spill their current size. A larger size that meets a value gives `#SPILL!`. A `FILTER` that keeps nothing gives `#CALC!` and spills nothing.

`A:A*2` in B1 spills to row 1,048,576, with 0 below the data. From B2 it does not fit, and any value below B1 blocks it. `1:1*2` in A2 spills to column XFD.

`RANDARRAY(3)` spills three draws. When the size itself is drawn (`SEQUENCE(RANDBETWEEN(1,3))`), Excel may give `#SPILL!` because the size changes between its calculation passes. A single such formula spilled in five runs out of five, but ten of them in one book gave `#SPILL!` 47 times out of 50. The engine draws once from its clock and spills.

## Edits and structure operations

In a session, a value typed into a spill range gives the array `#SPILL!`, and its readers (`SUM(A1#)`, `A1#*2`) follow. Clearing the value lets the array spill again. Merged areas beside the range, below it or holding the formula's cell do not change this.

Spilled cells are not cells of the book, so copying them without their formula pastes nothing. Copying the formula pastes it, and it spills at the target. Moving spilled cells moves the references to them (`E3+1` follows `E3`), and moving the formula moves `E2#`.

`E2#` changes like a reference to E2: inserted and deleted rows and columns move it, copies shift it unless it is absolute, renaming the sheet renames it. Deleting the formula's cell leaves `#REF!#` (`S!#REF!#` from another sheet, in a name too), and deleting its sheet leaves `#REF!E2#`. The parser reads these forms back.

## Results and sessions

In `ComputeResult`, spilled cells appear in `sheets[].cells` with `spillOf`, the address of their formula. The formula's cell has `spill`, its range. Spilled values are part of the values hash. A book without spills hashes as it did before spilling existed.

An error inside a spilled array is listed in the result's `diagnostics`, for example "division by zero, in the array spilled from 'S'!B1". Session changes carry the same fields and include a formula whose range changed even when its own value did not. `setLayout` recomputes the arrays whose range meets a merged area that was added or removed. `explain` of a spilled cell leads to its formula.

## Large arrays

An array of up to 65,536 values is stored whole. A larger one is stored up to the last row and column whose values differ from its bottom-right value; the rest of the range holds that one value and is not stored. Such a formula has `spillFill` with the value, and the hash has one line for the unstored part.

So `A:A*2` over 1,000 rows of data stores 1,000 cells and computes in about 16 ms, and `SUM(B:B)` over it reads the million values once, in about 50 ms. An array without such a tail is stored whole: `ROW(A:A)` lists 1,048,576 cells and takes about 3 s and 0.5 GB, as Excel also holds a million cells; a thread pool with a heap under 1 GB refuses such a spill by its default `limits.spill` ([async.md](async.md)).

## How the engine does it

`src/spill/place.ts` places an array: it checks the sheet's edge, the merged areas, the input cells in the way (an index of cells with a value or a formula, `src/spill/occupied.ts`) and the ranges of other arrays. An array that cannot spill gives `#SPILL!`, records which cell is in the way, and is computed again when that cell changes.

A formula that reads a cell of a spilled range has to run after the array's formula. The graph (`src/graph.ts`) adds an edge from each spilled formula to the formulas reading its range, and changes these edges only when a range changes. Evaluation (`src/evaluate.ts`) goes in rounds until nothing changes; a book without spills takes one round.

A formula that reads a range its own array's formula also reads closes a cycle, and both are `#CYCLE!`. What such a cycle holds depends on which arrays spilled first, so a session recomputes the whole book from scratch when a cycle runs through a spill range. This costs as much as a cold computation and happens only in such books.

On the synthetic book, which has no spills, cold computation, edits and the hash are unchanged within noise.

## How it was checked

Each case is a small book written the way Excel 365 saves dynamic array formulas (`cm="1"` and the `XLDAPR` metadata). Reading a range of values that contains errors is unsafe in Excel (errors drop out of the list, and a two-dimensional read crashed Excel), so every cell that is read has three mirror formulas on separate sheets that never fail: a type code, the number and the text. `npx tsx bench/coverage/run.ts spill <set>` runs a set, and `test/spill-excel.test.ts` compares the engine with the stored answers without Excel.

| set | cases | what |
|---|---|---|
| `base` | 50 | directions, blocking, collisions, edges, `E2#`, sizes from data, whole columns, volatile arrays |
| `order` | 21 | two arrays over the same cells, computed four times |
| `held` | 17 | predicted by the engine before Excel computed them; all 17 matched |
| `more` | 13 | sizes from a draw (five runs), edits in books with merged areas |
| struct `spill`, `spill-more`, `spill-copy` | 22 | `E2#` and spilled cells under structure operations |

Answers are in `test/excel/spill/` and `test/excel/struct/`. Cutting a formula with a spill to another sheet crashed Excel 16.113 twice, so it is not among the cases.

## Known differences

- Two independent arrays over the same cells: Excel's choice varies between runs, the engine always lets the lower one spill. Two cases of `base` that Excel ran once show the other choice.
- An array whose size depends on a random draw spills by the engine's clock. Excel gives `#SPILL!` when the size changes between its passes (three cases of `more`).
- A cycle through a spill range is `#CYCLE!`. Excel shows its circular reference warning instead (not measured: the warning is a window).
