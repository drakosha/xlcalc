# Spilling

A formula whose result is an array of several values fills the cells to the right of and below its own cell, as in Excel 365 with dynamic arrays. The formula's cell is the array's top-left value; the other cells are values of the book that no input holds. Reference: Excel 16.113 for Mac, en-US, dynamic array formulas; every rule below was measured by the spill sets (`bench/coverage/spill/`, answers in `test/excel/spill/`): `base` (50 cases), `order` (21 cases, computed four times), `held` (17 held-out cases whose answers were predicted from the rules of the first two and committed before Excel computed them: all 17 as predicted), and for structure operations the struct sets `spill` (8) and `spill-more` (4) in `test/excel/struct/`.

## The rules

**Where an array goes.** From the formula's cell, as many rows and columns as the array has. A reference of several cells spills the referenced values (`=A1:A4`). An empty element shows 0: `=A1:A4` over a blank A2 puts 0 there, `ISBLANK` of that cell is FALSE and `COUNTBLANK` does not count it. An array of one value is that value, as before (`={5}` is 5).

**What stops it.** The array gives `#SPILL!` and fills nothing when a cell of its range other than the formula's own:

- holds a value, an empty text too (a constant `""`, as xlsx can store one), or a formula, whatever it gives (`""`, an error, a reference to the array's own cell: `A3 = A1+1` stops `A1 = SEQUENCE(3)` and is `#SPILL!` itself);
- belongs to a merged area: one inside the range, one the range cuts, one that holds the formula's own cell (`A1:B1` merged stops `A1 = SEQUENCE(3)`); an area next to the range does not;
- lies past the sheet's edge (`SEQUENCE(3)` in A1048575, `SEQUENCE(1,3)` in XFC1; one row less fits).

A number format or a fill alone does not stop it. A cell typed into the range later stops the array; clearing it lets the array spill again.

**Two arrays over one cell.** When one formula reads the other, the one read spills and the other gives `#SPILL!` (six pairs in the set `order`, both orders, two runs each, and two pairs of `base`). When neither reads the other, Excel's choice varies between runs: the set `order` was computed four times, and of its 15 independent pairs 11 changed their winner at least once (`order-runs.json` keeps every answer seen). The array lower on the sheet won 45 of the 63 times (the four runs and three pairs of `base` run once). The engine is deterministic: the formula another one reads (directly or through other formulas) wins, otherwise the one further down (by sheet, row, column; `src/spill/rank.ts`). Its answer is one Excel gave in every pair of `order`; two pairs of `base`, run once, got the other answer there (known differences in `test/spill-excel.test.ts`).

**Spill references.** `E2#` is the range E2's array spilled into. When E2 holds a formula that did not spill (one value, or `#SPILL!`), `E2#` is E2 alone: `ROWS(E2#)` is 1 and `E2#` gives E2's value, `#SPILL!` included. When E2 holds no formula (a value, nothing, a cell inside another array's range) it is `#REF!`. A spill reference works in functions (`SUM(E2#)`, `ROWS`, `COLUMNS`, `INDEX(E2#,2)`), lifted (`E2#*2` spills itself), from another sheet (`S!E2#`) and in a name (`'S'!$E$2#`).

**Readers.** A formula reading a spilled cell gets its value, a whole column too: `COUNT(B:B)` over `B1 = A:A*2` is 1,048,576.

**Sizes from data.** `SEQUENCE(B1)`, `FILTER`, `UNIQUE`, `SORT` spill their current size; a size that meets a value gives `#SPILL!`, a `FILTER` that keeps nothing `#CALC!` without spilling.

**Whole columns and rows.** `A:A*2` in B1 spills down to row 1,048,576, 0 below the data; from B2 it does not fit; a value anywhere below stops it. `1:1*2` in A2 spills to column XFD.

**Volatile arrays.** `RANDARRAY(3)` spills three draws. Excel gives `#SPILL!` to an array whose size changes between its calculation passes (`SEQUENCE(RANDBETWEEN(1,9))`); the engine draws from its clock once and spills (not measured).

**Structure operations.** `E2#` moves as a reference to its cell: rows and columns inserted or deleted, the anchor cut and pasted, copies (relative and absolute), a renamed sheet (`'R'!E2#`). Deleting the anchor's cell leaves `#REF!#` (`S!#REF!#` from another sheet, in a name too); deleting its sheet `#REF!E2#`. The parser reads these back. Cutting an anchor to another sheet crashed Excel 16.113 twice, so it is not a case.

## Reading Excel's answers

A case is a small book written as Excel 365 saves dynamic array formulas (`cm="1"` and the `XLDAPR` metadata; `E2#` is stored as `_xlfn.ANCHORARRAY(E2)`), one book per case. Reading the `value` of a range with errors is unsafe (errors drop out of the list, a two-dimensional read crashed Excel), so every cell of the rectangles read has three mirror formulas on sheets of their own that never give an error: a type code (`-5` blank, `0` number, `-1`/`-2` TRUE/FALSE, `-3` text, `ERROR.TYPE` for errors), the number, the text. The mirrors are read as ranges, and the rectangles' formulas too (a spilled cell has none). `npx tsx bench/coverage/run.ts spill <set>` runs a set; `test/spill-excel.test.ts` compares without Excel.

## The engine

**Placing an array** (`src/spill/place.ts`). When a formula's value is an array of several values, its range is checked: the edge, the merged areas, the input cells that stop it (an index of the cells with a value or a formula, by column, `src/spill/occupied.ts`), then the other arrays' ranges. An array that cannot spill gives `#SPILL!`, keeps the reason for `issues` ("the array of 3×1 values cannot spill: 'S'!A3 is not empty") and waits for its range: a change of a cell there computes it again. An array that wins over another takes its place, and the other waits.

**The graph.** A formula that read a cell of a spilled range must come after the array's formula. Besides the references, the graph has an edge from each spilled formula to the formulas reading its range (its claim, `src/graph.ts`); the edges change only when a range does. Evaluation goes in rounds (`src/evaluate.ts`): when an array's cells change after a formula read them, or a range an array waits for is freed, those formulas are computed again, until nothing changes; for a book without spills it is one round, as before. A reader of a range that the array's own formula reads closes a cycle: both are `#CYCLE!`, and the formula keeps its claim while in the cycle, so the cycle stays found. What such a cycle holds depends on which arrays spilled first, so a session computes the whole book again from no spills when a cycle runs through a spill range, before or after an edit, as a new computation does (the cost of a cold computation, only for such books).

**Large arrays.** An array of up to 65,536 values is stored whole. A larger one is stored from its top-left up to the last row and column holding a value other than its bottom-right one; the cells past that hold that one value and are not stored (`src/spill/fill.ts`): the runtime reads them where a range or a cell meets them. So `A:A*2` over 1,000 rows stores 1,000 cells and computes in about 16 ms; `SUM(B:B)` and `COUNT(B:B)` over it read 1,048,576 values once (cached, about 50 ms more). The result lists the stored cells (`spillOf`) and gives the formula `spill: 'B1:B1048576'` and `spillFill: 0`; the hash has one line for the rest of the range. An array without such a tail is stored whole: `ROW(A:A)` lists 1,048,576 cells, about 3 s and 0.5 GB, as Excel holds a million cells too.

**Results and sessions.** Spilled cells are in `sheets[].cells` with `spillOf` (the formula's address), the formula has `spill` (its range) and, for a large array, `spillFill`; both are in the values hash, and a book without spills hashes as before. A cell of an array holding an error is in `issues` ("division by zero, in the array spilled from 'S'!B1"). Session changes carry the same fields and list a formula whose range changed even when its own value did not; `setLayout` computes again the arrays whose range a merged area added or removed meets; `explain` of a spilled cell leads to its formula. The contract: `CellResult`, `CellChange`, `ExplainNode` in `src/contract.ts`.

**Speed.** On the synthetic book (no spills) the cold computation, the edits and the hash are as before within noise (medians of four interleaved runs on a loaded machine: cold 287 ms before, 281 ms after; criterion edit p95 1.3 and 1.4 ms; data edit 21.7 and 22.5 ms; heap 30.4 and 31.4 MB, the index of input cells).

## Known differences and limits

- Two arrays over one cell that do not depend on each other: Excel's choice varies between runs; the engine always takes the lower one (two cases of `base` pinned).
- An array whose size depends on a volatile draw spills by the engine's clock; Excel gives `#SPILL!` when the size changes between its passes (not measured).
- A cycle through a spill range is `#CYCLE!`; Excel shows its circular reference warning instead (not measured: the warning is a window).
- Structure operations copy and move the cells of the input: a spilled cell is not input, so copying a range of spilled cells copies their formula's cell only if it is in the range (not measured).
