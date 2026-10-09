# Range and lookup caches

The runtime keeps large range values and SUMIFS/COUNTIFS indexes in `RangeCache` (`src/range-cache.ts`). Dependency listeners and range caches share `AreaIndex` (`src/area-index.ts`): narrow ranges are indexed by sheet and column, wide ranges in a separate per-sheet bucket. A write checks the entries in its columns and that sheet's wide ranges, rather than parsing and checking every cached key. The bounds are the referenced area, not just its occupied cells: writing below the occupied part of `A:A` invalidates it too. Deletion and clearing remove the entries from both the values and the column index. The graph and spill occupancy also share the sorted-row lower-bound search in `src/sorted.ts`.

`test/range-index.test.ts` compares both graph readers and cache invalidation to a full rectangle scan, including multiple sheets, the narrow/wide boundary, whole rows and columns, and deletion during traversal.

The runtime also keeps each cached value array's materialized shape in a WeakMap. A write outside the referenced range can grow the sheet without invalidating that array. Reads keep the snapshot's original height and row stride; newly occupied sheet bounds cannot be applied to an older, shorter array. The unstored part of the range remains blank. `test/range-shape.test.ts` checks growth in both dimensions and restoration of a blocked whole-column spill, comparing session and full results.

Exact VLOOKUP, HLOOKUP, MATCH, XLOOKUP and XMATCH searches reuse indexes over immutable range or array values (`src/functions/lookup/exact.ts`). The storage array, its stride and length, and the book's comparator identify an index. An edit invalidates the range array; its old lookup indexes are held weakly and can be collected with it.

A scalar index keeps the first and last positions of each number or boolean. Numbers compare exactly, as lookups do, without the output's 15-digit rounding; booleans and numbers stay distinct. A text index groups spellings with the book's comparator, preserving case-insensitive, accent-sensitive equality, including Unicode-equivalent spellings. Lower-cased strings alone are not equality keys. First and last positions preserve duplicate direction.

Small lines, wildcard patterns and approximate searches use the existing scan or binary search. Blank lookup keys keep each function's existing rules. Indexes are built lazily after enough positions have actually been scanned: one line's length for scalar keys, four for text. Repeated hits near the start therefore do not force a full index build. Index construction checks the computation budget.

Index construction runs inside one formula's evaluation, with budget checks; it does not yield to the event loop inside that formula.

`test/range-cache.test.ts` checks invalidation across sheets, columns, wide ranges and the unoccupied part of a whole column. `test/lookup-index.test.ts` checks scan work, types, duplicate direction, Unicode, wildcards, horizontal ranges, edits, structure changes and spills, with closures and eval. The existing Excel, session and worker tests also exercise these paths.

## Measurements

Five alternating pairs against `a33993d`, Node 26.7 on the development Mac, 2026-10-09. The clock was `2026-10-04T12:00:00Z`, seed 7. Session benchmarks used closures and `display: false`; result comparisons, including both hashes, were outside the measured calls. Times are medians in milliseconds and depend on machine load.

| Book and operation | Before | After |
|---|---:|---:|
| 3,000 rows: A holds 1, B holds `SUM($A$1:A<row>)`; open session | 1,018 | 86 |
| Same book: change A1500 to 2 | 664 | 49 |
| 34,000 shuffled numeric keys, 1,000 VLOOKUP formulas; open session | 314 | 54 |
| Same lookup book: replace all 1,000 query keys in one setCells call | 138 | 12 |
| Same lookup book with text keys; open session | 1,270 | 81 |
| Text keys: replace all 1,000 query keys | 638 | 12 |
| synthBook, seed 1, 3,400 rows: full compute with display and hashes | 263 | 259 |
| synthBook, seed 1, 34,000 rows: full compute with display and hashes | 2,405 | 2,360 |

Eval also improved on the prefix and numeric lookup books, with the same values and hashes. A separate comparison with the frozen baseline matched 140 results across both execution modes, both locales, mixed lookup types, edits and structure operations.

## Shared aggregate prefixes

Single-range SUM, COUNT, MIN and MAX use `src/range-aggregate.ts`. One cached state records the unrounded running sum, number count, minimum, maximum and first reference error. Shorter row-major prefixes can feed longer ones. Checkpoints about every 256 cells also let a shorter range reuse work when the longest range was evaluated first. Each extension copies the state before adding cells, so it cannot alter a cached prefix.

Addition stays in the original row-major order. Intermediate sums are not fitted to Excel's number range: premature rounding would lose subnormal contributions or change the error returned after overflow. Horizontal ranges extend by columns; taller ranges extend by whole rows of the same referenced width. Extending a multi-row range by columns would reorder additions and is not reused.

An edit invalidates states through the same range index as the value cache. Spills use that invalidation path too. The states do not allocate a flat array of all preceding values for every prefix. Multi-argument calls, arrays and unions use the existing argument path. A single range is never made into an array here, so `limits.cells`, which bounds the values of one range read whole, does not bound it, and a pool with `maxHeapMb` takes this path too. Through the argument path a running total over 10,000 rows (`SUM($A$1:A<row>)`) held an array for every prefix and filled a thread of 384 MB; now it opens in such a pool.

`test/range-aggregate.test.ts` checks read counts in descending evaluation order, addition order, raw underflow/overflow, error precedence, edits, structure operations, horizontal ranges, spills and limits. A separate frozen-build comparison checked 312 results across execution modes and locales, including values, issues and both hashes. That comparison also exposed a pre-existing shape mismatch in cached ranges after sheet growth, addressed separately.

Five alternating pairs against the statistics checkpoint `71fd5e6`, same machine, clock and options as above, with tests stopped. These changes are additional to the earlier cache invalidation and lookup indexes:

| Book and operation | Before | After |
|---|---:|---:|
| 5,000 prefix SUM formulas; open session, closures | 166 ms | 35 ms |
| Same book: change A2500 to 2 | 85 ms | 9.3 ms |
| Same book: open session, eval | 170 ms | 41 ms |
| Same book: edit, eval | 82 ms | 9.3 ms |
| synthBook, seed 1, 3,400 rows; full compute | 227 ms | 227 ms |
| synthBook, seed 1, 34,000 rows; full compute | 1,569 ms | 1,535 ms |

The synthetic business book did not show a consistent speed improvement from shared aggregate prefixes. Before the cached-shape fix, a separate paired run at 34,000 rows gave 1,522/1,543 ms, so the sign of that small difference changed between runs. Edit-A p95 was 1.13/1.01 ms at 3,400 rows and 3.62/3.56 ms at 34,000 rows (before/after).

Separate fresh-process measurements of the 5,000-prefix book retained about 102 MiB before and 7.2 MiB after, measured as heap growth after GC while the session was open, median of five runs. The new implementation opened 34,000 prefix formulas in 194 ms and edited the middle input in 62 ms, retaining about 53 MiB. The old implementation was not run at that size because its separate prefix arrays contain roughly 578 million values.

After sharing `AreaIndex` between the graph and caches, five alternating pairs against `9e6dd70` gave 36.8/35.3 ms for opening 5,000 prefix formulas and 9.6/9.4 ms for the middle edit. Synthetic full compute was 212/210 ms at 3,400 rows and 1,525/1,530 ms at 34,000 rows; edit-A p95 at 34,000 rows was 3.56/3.57 ms. These measurements show no material change in those workloads. A separate comparison against the same frozen build matched 856 complete results and both hashes across both execution modes and locales, including edits, structure operations, spills, numerical edge cases and aggregate function families.

## Cache budget

What the range caches hold is bounded, so a book whose formulas read many large ranges no longer fills the heap with copies of them. Running means, counts and searches down a column (`AVERAGE($A$1:A<row>)`, `COUNTA`, `COUNTIF($A$1:A<row>,">0")`, `MATCH(<row>,$A$1:A<row>,0)`) held an array for every prefix: over 10,000 rows they ran a pool thread of 384 MB out of memory or took its process down, and over 34,000 rows a synchronous session took a Node process of 4 GiB down. They now open and edit in both.

A range's values are kept from its second read on, counted since the range last changed; the first read leaves only a mark. A running mean reads each prefix once, so it fills no cache. The criteria indexes of SUMIFS, COUNTIFS and their kin are kept the same way. A range that one formula reads twice (`AVERAGE(x)+COUNTA(x)`), or that several formulas read, is kept as before.

The kept values and indexes count against one budget in bytes, shared by the books of one heap: the sessions and computations of a pool's thread, or those of the main thread. A range's values count 8 bytes each, a criteria index 16 bytes a position and 80 a distinct value, and each entry or mark 256 bytes more. On Node these estimates came within 13 % of the heap the caches retained, with numbers, text and mixed values. Between two formulas, when the caches hold more than the budget, the entries read least recently are dropped until three quarters of the budget remain. The entries the last formula read stay, so a range larger than the budget that every formula reads is not made again for each. A formula's ranges are not dropped while it computes, so each keeps one shape during it. A dropped range is read again when a formula needs it, so the budget changes times, never values or hashes.

In a pool with `maxHeapMb` the budget is a quarter of a thread's heap: 96 MiB for 384 MB. Elsewhere it is 512 MiB per heap, since the engine reads no heap size from its platform. With it a synchronous session of `AVERAGE($A$1:A<row>)+COUNTA($A$1:A<row>)` over 34,000 rows, whose ranges would hold about 4.6 GB, opens in a Node heap of 1 GiB.

`test/cache-budget.test.ts` checks the order of dropping, the second-read rule, and sessions under a budget of 32 KiB, alone and two sharing it, against full computations through edits, sheet growth and a structure operation, with closures and eval. `test/async-pool.test.ts` opens and edits the running books above in a pool of 384 MB. A random session test with edits, structure operations and spills compared 1,600 hashes with full computations, under the default budget and under one that drops everything after each formula. `test/session-fuzz.test.ts` keeps such sessions in the suite, seeded, with closures and eval and under a budget of 8 KiB: a data sheet of one column read through wider rectangles and whole columns, written past its edge, then read again from the caches. It fails on the code before the cached-shape fix (`9e6dd70`) and when `bounds()` ignores a cached array's shape.

One thread of 384 MB, 20 edits spread over column A, hash equal to the synchronous computation's. Before, none of these books opened there at 10,000 rows: the thread ran out of memory, and with `COUNTIF($A$1:A<row>,">0")` the whole process went down.

| Running formula | Rows | Open | Edit p50 | Edit max |
|---|---:|---:|---:|---:|
| `AVERAGE($A$1:A<row>)` | 10,000 | 0.56 s | 0.34 s | 0.45 s |
| `COUNTA($A$1:A<row>)` | 10,000 | 0.24 s | 0.09 s | 0.13 s |
| `COUNTIF($A$1:A<row>,">0")` | 10,000 | 0.58 s | 0.36 s | 0.47 s |
| `COUNTIF($A$1:A<row>,1)` | 10,000 | 1.7 s | 1.3 s | 1.6 s |
| `AVERAGE(...)+COUNTA(...)` | 10,000 | 0.84 s | 0.52 s | 0.80 s |
| `AVERAGE($A$1:A<row>)` | 34,000 | 6.1 s | 5.2 s | 6.1 s |
| `COUNTA($A$1:A<row>)` | 34,000 | 1.8 s | 1.4 s | 1.7 s |
| `COUNTIF($A$1:A<row>,">0")` | 34,000 | 6.4 s | 5.1 s | 6.2 s |
| `COUNTIF($A$1:A<row>,1)` | 34,000 | 18.7 s | 15.2 s | 18.7 s |
| `AVERAGE(...)+COUNTA(...)` | 34,000 | 9.3 s | 7.4 s | 9.0 s |

Twelve sessions of `AVERAGE(...)+COUNTA(...)` over 5,000 rows, six on each of two threads of 384 MB, opened in 1.5 s and took an edit each in 1.1 s with the synchronous hashes; before, a thread ran out of memory. These books stay quadratic in time: each formula reads its whole prefix. Keeping every range from its first read, under the same budget, was slower on the running mean over 10,000 rows in the pool (open 0.64 to 0.67 s against 0.56 to 0.58 s, edit p50 0.44 against 0.34 to 0.39 s), and a synchronous session would keep about 400 MB of prefixes where it now keeps none.

Ordinary books are not affected: the synthetic book of 34,000 rows keeps 3.6 MiB by the estimate. Two rounds of alternating pairs with the engine without the budget at a load average of 3 to 5, cold full computation, median of five: 185/196 and 184/185 ms with closures, 194/197 and 183/183 ms with eval at 3,400 rows; 1,599/1,581 and 1,589/1,610 ms with closures, 1,591/1,659 and 1,600/1,601 ms with eval at 34,000 rows; hashes `101578a9c6b600b1` and `e64d5f4d0f0ac632` unchanged. A session of the synthetic book of 34,000 rows under the limits of a 384 MB pool opened in 644/674 and 654/662 ms, retained 52/52 MiB and edited at p95 in 3.47/3.62 and 3.56/3.57 ms; 5,000 running sums opened in 39/40 and 40/41 ms with 8.1/8.2 MiB retained.

Lookup indexes over cached arrays (`src/functions/lookup/exact.ts`) are not counted; they go with their array. The ranges one formula reads are counted only between formulas, so a single formula that reads more than the budget holds it while it computes.
