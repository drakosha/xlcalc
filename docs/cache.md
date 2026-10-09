# Range and lookup caches

The runtime keeps large range values and SUMIFS/COUNTIFS indexes in `RangeCache` (`src/range-cache.ts`). A cached entry keeps its numeric bounds and is indexed by sheet and column. A write checks the entries in its columns and that sheet's wide ranges, rather than parsing and checking every cached key. The bounds are the referenced area, not just its occupied cells: writing below the occupied part of `A:A` invalidates it too. Deletion and clearing remove the entries from both the values and the column index.

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

An edit invalidates states through the same range index as the value cache. Spills use that invalidation path too. The states do not allocate a flat array of all preceding values for every prefix. Multi-argument calls, arrays, unions and calls with a finite `limits.cells` use the existing argument path. The finite-limit fallback preserves allocation checks and their cache-hit behavior; this optimization therefore does not apply to a pool that supplies a finite cells limit.

`test/range-aggregate.test.ts` checks read counts in descending evaluation order, addition order, raw underflow/overflow, error precedence, edits, structure operations, horizontal ranges, spills and limits. A separate frozen-build comparison checked 312 results across execution modes and locales, including values, issues and both hashes. That comparison also exposed a pre-existing shape mismatch in cached ranges after sheet growth, addressed separately.

Five alternating pairs against the statistics checkpoint `71fd5e6`, same machine, clock and options as above, with tests stopped. These changes are additional to the earlier cache invalidation and lookup indexes:

| Book and operation | Before | After |
|---|---:|---:|
| 5,000 prefix SUM formulas; open session, closures | 161 ms | 35 ms |
| Same book: change A2500 to 2 | 83 ms | 9.3 ms |
| Same book: open session, eval | 170 ms | 41 ms |
| Same book: edit, eval | 82 ms | 9.3 ms |
| synthBook, seed 1, 3,400 rows; full compute | 227 ms | 227 ms |
| synthBook, seed 1, 34,000 rows; full compute | 1,522 ms | 1,543 ms |

The synthetic business book did not show a speed improvement from shared aggregate prefixes. Its edit-A p95 was 1.13/1.01 ms at 3,400 rows and 3.50/3.52 ms at 34,000 rows (before/after).

Separate fresh-process measurements of the 5,000-prefix book retained about 102 MiB before and 7.2 MiB after, measured as heap growth after GC while the session was open, median of five runs. The new implementation opened 34,000 prefix formulas in 194 ms and edited the middle input in 60 ms, retaining about 53 MiB. The old implementation was not run at that size because its separate prefix arrays contain roughly 578 million values.
