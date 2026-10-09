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
