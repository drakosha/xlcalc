# xlcalc against other JavaScript spreadsheet libraries

This optional benchmark compares xlcalc with HyperFormula and IronCalc as calculation engines, and with ExcelJS and SheetJS as xlsx readers and writers, on every operation both sides have. It adds no dependency to the xlcalc package: the libraries compared are pinned exactly in this directory's `package.json`.

## Run

From the repository root, after `npm run build`:

```sh
npm ci --prefix bench/hyperformula --ignore-scripts
npm ci --prefix bench/compare --ignore-scripts
node bench/hyperformula/run.mjs --out /tmp/xlcalc-hf
node bench/compare/run.mjs --out /tmp/xlcalc-compare
node --max-old-space-size=8192 bench/compare/read-bench.mjs "$PWD/../base" "$PWD" "$PWD/bench/compare"
node --expose-gc --max-old-space-size=8192 bench/compare/workload-pair.mjs "$PWD/../base" "$PWD"
node --max-old-space-size=8192 bench/compare/synth-pair.mjs "$PWD/../base" "$PWD"
```

`XLCALC_BENCH_ROOT=<directory>` measures another compiled xlcalc with the same harness, as in `bench/hyperformula`. `run.mjs --part engines,structure,xlsx,memory` selects parts. The two pair scripts take two compiled xlcalc directories (absolute paths), a base and a branch.

## Protocol

- Libraries: HyperFormula and IronCalc compute; ExcelJS and SheetJS read and write xlsx and compute nothing, so they are compared on xlsx alone. Sets of functions without a workbook and a dependency graph (formula.js and the like) have nothing to recompute and are left out. Exact versions are in `package.json` and in each run's `raw.json`.
- Opening and editing are `bench/hyperformula`'s nine workloads and protocol: that benchmark's own run for xlcalc and HyperFormula, and its child process for IronCalc through an adapter (`iron.mjs`). IronCalc gets each cell by `setUserInput` with its calculation paused, then one `evaluate`; a text is typed after an apostrophe and a formula with `=`, as HyperFormula gets them. Edits of several cells pause the calculation the same way; an edit of one cell recalculates as IronCalc does on its own.
- Structure operations insert one row in the middle of a book's first sheet, delete one row there, and insert 100 rows, each undone (`structure-child.mjs`): HyperFormula's `addRows`, `removeRows` and `undo`, IronCalc's `insertRows`, `deleteRows` and `undo`, xlcalc's `editStructure` and `undo`. One warmup round and five measured ones (three on the 34,000-row book, which only xlcalc opens in time).
- xlsx is the synthetic book of 3,400 and of 34,000 rows written by xlcalc's `writeXlsx`, read from its bytes by each library (`readXlsx`, ExcelJS `workbook.xlsx.load`, SheetJS `XLSX.read`) and written back from what it read (`writeXlsx({ book, result: cached })`, `workbook.xlsx.writeBuffer()`, `XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })`). The first read is cold and gives the retained heap; one more read warms up and five are timed; then the same for writing.
- Reading cells runs in one process (`read-bench.mjs`): HyperFormula's `getRangeValues`, `getSheetValues` and `getCellValue` against `readCells` with the field `value` and `getCell`, nine rounds taking turns, median. An answer of `readCells` is an object with the sheet and the cell's address per cell, so it cannot reach a bare array of values; that is the contract, not a target.
- Every populated cell is compared after opening and after every edit and operation, outside the timer, against an independent full calculation by xlcalc of the book as xlcalc's session leaves it; xlsx values against xlcalc's `readXlsx`. Values are brought to one form first: a date as its 1900 serial, a formula's cached result, rich text joined, a missing cell and an empty text alike; numbers agree within `1e-9 * max(1, |a|, |b|)`. IronCalc reads values only as shown text, so its populated cells get the format `0.00000000000000E+00` (15 digits) once, outside the timer, before they are first read. A case whose values differ keeps its time for reference only, with the count of differing cells and an example; books are not fitted to a library and differing cells are not left out.
- Each case and library is its own Node process, one after another: the engines' cases with a 1024 MiB heap and 30 s for a phase as in `bench/hyperformula`, the others with 4096 MiB and 60 s. Medians; p95 where a sample has more than ten values. `raw.json` keeps every sample, the load average before and after each case and `uptime` at the start and end.
- Retained memory is the JavaScript heap after a full collection with the opened book held, above what was there before. IronCalc's book lives in WebAssembly memory, which grows and is never given back: its growth while opening one book in a fresh process (`iron-memory.mjs`) is an upper bound, added to its heap.
- Two builds of xlcalc (a base and a branch) are compared in one process taking turns (`read-bench.mjs`, `workload-pair.mjs`, `synth-pair.mjs`), so that the load of a busy host falls on both alike. Between libraries, separate processes are compared at a load average below 10.

## Results on xlcalc 0.47.0

Node 26.7.0, Apple M5 Pro, 48 GiB RAM, 17.10.2026. xlcalc 0.47.0 (`68d530c`), HyperFormula 3.5.0, IronCalc 0.9.0 (`@ironcalc/wasm`), ExcelJS 4.4.0, SheetJS 0.18.5. Milliseconds, medians. A bold factor marks a case where xlcalc's better execution mode is more than 1.5 times slower than the best other library; a time with differing values is for reference and is not ranked.

The opening, editing and memory figures of xlcalc and HyperFormula are `bench/hyperformula`'s own run (`results/hyperformula/`, one-minute load average 4.9 to 10.8); IronCalc's are `results-ironcalc/` (load 1.8 to 6.1). Structure operations and xlsx are `results/raw.json` (load 1.6 to 3.7), with the 34,000-row structure cases and two memory cases run again in `results-rerun/` after a harness fix. The IronCalc rows of `results/raw.json` are 0.8.4 and are superseded: 0.9.0 is up to 25 times faster on the fanout book. Reading is `results/read-0.47.0.txt` (load about 5).

| Opening, ms | xl closures | xl eval | HF | HF tuned | IronCalc |
|---|---:|---:|---:|---:|---:|
| business-3400 | 128 | 157 | 3,973 |  | 2,126 |
| prefix-sum | 33.5 | 38.1 | 47.4 |  | 180 |
| prefix-statistics | 55.7 | 59.8 | 80.8 |  | 1,113 |
| lookup | 61.3 | 61.3 | 2,107 | 1,266 | 13,121 |
| chain **1.7×** | 43.2 | 52.1 | 31.5 |  | 24.8 |
| fanout **1.7×** | 185 | 228 | 372 |  | 110 |
| bulk | 17.5 | 18.6 | 28.2 |  | 12.6 |
| sparse **1.6×** | 7.37 | 7.45 | 60.0 | 17.0 | 4.51 |
| business-34000 | 729 | 758 | over 30 s |  | over 30 s |

HyperFormula's tuned engine is its column index for the lookup book and automatic sparse mapping for the sparse book. Neither HyperFormula nor IronCalc opens the 34,000-row business book within 30 s.

| Edit, ms | xl closures | xl eval | HF | HF tuned | IronCalc |
|---|---:|---:|---:|---:|---:|
| business-3400, criterion | 0.220 | 0.232 | 1.73 |  | 2,025 |
| business-3400, data | 18.6 | 17.9 | 3,007 |  | 2,017 |
| prefix-sum, middle-input | 4.04 | 3.85 | 4.40 |  | 163 |
| prefix-statistics, middle-input | 8.86 | 9.76 | 10.6 |  | 1,079 |
| lookup, lookup-key | 0.206 | 0.222 | 1.09 | 0.165 | 12,746 |
| lookup, lookup-data | 1.21 | 1.22 | 880 | 2.33 | 12,901 |
| chain, root **3.0×** | 5.35 | 5.45 | 5.20 |  | 1.79 |
| fanout, root **2.4×** | 40.4 | 42.5 | 47.1 |  | 16.7 |
| bulk, 2000-inputs **2.2×** | 5.52 | 5.87 | 3.37 |  | 2.49 |
| sparse, one-input **2.0×** | 0.060 | 0.078 | 0.031 | 0.031 | 0.353 |
| business-34000, criterion | 0.469 | 0.470 | over 30 s |  | over 30 s |
| business-34000, data | 103 | 103 | over 30 s |  | over 30 s |

| Retained heap after opening, MiB | xl closures | xl eval | HF | HF tuned | IronCalc, heap + wasm |
|---|---:|---:|---:|---:|---:|
| business-3400 | 17.6 | 14.5 | 17.7 |  | 22.0 (wasm 20.9) |
| prefix-sum | 7.38 | 7.50 | 9.30 |  | 5.24 (wasm 4.88) |
| prefix-statistics **1.9×** | 13.3 | 11.1 | 12.7 |  | 5.71 (wasm 5.19) |
| lookup | 9.37 | 9.07 | 20.5 | 39.0 | 26.5 (wasm 26.3) |
| chain **2.6×** | 13.0 | 12.6 | 4.77 |  | 6.52 (wasm 6.31) |
| fanout | 62.9 | 61.0 | 44.8 |  | 41.6 (wasm 41.4) |
| bulk | 3.92 | 3.10 | 2.23 |  | 2.29 (wasm 2.00) |
| sparse **1.7×** | 1.52 | 1.46 | 236 | 0.863 | 3.77 (wasm 3.56) |
| business-34000 | 52.1 | 49.0 |  |  | 181 (wasm 180) |

IronCalc's book lives in WebAssembly memory, which only grows: its figure is an upper bound.

| Structure operation, ms | xl closures | xl eval | HF | IronCalc |
|---|---:|---:|---:|---:|
| business-3400, insert 1 row | 123 | 129 | 2,734 | 2,123 (differ: 974) |
| business-3400, undo insert 1 row | 100 | 102 | 2,757 | 2,158 (differ: 974) |
| business-3400, delete 1 row | 125 | 135 | 2,748 | 2,126 (differ: 974) |
| business-3400, undo delete 1 row | 101 | 102 | 3,881 | 2,136 (differ: 974) |
| business-3400, insert 100 rows | 126 | 124 | 2,787 | 2,133 (differ: 974) |
| business-3400, undo insert 100 rows | 95.2 | 99.7 | 2,735 | 2,117 (differ: 974) |
| prefix-sum, insert 1 row **4.6×** | 28.8 | 34.3 | 6.28 | 186 |
| prefix-sum, undo insert 1 row **2.9×** | 23.0 | 28.7 | 7.84 | 181 |
| prefix-sum, delete 1 row | 28.6 | 33.4 | 158 | 178 |
| prefix-sum, undo delete 1 row | 22.3 | 26.6 | 172 | 179 |
| prefix-sum, insert 100 rows | 27.2 | 35.8 | 164 | 187 |
| prefix-sum, undo insert 100 rows | 22.7 | 31.3 | 216 | 179 |
| business-34000, insert 1 row | 834 | 878 |  |  |
| business-34000, undo insert 1 row | 820 | 802 |  |  |
| business-34000, delete 1 row | 798 | 827 |  |  |
| business-34000, undo delete 1 row | 769 | 781 |  |  |
| business-34000, insert 100 rows | 826 | 854 |  |  |
| business-34000, undo insert 100 rows | 760 | 781 |  |  |

IronCalc's values differ after every structure operation on the business book: texts typed after an apostrophe (`'25%`, `'=1+1`) are read again as numbers once their rows move, in 974 cells.

| xlsx | xlcalc | ExcelJS | SheetJS |
|---|---:|---:|---:|
| 3400 rows, read, ms | 55.1 | 62.2 (differ: 151) | 51.8 |
| 3400 rows, write back, ms **2.4×** | 76.0 | 106 (differ: 151) | 32.2 |
| 3400 rows, retained heap after reading, MiB | 8.48 | 23.8 (differ: 151) | 9.78 |
| 3400 rows, written file, KiB | 400 | 409 (differ: 151) | 2,710 |
| 34000 rows, read, ms | 463 | 561 | 525 |
| 34000 rows, write back, ms **2.1×** | 727 | 1,019 | 349 |
| 34000 rows, retained heap after reading, MiB | 58.2 | 197 | 76.2 |
| 34000 rows, written file, KiB | 3,416 | 3,459 | 19,527 |

ExcelJS drops a cached result of 0 of an array formula, 151 cells of the 3,400-row book (`ПоПроектам!B2`: 0 expected, nothing read). SheetJS writes without compression by default: its files are 5.7 to 6.8 times larger, which makes its writing faster.

| Reading, ms | xlcalc | HF |
|---|---:|---:|
| window A1:P100 **3.5×** | 0.420 | 0.120 |
| column B, its rows **5.3×** | 29.0 | 5.46 |
| column B whole | 31.0 | 140 |
| the sheet by pages of 10,000 **18.9×** | 340 | 18.0 |
| 10,000 single cells **11.0×** | 19.0 | 1.72 |
| window ПоНаправлениям!A1:P50 **4.1×** | 0.330 | 0.080 |
| the sheets of formulas by pages of 10,000 **4.2×** | 1.91 | 0.450 |
| the data sheet by pages of 10,000 **14.6×** | 27.0 | 1.85 |
| 10,000 single cells of ПоНаправлениям **14.3×** | 10.0 | 0.700 |

Reading is the values of the data sheet alone (506,594 values) and then the synthetic book with its formulas (3,400 rows), `readCells` with the field `value` against HyperFormula's arrays of values.

## Where xlcalc loses, and why

Profiles are `node --cpu-prof` self time by function, the operation repeated alone after warming up.

- **Reading cells.** Reading the whole data sheet by pages: `readQuery` 17 %, `cellResult` 15 %, `readOrder` 15 %, `projected` 8 %, garbage collection 5 %, `parseAddress` 4 %. Each cell went through a generator of its own, built its whole result (value, formula, format, spill fields) for one field asked, formatted its address twice and parsed it back, and was merged from three streams (input cells, spills, looks) by generators even where no spill or look meets the range. `getCell` read its cell as a one-cell rectangle of `getCells`. An answer of `readCells` is an object with the sheet and address per cell; HyperFormula gives bare arrays of values, which the contract does not.
- **Edits against IronCalc** (chain, fanout, a batch of inputs). Editing the root of the fanout book, 50,000 formulas: evaluating them 31 %; keeping each changed cell's old value by a text key and listing the changes `setCells` answers with 37 %; dropping cached ranges at each write 8 %; ordering 8 %; garbage collection 6 %. IronCalc's `setUserInput` answers with nothing, and the rest is its evaluation in compiled Rust. The list of changes is the contract of an edit's answer; HyperFormula, which also answers with its changes, is slower than xlcalc on these books.
- **2,000 inputs in one call** against HyperFormula: evaluating the 2,000 formulas was 16 % of the call; the rest was bookkeeping per edited cell: an edit's fields put over the cell's input by deleting from a copy (which leaves a slow object for every later read), content compared by writing JSON, inputs kept for undo, argument checks, and the sheet's name upper-cased about five times per cell.
- **One input of the sparse book**, 0.06 against 0.03 ms: the fixed cost of one call (checks, the answer, the undo entry). Not pursued: below 0.1 ms.
- **Opening against IronCalc** (chain, fanout, sparse): compiling each formula and building the graph in JavaScript against compiled Rust; not profiled here. Shared formula templates (see `bench/hyperformula/README.md`) are the known candidate.
- **Inserting one row** in the cumulative SUM book against HyperFormula: every structure operation builds the book's state again (tokenizing, parsing and printing formulas about 15 %, range caches 10 %, comparing moved cells for the answer 8 %, garbage collection 8 %), about what opening costs. HyperFormula transforms references lazily, so one row is cheap for it, while deleting a row or inserting 100 costs it 158 to 216 ms against xlcalc's 22 to 29.
- **Writing xlsx** against SheetJS: SheetJS does not compress by default. Against ExcelJS, which does, xlcalc is faster.
- **Memory** of the chain and of cumulative formulas: compiled code and graph per formula; not profiled here.

## Not compared

HyperFormula has no Find, no `summarizeRanges` and no text a cell shows by its number format; IronCalc's WebAssembly package reads and writes no xlsx; ExcelJS and SheetJS compute nothing. Browsers and the transfer of pages between threads of a calculation pool are not measured.
