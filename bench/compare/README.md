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

## Speed-ups on the branch `perf/speed`

Both against 0.47.0, two builds taking turns in one process, medians. Every answer stays as it was: `test/read-answers.test.ts` and `test/edit-answers.test.ts` compare each answer's JSON (fields, order, pages, cursors, counts, errors, the book after each edit) with fixtures written by 0.47.0, in every kind of session and both executions; the synthetic hashes are unchanged.

### Reading cells

A field that needs no shown text is read from the cell's value and input alone, without a generator per cell or the cell's whole result; with `text`, and in `getCell`, one cell is read by its own path instead of as a one-cell rectangle of `getCells`. Where no spill and no look meets a range, its cells come from the index of the sheet's cells alone, and the address built to check a cell is the one the answer gives. A sheet named as the book spells it is found without upper-casing the name. `to15` keeps a number of hundredths below 1e13 as it is, without making text: it is its own 15-digit rounding (`test/to15.test.ts` checks it against rounding through text).

Median of three runs of `read-bench.mjs` (each nine rounds), display false, field `value` unless named; HyperFormula's figure is its own call in the same runs.

| Case | 0.47.0, ms | branch, ms | Faster | HyperFormula, ms |
|---|---:|---:|---:|---:|
| window A1:P100 | 0.41 | 0.21 | 1.95× | 0.12 |
| window A1:P100, value, text and formula | 1.62 | 0.86 | 1.88× | |
| column B, its rows | 28 | 14 | 2.0× | 5.48 |
| column B whole | 31 | 14 | 2.2× | 107 |
| the sheet (506,594 values) by pages of 10,000 | 358 | 157 | 2.28× | 19 |
| 10,000 single cells (`getCell`) | 19 | 8.74 | 2.17× | 1.76 |
| book with formulas, 3,400 rows: window of a sheet of formulas | 0.35 | 0.15 | 2.3× | 0.08 |
| the same: its six sheets of formulas | 2.15 | 0.88 | 2.4× | 0.65 |
| the same: its data sheet by pages | 30 | 10 | 3.0× | 1.77 |
| the same: 10,000 single cells of a sheet of formulas | 14 | 3.69 | 3.8× | 0.57 |
| book with formulas, 34,000 rows: its data sheet by pages | 355 | 159 | 2.2× | |
| the same: 10,000 single cells of a sheet of formulas | 10 | 3.83 | 2.6× | |

The single cells take more after the whole sheet's pages in the same process than alone: run alone, 31 rounds, `getCell` goes from 16 to 6.8 ms.

Display true gives the same figures within noise, every case faster than 0.47.0; `results/read-*.txt` holds the three runs. After the change the whole sheet's read spends its time in the walk of the index (`readOrder` 24 %), the page loop (`readQuery` 14 %), reading values with their 15-digit rounding (`valueAt`, `cellValue`, `to15`, about 13 %) and the answer's objects (`projected` 3 %); `getCell` in building the cell's view (`resultAt` 21 %, `cellAt` 11 %, `viewOf` 7 %), its value (12 %), and checking and formatting the address (10 %). The answer's object per cell with its address is the contract.

### 2,000 inputs in one call

An edit's fields go over a cell's input by copying the fields that stay instead of deleting from a copy, so inputs stay fast objects; content is compared without writing JSON; one engine helper finds a sheet by its name for the session's checks, edits, undo inputs and answers.

`workload-pair.mjs` over the nine workloads, 30 edits per build:

| 2,000 inputs, ms | 0.47.0 | branch | Faster |
|---|---:|---:|---:|
| closures, p50 | 8.03 | 5.89 | 1.36× |
| eval, p50 | 6.53 | 4.63 | 1.41× |

A separate run of this workload alone gave 5.34 to 3.62 ms (1.47×) with closures and 5.02 to 3.34 ms (1.50×) with eval. HyperFormula took 3.37 ms and IronCalc 2.49 ms in the comparison above. Evaluating the formulas is now 23 % of the call, against 16 % before; putting fields over inputs fell from 15.5 % to 6 %.

## Other cases on the branch

`workload-pair.mjs` (seven openings and 30 edits per build) and `synth-pair.mjs` (the synthetic book: five cold calculations, 200 edits, and structure operations at row 10 with undo and redo, seven each), medians; 0.47.0 against the branch, ms. None is slower by more than 10 %.

| Case | closures | eval |
|---|---|---|
| business-3400 open; criterion; data | 122.9 / 125.2; 0.059 / 0.059; 18.0 / 17.8 | 130.1 / 141.6; 0.051 / 0.061; 18.1 / 18.1 |
| prefix-sum open; edit | 36.3 / 37.9; 4.95 / 4.80 | 41.2 / 40.5; 4.54 / 4.71 |
| prefix-statistics open; edit | 62.4 / 62.2; 9.98 / 9.41 | 65.2 / 66.7; 9.15 / 8.60 |
| lookup open; key; data | 72.6 / 71.5; 0.018 / 0.017; 0.921 / 0.803 | 71.1 / 72.5; 0.018 / 0.017; 0.841 / 0.879 |
| chain open; root | 54.1 / 51.6; 6.95 / 7.03 | 73.5 / 75.6; 8.72 / 8.07 |
| fanout open; root | 192.0 / 191.7; 40.7 / 38.7 | 237.8 / 237.9; 48.0 / 44.0 |
| sparse open; one input | 7.33 / 7.23; 0.014 / 0.012 | 7.88 / 7.49; 0.014 / 0.012 |
| business-34000 open; criterion; data | 738.8 / 731.3; 0.263 / 0.262; 103.9 / 103.5 | 773.7 / 760.3; 0.265 / 0.263; 103.5 / 103.5 |
| synthetic 3,400: cold; edit p50 | 181.0 / 187.2; 0.638 / 0.619 | 186.4 / 189.5; 0.636 / 0.626 |
| synthetic 34,000: cold; edit p50 | 1724.6 / 1618.6; 2.795 / 2.758 | 1751.2 / 1680.3; 2.770 / 2.761 |

The criterion edit of business-3400 with eval (0.051 against 0.061 ms in 30 edits) was measured again alone over 300 pairs: 0.0545 against 0.0520 ms, the branch 0.002 ms faster by the median of paired differences; with closures 0.0547 against 0.0518.

Structure operations at row 10 of the synthetic book, ms, 0.47.0 / branch:

| Operation | 3,400, display false | 3,400, display true | 34,000, display false | 34,000, display true |
|---|---|---|---|---|
| insert a row | 125.9 / 124.3 | 167.3 / 158.0 | 883.8 / 882.1 | 1200.0 / 1160.8 |
| its undo | 100.6 / 95.0 | 133.9 / 130.5 | 809.2 / 818.5 | 1130.3 / 1076.7 |
| its redo | 100.3 / 97.8 | 133.7 / 131.2 | 822.7 / 824.7 | 1102.9 / 1104.0 |
| delete a row | 135.0 / 136.4 | 161.5 / 155.1 | 935.2 / 912.5 | 1155.4 / 1162.3 |
| its undo | 102.9 / 96.3 | 136.3 / 128.5 | 851.1 / 832.9 | 1110.4 / 1068.5 |
| its redo | 100.4 / 101.0 | 140.8 / 127.4 | 842.7 / 833.0 | 1109.2 / 1091.2 |
| undo again | 95.5 / 103.5 | 134.1 / 131.6 | 837.3 / 858.5 | 1088.3 / 1076.7 |

A first run of the synthetic pair was discarded: the host went to sleep with its lid closed during it (`pmset -g log`, 12:28), which the second run, held awake by `caffeinate -i`, did not.

## Parity on the branch `perf/parity`

The aim: in every case above, xlcalc no more than 1.5 times slower than the best other library that gives the same values, and no more than 1.5 times its retained memory. The base is 0.48.0 (`f45c8d0`), measured by both stands at the start (one-minute load average 2 to 10); the branch is `018ab99` (0.49.1 merged in), measured by both stands at the end (load 2 to 5), with xlcalc now beside IronCalc in the comparison's own run, so that a ratio is taken within one run: the best other library and xlcalc's branch figure come from the same run of the same stand. Base and branch are also compared taking turns in one process (`workload-pair.mjs`, `synth-pair.mjs`, `read-bench.mjs`). Each mode of execution has its own figures, closures / eval: a host that forbids `new Function` has closures alone. The raw results are in `results-parity/`.

Every answer stays as it was: `test/edit-answers.test.ts` (with `test/recompute-answers.ts`: a fanout of 120, a chain of 100 with cumulative sums, statistics and a lookup over them, an unchanged INT, a spill sized by the edit, a batch of 70 inputs, two independent roots, a formula replaced in the middle) and `test/read-answers.test.ts` compare each answer's JSON with fixtures written by the base, in every kind of session and both executions; the synthetic hashes are unchanged; a written xlsx has the same bytes.

### What changed

- **An edit's recalculation.** A changed cell is kept by a number of its place instead of an `s:r:c` text, from the tracker of old values to the list an edit answers with; the formulas to recompute are ordered by filtering the book's order when they are a large part of it, after a walk of the graph that marks vertices in a typed array instead of filling a set; a formula's step in a round is a typed array instead of a map; a write does not look up range caches that are empty; the readers of an edited cell go straight into the set to recompute; an edit looks for old-kind array areas only in a book that has them; two numbers compare without looking for errors. The list of changes `setCells` answers with is kept: no option to leave it out was needed.
- **Opening and memory.** A formula is found by its cell's number; references and successors are arrays of their own length (an array grown by pushes keeps room for 16 more); a range's cache key is made when a cache first asks for it; in closures, an operator over two cells or a cell and a constant is one closure instead of three, and a call of up to three arguments makes no array of them each time.
- **Reading.** `getValues(sheet, range)` gives a rectangle's values as rows, as HyperFormula's `getRangeValues` and `getSheetValues` give arrays of values (`docs/live-books.md`, "Values of a rectangle").
- **Writing xlsx.** A sheet's cells' XML is kept by row and column instead of in a map of every cell's place.

### Results

Medians, ms unless named; closures / eval. "Best other" is the fastest (or the smallest) other library that gives the same values, from the branch's run; a library whose values differ is not ranked (IronCalc after structure operations on the business book, ExcelJS on the 3,400-row file, as above). The base column is the base's run, a figure of its own run, not a ratio.

| Opening, ms | base, closures / eval | branch | best other | branch against it |
|---|---:|---:|---:|---:|
| business-3400 | 121 / 123 | 109 / 115 | IronCalc 2,112 | 0.05× / 0.05× |
| prefix-sum | 32.4 / 42.3 | 32 / 33.2 | HF 46.2 | 0.69× / 0.72× |
| prefix-statistics | 61.6 / 60.8 | 57.2 / 54.1 | HF 77.8 | 0.74× / 0.70× |
| lookup | 57.9 / 57.8 | 48.5 / 57.7 | HF index 1,215 | 0.04× / 0.05× |
| chain | 40.1 / 47.8 | 39.1 / 41.5 | IronCalc 24.1 | 1.63× / 1.73× |
| fanout | 167 / 223 | 140 / 191 | IronCalc 109 | 1.28× / 1.75× |
| bulk | 16.9 / 18 | 16 / 17.1 | IronCalc 12.3 | 1.30× / 1.39× |
| sparse | 6.48 / 6.95 | 5.95 / 6.5 | IronCalc 4.28 | 1.39× / 1.52× |
| business-34000 | 718 / 822 | 600 / 676 |  |  |

| Edit, ms | base, closures / eval | branch | best other | branch against it |
|---|---:|---:|---:|---:|
| business-3400, criterion | 0.213 / 0.205 | 0.171 / 0.170 | HF 1.48 | 0.12× / 0.11× |
| business-3400, data | 18.2 / 18.9 | 17.8 / 16.8 | IronCalc 2,006 | 0.01× / 0.01× |
| prefix-sum, middle-input | 4.36 / 4.02 | 3.59 / 3.86 | HF 4.59 | 0.78× / 0.84× |
| prefix-statistics, middle-input | 9.15 / 9.15 | 8.1 / 7.81 | HF 9.56 | 0.85× / 0.82× |
| lookup, lookup-data | 1.16 / 1.14 | 0.970 / 0.928 | HF index 3.05 | 0.32× / 0.30× |
| lookup, lookup-key | 0.185 / 0.205 | 0.152 / 0.163 | HF index 0.133 | 1.14× / 1.23× |
| chain, root | 4.9 / 4.76 | 2.15 / 2.12 | IronCalc 1.9 | 1.13× / 1.11× |
| fanout, root | 36.1 / 43.2 | 15.7 / 21.1 | IronCalc 15.4 | 1.02× / 1.37× |
| bulk, 2000-inputs | 3.47 / 3.39 | 2.63 / 2.72 | IronCalc 2.39 | 1.10× / 1.14× |
| sparse, one-input | 0.051 / 0.051 | 0.052 / 0.050 | HF sparse 0.021 | 2.43× / 2.36× |
| business-34000, criterion | 0.434 / 0.436 | 0.405 / 0.415 |  |  |
| business-34000, data | 103 / 104 | 101 / 102 |  |  |

| Retained heap after opening, MiB | base, closures / eval | branch | best other | branch against it |
|---|---:|---:|---:|---:|
| business-3400 | 17.6 / 14.5 | 15.7 / 13.2 | HF 17.8 | 0.88× / 0.74× |
| prefix-sum | 7.37 / 7.5 | 6.64 / 6.72 | IronCalc 5.24 | 1.27× / 1.28× |
| prefix-statistics | 13.3 / 11.1 | 12.7 / 10.3 | IronCalc 5.71 | 2.22× / 1.80× |
| lookup | 9.41 / 9.12 | 9.18 / 8.86 | HF 20.4 | 0.45× / 0.43× |
| chain | 13 / 12.6 | 7.18 / 8.53 | HF 4.77 | 1.51× / 1.79× |
| fanout | 62.9 / 61 | 37 / 43.8 | IronCalc 41.6 | 0.89× / 1.05× |
| bulk | 3.93 / 3.09 | 3.1 / 2.6 | HF 2.23 | 1.39× / 1.17× |
| sparse | 1.52 / 1.48 | 1.04 / 1.18 | HF sparse 0.839 | 1.24× / 1.41× |
| business-34000 | 52.1 / 48.9 | 50.2 / 47.7 | IronCalc 181 | 0.28× / 0.26× |

| Structure operation, ms | base, closures / eval | branch | best other | branch against it |
|---|---:|---:|---:|---:|
| business-3400, insert 1 row | 124 / 131 | 117 / 122 | HF 2,684 | 0.04× / 0.05× |
| business-3400, undo insert 1 row | 97 / 102 | 86.4 / 93.2 | HF 2,723 | 0.03× / 0.03× |
| business-3400, delete 1 row | 124 / 127 | 113 / 127 | HF 2,699 | 0.04× / 0.05× |
| business-3400, undo delete 1 row | 97 / 100 | 89.4 / 95.1 | HF 3,831 | 0.02× / 0.02× |
| business-3400, insert 100 rows | 125 / 128 | 116 / 119 | HF 2,744 | 0.04× / 0.04× |
| business-3400, undo insert 100 rows | 94.6 / 102 | 91.8 / 94 | HF 2,692 | 0.03× / 0.03× |
| prefix-sum, insert 1 row | 28.9 / 33.3 | 25.9 / 31 | HF 6.55 | 3.95× / 4.72× |
| prefix-sum, undo insert 1 row | 23.1 / 28.4 | 19.5 / 25.8 | HF 8.16 | 2.40× / 3.17× |
| prefix-sum, delete 1 row | 31.2 / 35 | 28.1 / 31.3 | HF 160 | 0.18× / 0.20× |
| prefix-sum, undo delete 1 row | 23.6 / 28.4 | 21.1 / 24.3 | HF 173 | 0.12× / 0.14× |
| prefix-sum, insert 100 rows | 28.1 / 36.4 | 25.9 / 31.4 | HF 163 | 0.16× / 0.19× |
| prefix-sum, undo insert 100 rows | 22.7 / 29.3 | 20 / 24 | IronCalc 177 | 0.11× / 0.14× |
| business-34000, insert 1 row | 839 / 838 | 733 / 757 |  |  |
| business-34000, undo insert 1 row | 800 / 775 | 701 / 694 |  |  |
| business-34000, delete 1 row | 825 / 836 | 748 / 750 |  |  |
| business-34000, undo delete 1 row | 783 / 774 | 697 / 701 |  |  |
| business-34000, insert 100 rows | 842 / 842 | 745 / 730 |  |  |
| business-34000, undo insert 100 rows | 777 / 779 | 686 / 722 |  |  |

| xlsx | base | branch | best other | branch against it |
|---|---:|---:|---:|---:|
| 3400 rows, read, ms | 54.1 | 53.9 | SheetJS 53.3 | 1.01× |
| 3400 rows, write back, ms | 81.6 | 75.8 | SheetJS 33.3, not compressed | 2.27× |
| 3400 rows, write back against compressing, ms | 81.6 | 75.8 | SheetJS compressed 50.5 | 1.50× |
| 3400 rows, retained heap after reading, MiB | 8.48 | 8.48 | SheetJS 9.78 | 0.87× |
| 3400 rows, written file | | xlcalc 400 KiB, ExcelJS 408 KiB, SheetJS 2,710 KiB, SheetJS compressed 983 KiB | | |
| 34000 rows, read, ms | 540 | 440 | SheetJS 540 | 0.82× |
| 34000 rows, write back, ms | 751 | 680 | SheetJS 363, not compressed | 1.88× |
| 34000 rows, write back against compressing, ms | 751 | 680 | SheetJS compressed 463 | 1.47× |
| 34000 rows, retained heap after reading, MiB | 58.1 | 58.1 | SheetJS 76.2 | 0.76× |
| 34000 rows, written file | | xlcalc 3,416 KiB, ExcelJS 3,458 KiB, SheetJS 19,527 KiB, SheetJS compressed 8,486 KiB | | |

Beside the JavaScript heap the branch keeps, in typed arrays, 8 bytes per formula and, after the first edit, 4 per vertex of the graph: about 0.1 MiB for the chain and 0.6 MiB for fanout, which the heap figures above do not count.

SheetJS writes without compression by default, so its files are 5.7 to 6.8 times larger; with `compression: true` they are still 2.5 times larger than xlcalc's. Every library's written file, read back by xlcalc, gives the first file's values.

Reading is one process taking turns (`read-bench.mjs`, display false), the branch's `getValues` beside `readCells`; every value of `getValues` was compared with HyperFormula's arrays, none differs.

| Reading, ms | base | branch | HyperFormula | branch against it |
|---|---:|---:|---:|---:|
| window A1:P100, `readCells` | 0.21 | 0.21 | 0.10 | 2.1× |
| window A1:P100, `getValues` | | 0.04 | 0.05 | 0.8× |
| column B, its rows, `readCells` by pages | 14 | 13 | 5.12 | 2.5× |
| column B, its rows, `getValues` | | 0.45 | 4.55 | 0.10× |
| column B whole, `readCells` by pages | 16 | 14 | 107 | 0.13× |
| column B whole, `getValues` with `maxCells` | | 57 | 108 | 0.53× |
| the sheet (506,594 values), `readCells` by pages of 10,000 | 157 | 156 | 14 | 11× |
| the sheet, `getValues` | | 9.96 | 13 | 0.77× |
| 10,000 single cells, `getCell` | 5.99 | 4.98 | 2.06 | 2.4× |
| book with formulas, 3,400 rows: window of a sheet of formulas, `readCells` | 0.22 | 0.22 | 0.06 | 3.7× |
| the same, `getValues` | | 0.10 | 0.05 | 2.0× |
| the same: its data sheet, `readCells` by pages | 13 | 10 | 1.72 | 5.8× |
| the same: its data sheet, `getValues` | | 0.77 | 1.63 | 0.47× |
| the same: 10,000 single cells of a sheet of formulas, `getCell` | 3.39 | 2.72 | 0.45 | 6.0× |

`readCells` gives an object with the sheet and the address for each cell and `getCell` a cell's whole view: those are their contracts, and `getValues` is the call that gives bare values, as HyperFormula's arrays do.

### Base against branch in one process

`workload-pair.mjs` (seven openings and 30 edits per build) and `synth-pair.mjs` (the synthetic book: cold calculations, 200 edits, structure operations at row 10 with undo and redo), medians: no median is slower than the base by more than 1 %. Opening is up to 1.22 times faster, the chain's and fanout's root edits 1.9 to 2.3 times, the batch of 2,000 inputs 1.13 to 1.23 times, one input of the sparse book 1.6 times, the synthetic book's cold calculation 1.14 to 1.18 times and its structure operations up to 1.12 times (`results-parity/workload-pair.txt`, `synth-pair.txt`).

### Not at parity, and why

- **Opening the chain** (1.63× / 1.73× against IronCalc) **and eval's opening of fanout and sparse** (1.75×, 1.52×): preparing formulas, that is tokenizing, parsing, checking and printing each formula, is 37 % of opening the chain, compiling 13 %, the graph 11 %. HyperFormula's answer is to parse a formula's relative shape once for a column of copied formulas; for xlcalc that is a cache of parsed shapes in `engine.ts: prepare`, about 300 to 400 lines, not done in this round.
- **One input of the sparse book** (2.4× against HyperFormula's 0.021 ms): the fixed cost of one `setCells` call, which the stand measures cold (20 calls after two). Warm, in one process, the call takes 4.8 µs against HyperFormula's 2.9.
- **Retained memory of the chain** (1.51× / 1.79× against HyperFormula's 4.77 MiB) **and of prefix-statistics** (2.2× / 1.8× against IronCalc's 5.7): each formula's compiled code (closures, or eval's parameters), its reference and its listener entry; prefix-statistics also keeps the prefix sums of its range caches, which make its edits faster than HyperFormula's.
- **Inserting a row in the cumulative SUM book** (4.0× / 4.7×) and its undo (2.4× / 3.2×): every structure operation builds the book's state again (70 % of the operation, recomputing every SUM 25 %). Deleting a row or inserting 100 is 5 to 9 times faster than HyperFormula. Shifting references and the graph in place instead is a larger rewrite of the structure operations, planned separately.
- **Writing xlsx** (1.88× / 2.27× against SheetJS without compression, 1.47× / 1.50× against SheetJS compressing): writing the XML is about 70 % of the time, deflate 25 %. The platform's `CompressionStream` takes no level, so a faster level would need another codec.

IronCalc's native binding for Node, not this stand's WebAssembly build, was measured once on 0.47.0 by another harness: two to three times faster than the WebAssembly build on the chain, fanout and batch edits. It is not ranked here.

## Not compared

HyperFormula has no Find, no `summarizeRanges` and no text a cell shows by its number format; IronCalc's WebAssembly package reads and writes no xlsx; ExcelJS and SheetJS compute nothing. Browsers and the transfer of pages between threads of a calculation pool are not measured.
