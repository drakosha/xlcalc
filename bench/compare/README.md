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

`XLCALC_BENCH_ROOT=<directory>` measures another compiled xlcalc with the same harness, as in `bench/hyperformula`. `run.mjs --part engines,structure,xlsx,memory` selects parts and `--only part:arguments,...` the cases whose part and arguments begin so. The two pair scripts take two compiled xlcalc directories (absolute paths), a base and a branch.

## Protocol

- Libraries: HyperFormula and IronCalc compute; ExcelJS and SheetJS read and write xlsx and compute nothing, so they are compared on xlsx alone. Sets of functions without a workbook and a dependency graph (formula.js and the like) have nothing to recompute and are left out. Exact versions are in `package.json` and in each run's `raw.json`.
- Opening and editing are `bench/hyperformula`'s nine workloads and protocol: that benchmark's own run for xlcalc and HyperFormula, and its child process for IronCalc through an adapter (`iron.mjs`). IronCalc gets each cell by `setUserInput` with its calculation paused, then one `evaluate`; a text is typed after an apostrophe and a formula with `=`, as HyperFormula gets them. Edits of several cells pause the calculation the same way; an edit of one cell recalculates as IronCalc does on its own.
- Structure operations insert one row in the middle of a book's first sheet, delete one row there, and insert 100 rows, each undone (`structure-child.mjs`): HyperFormula's `addRows`, `removeRows` and `undo`, IronCalc's `insertRows`, `deleteRows` and `undo`, xlcalc's `editStructure` and `undo`. One warmup round and five measured ones (three on the 34,000-row book, which only xlcalc opens in time).
- xlsx is the synthetic book of 3,400 and of 34,000 rows written by xlcalc's `writeXlsx`, read from its bytes by each library (`readXlsx`, ExcelJS `workbook.xlsx.load`, SheetJS `XLSX.read`) and written back from what it read (`writeXlsx({ book, result: cached })`, `workbook.xlsx.writeBuffer()`, `XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })`). The first read is cold and gives the retained heap; one more read warms up and five are timed; then the same for writing.
- Reading cells runs in one process (`read-bench.mjs`): HyperFormula's `getRangeValues`, `getSheetValues` and `getCellValue` against `readCells` with the field `value` and `getCell`, nine rounds taking turns, median. An answer of `readCells` is an object with the sheet and the cell's address per cell, so it cannot reach a bare array of values; that is the contract, not a target.
- Every populated cell is compared after opening and after every edit and operation, outside the timer, against an independent full calculation by xlcalc of the book as xlcalc's session leaves it; xlsx values against xlcalc's `readXlsx`. Values are brought to one form first: a date as its 1900 serial, a formula's cached result, rich text joined, a missing cell and an empty text alike; numbers agree within `1e-9 * max(1, |a|, |b|)`. IronCalc reads values only as shown text, so its populated cells get the format `0.00000000000000E+00` (15 digits) once, outside the timer, before they are first read. A case whose values differ keeps its time for reference only, with the count of differing cells and an example; books are not fitted to a library and differing cells are not left out.
- Each case and library is its own Node process, one after another: the engines' cases with a 1024 MiB heap and 30 s for a phase as in `bench/hyperformula`, the others with 4096 MiB and 60 s. Medians; p95 where a sample has more than ten values. `raw.json` keeps every sample, the load average before and after each case and `uptime` at the start and end.
- Retained memory is the JavaScript heap after a full collection with the opened book held, above what was there before. IronCalc's book lives in WebAssembly memory, which grows and is never given back: its growth while opening one book in a fresh process (`iron-memory.mjs`) is an upper bound, added to its heap.
- Two builds of xlcalc (a base and a branch) are compared in one process taking turns (`read-bench.mjs`, `workload-pair.mjs`, `synth-pair.mjs`), so that the load of a busy host falls on both alike. Between libraries, separate processes are compared, each case counted only at a one-minute load average of 6 or less before and after it; a case above that is run again alone (`run.mjs --only engines:chain ironcalc`). A run is held awake by `caffeinate -i`, and `pmset -g log` is checked for sleep after it.
- Each mode of execution of xlcalc has its own figures, closures and eval: a host that forbids `new Function` has closures alone.
- Retained heap leaves out typed arrays: xlcalc keeps 8 bytes per formula in them and, after the first edit, 4 per vertex of the graph (about 0.1 MiB for the chain and 0.6 MiB for fanout when last counted).

## Results on xlcalc 0.67.0

Node 26.7.0, Apple M5 Pro, 48 GiB RAM, 29.10.2026. xlcalc 0.67.0 (`75a6bcd2`), HyperFormula 3.5.0, IronCalc 0.9.0 (`@ironcalc/wasm`), ExcelJS 4.4.0, SheetJS 0.18.5. Milliseconds, medians. A bold factor marks a case where xlcalc's better execution mode is more than 1.5 times slower (or larger) than the best other library that gives the same values; a time with differing values is for reference and is not ranked.

The opening, editing and memory figures of xlcalc and HyperFormula are `bench/hyperformula`'s own run (`results-067/hyperformula/`, one-minute load average 3.7 to 4.4). IronCalc's, structure operations, xlsx and IronCalc's memory are `results-067/compare/`: `raw.json` is the first run, its cases above a load of 6 were run again alone (`results-067/rerun/`), and `merged.json` takes each case from a run at 6 or less (load 3.5 to 5.8). IronCalc 0.9.1, the latest, was measured the same way (`results-067/compare-ironcalc-0.9.1/`, load 3.3 to 5.4): within 8 % of 0.9.0 in every case but its chain's root edit (2.22 against 1.96 ms), with the same memory, so the stand keeps 0.9.0. The README's agreement with Excel, on the coverage run's formulas, is `results-067/accuracy.txt`.

| Opening, ms | xl closures | xl eval | HF | HF tuned | IronCalc |
|---|---:|---:|---:|---:|---:|
| business-3400 | 105 | 104 | 3,765 |  | 2,208 |
| prefix-sum | 27.9 | 30.8 | 45.6 |  | 181 |
| prefix-statistics | 51.1 | 53.0 | 79.8 |  | 1,118 |
| lookup | 44.5 | 47.3 | 1,970 | 1,272 | 13,478 |
| chain | 34.8 | 39.0 | 31.0 |  | 25.0 |
| fanout | 131 | 182 | 366 |  | 113 |
| bulk | 14.2 | 15.2 | 27.9 |  | 13.1 |
| sparse | 5.34 | 5.72 | 65.2 | 16.8 | 4.45 |
| business-34000 | 575 | 545 | over 30 s |  | over 30 s |

HyperFormula's tuned engine is its column index for the lookup book and automatic sparse mapping for the sparse book. Neither HyperFormula nor IronCalc opens the 34,000-row business book within 30 s.

| Edit, ms | xl closures | xl eval | HF | HF tuned | IronCalc |
|---|---:|---:|---:|---:|---:|
| business-3400, criterion | 0.175 | 0.163 | 0.964 |  | 2,046 |
| business-3400, data | 7.14 | 6.46 | 2,800 |  | 2,065 |
| prefix-sum, middle-input | 3.62 | 3.30 | 5.21 |  | 164 |
| prefix-statistics, middle-input | 8.91 | 7.62 | 9.78 |  | 1,086 |
| lookup, lookup-key | 0.146 | 0.152 | 1.01 | 0.100 | 13,599 |
| lookup, lookup-data | 1.01 | 0.997 | 880 | 2.54 | 13,521 |
| chain, root | 2.41 | 2.30 | 4.51 |  | 1.96 |
| fanout, root | 16.0 | 22.7 | 42.3 |  | 18.1 |
| bulk, 2000-inputs | 3.24 | 2.85 | 3.31 |  | 2.62 |
| sparse, one-input **2.8×** | 0.076 | 0.058 | 0.020 | 0.022 | 0.361 |
| business-34000, criterion | 0.183 | 0.179 | over 30 s |  | over 30 s |
| business-34000, data | 9.32 | 8.24 | over 30 s |  | over 30 s |

| Retained heap after opening, MiB | xl closures | xl eval | HF | HF tuned | IronCalc, heap + wasm |
|---|---:|---:|---:|---:|---:|
| business-3400 | 16.6 | 14.2 | 17.7 |  | 22.0 (wasm 20.9) |
| prefix-sum | 6.27 | 6.37 | 9.28 |  | 5.24 (wasm 4.88) |
| prefix-statistics **1.6×** | 11.4 | 9.02 | 12.7 |  | 5.71 (wasm 5.19) |
| lookup | 7.75 | 7.42 | 20.4 | 38.9 | 26.5 (wasm 26.2) |
| chain **1.6×** | 7.41 | 8.77 | 4.76 |  | 6.52 (wasm 6.31) |
| fanout | 38.5 | 45.3 | 44.8 |  | 41.6 (wasm 41.4) |
| bulk | 2.90 | 2.43 | 2.23 |  | 2.29 (wasm 2.00) |
| sparse **1.6×** | 1.36 | 1.48 | 236 | 0.848 | 3.77 (wasm 3.56) |
| business-34000 | 51.2 | 48.7 |  |  | 181 (wasm 180) |

IronCalc's workbook lives in WebAssembly memory, which only grows: its figure is an upper bound.

| Structure operation, ms | xl closures | xl eval | HF | IronCalc |
|---|---:|---:|---:|---:|
| business-3400, insert 1 row | 91.0 | 96.3 | 2,805 | 2,156 (differ: 974) |
| business-3400, undo insert 1 row | 72.5 | 76.3 | 2,994 | 2,170 (differ: 974) |
| business-3400, delete 1 row | 98.9 | 106 | 2,897 | 2,153 (differ: 974) |
| business-3400, undo delete 1 row | 76.5 | 79.3 | 4,018 | 2,162 (differ: 974) |
| business-3400, insert 100 rows | 103 | 95.8 | 2,829 | 2,155 (differ: 974) |
| business-3400, undo insert 100 rows | 79.0 | 77.3 | 2,779 | 2,153 (differ: 974) |
| prefix-sum, insert 1 row **4.0×** | 25.2 | 30.6 | 6.36 | 185 |
| prefix-sum, undo insert 1 row **2.6×** | 20.2 | 25.7 | 7.64 | 179 |
| prefix-sum, delete 1 row | 29.1 | 29.7 | 166 | 179 |
| prefix-sum, undo delete 1 row | 23.4 | 24.5 | 178 | 179 |
| prefix-sum, insert 100 rows | 26.9 | 29.8 | 163 | 189 |
| prefix-sum, undo insert 100 rows | 21.5 | 23.4 | 216 | 182 |
| business-34000, insert 1 row | 654 | 633 |  |  |
| business-34000, undo insert 1 row | 582 | 657 |  |  |
| business-34000, delete 1 row | 674 | 654 |  |  |
| business-34000, undo delete 1 row | 658 | 578 |  |  |
| business-34000, insert 100 rows | 708 | 661 |  |  |
| business-34000, undo insert 100 rows | 587 | 601 |  |  |

IronCalc's values differ after every structure operation on the business book: texts typed after an apostrophe (`'25%`, `'=1+1`) are read again as numbers once their rows move, in 974 cells.

| xlsx | xlcalc | ExcelJS | SheetJS | SheetJS compressed |
|---|---:|---:|---:|---:|
| 3400 rows, read, ms | 50.3 | 63.8 (differ: 151) | 54.9 | 54.3 |
| 3400 rows, write back, ms | 56.4 | 109 (differ: 151) | 33.8 | 50.0 |
| 3400 rows, retained heap after reading, MiB | 8.46 | 23.9 (differ: 151) | 9.84 | 9.83 |
| 3400 rows, written file, KiB | 400 | 409 (differ: 151) | 2,710 | 984 |
| 34000 rows, read, ms | 408 | 579 | 527 | 534 |
| 34000 rows, write back, ms | 518 | 1,068 | 368 | 473 |
| 34000 rows, retained heap after reading, MiB | 58.2 | 197 | 76.2 | 76.2 |
| 34000 rows, written file, KiB | 3,416 | 3,459 | 19,527 | 8,486 |

ExcelJS drops a cached result of 0 of an array formula, 151 cells of the 3,400-row book (`ПоПроектам!B2`: 0 expected, nothing read). SheetJS writes without compression by default: its files are 5.7 to 6.8 times larger, which makes its writing faster; with `compression: true` they are still 2.5 times larger than xlcalc's, and xlcalc's write back takes 1.1 times as long as SheetJS's with compression. Every library's written file, read back by xlcalc, gives the first file's values.

Reading is the median of three runs of `read-bench.mjs` (`results-067/read-2.txt` to `read-4.txt`, load 3.8 to 5.6; `read-1.txt` ran above 6 and is not counted), display false. `readCells` gives an object with the sheet and the address for each cell and `getCell` a cell's whole view: those are their contracts; `getValues` is the call that gives bare values, as HyperFormula's arrays do, and none of its values differs from HyperFormula's. The factor is the faster xlcalc call against HyperFormula.

| Reading, ms | `readCells` | `getValues` | HF |
|---|---:|---:|---:|
| window A1:P100 | 0.38 | 0.04 | 0.05 |
| window A1:P100, value, text and formula | 0.71 | | |
| column B, its rows | 22 | 0.71 | 6.00 |
| column B whole (`getValues` with `maxCells`) | 19 | 47 | 122 |
| the sheet (506,594 values), `readCells` by pages of 10,000 | 169 | 13 | 14 |
| 10,000 single cells, `getCell` **3.1×** | 6.37 | | 2.07 |
| workbook with formulas, 3,400 rows: window of a sheet of formulas **2.8×** | 0.28 | 0.14 | 0.05 |
| the same: its six sheets of formulas by pages **1.5×** | 0.97 | | 0.64 |
| the same: its data sheet | 11 | 0.92 | 1.78 |
| the same: 10,000 single cells of a sheet of formulas, `getCell` **7.4×** | 3.42 | | 0.46 |
| workbook with formulas, 34,000 rows: its data sheet | 179 | 11 | |
| the same: 10,000 single cells of a sheet of formulas, `getCell` | 3.47 | | |

## Where xlcalc loses, and why

The cases in bold above. None was profiled again on 0.67.0: the reasons are from the last profiles of these cases, on 0.47.0 to 0.52.0 (`node --cpu-prof` self time by function, the operation repeated alone after warming up).

- **One input of the sparse book** (2.8× against HyperFormula's 0.020 ms): the fixed cost of one `setCells` call (checks, the answer, the undo entry), which the stand measures cold, 20 calls after two. Warm, in one process, the call took 4.8 µs against HyperFormula's 2.9 when last profiled.
- **Retained memory of the chain** (1.6× against HyperFormula's 4.76 MiB), **of prefix-statistics** (1.6× against IronCalc's 5.71) **and of the sparse book** (1.6× against HyperFormula's sparse mapping, 0.848): each formula's compiled code (closures, or eval's parameters), its reference and its listener entry; prefix-statistics also keeps the prefix sums of its range caches, which make its edits faster than HyperFormula's.
- **Inserting a row in the cumulative SUM book** (4.0×) and its undo (2.6×): every structure operation builds the workbook's state again (about 70 % of the operation, recomputing every SUM 25 %). HyperFormula transforms references lazily, so one row is cheap for it, while deleting a row or inserting 100 costs it 163 to 216 ms against xlcalc's 21 to 29. Shifting references and the graph in place is a larger rewrite of the structure operations.
- **Reading single cells** (`getCell`, 3.1× and 7.4×) and **a sheet of formulas by pages** (`readCells`, 1.5×): an answer per cell with its address and view, where HyperFormula returns a bare value. The window of a sheet of formulas by `getValues` (2.8×, 0.14 against 0.05 ms) took 0.02 ms in 0.66.0 in the same runs, and in a loop of its own the call takes 0.023 ms in both versions, 0.03 to 0.04 ms on a first call; why the stand gives 0.67.0 more is not found.

Near the line, not in bold: opening the chain in eval (1.56× against IronCalc, closures 1.39×), opening fanout in eval (1.61×, closures 1.16×), and a lookup key's edit against HyperFormula's column index (1.46×). Opening against IronCalc is compiling each formula and building the graph in JavaScript against compiled Rust; formulas alike but for their references and numbers are already parsed once per shape (`src/prepare.ts`). IronCalc's native binding for Node, not this stand's WebAssembly build, was measured once on 0.47.0 by another harness: two to three times faster than the WebAssembly build on the chain, fanout and batch edits; it is not ranked here.

## Not compared

HyperFormula has no Find, no `summarizeRanges` and no text a cell shows by its number format; IronCalc's WebAssembly package reads and writes no xlsx; ExcelJS and SheetJS compute nothing. Browsers and the transfer of pages between threads of a calculation pool are not measured.
