# Development

The commands, the tests and how the rules were checked against Excel.

```sh
npm ci
npm run typecheck
npm test                # build, then all tests against src
npm run test:dist       # build, then all tests against the built dist
npm run test:oracle     # recalculate the LibreOffice answers with soffice and check them
npm run check:package   # build, then publint and attw on the packed tarball
npm run pack:release    # build, then the tarball in release/
FORMULA_BENCH=1 npx vitest run test/budgets.test.ts   # budgets on the full synthetic book
```

The tests cover the whitelist against Excel's answers on the coverage measurement's cases (`test/excel.test.ts`) and on the probes (`test/excel-probes.test.ts`), the argument layer rule by rule, the semantics cases against Excel's documentation and the stored LibreOffice answers (`test/oracle/libreoffice.json`, refreshed by `npm run test:oracle`, which needs LibreOffice's `soffice` on `PATH`; the usual tests need no office suite); determinism (shuffled input, a child process, three time zones, cents over 3,400 amounts); references, errors, volatile functions, parsing and the canonical form; sessions (200 random edits against full computations) and `explain`; structure operations against Excel's answers in `test/excel/struct/`, on the synthetic book and on seeded books; budgets; SHA-256 against test vectors and `node:crypto`; a browser bundle run in an empty JavaScript context and with code generation disabled; the two execution modes against each other on every case book and over 200 session edits, the closures mode with code generation disabled; computation in portions and in threads against the synchronous engine, threads lost and sessions restored, the pool on Web Workers in headless Chromium, Firefox and WebKit (through `playwright-core`; each skipped where its engine is not found, `CHROME_PATH` names a Chromium); the published package (exports, tarball contents, a consumer importing it by name). The code blocks of this README run as a test too.

Excel's answers that the tests compare with are stored in `test/excel/`, so the tests need no Excel: the coverage measurement's cases and Excel columns in `coverage/`, the case sets in `cases/` (`cases/ru-RU/` for the sets computed in ru-RU), the probes in `probes.json`, structure operations in `struct/`, spilling in `spill/`, Excel's ru-RU formula text in `locale/`, the files Excel saved for the xlsx subpath in `xlsx/`. They are recomputed only with Excel for Mac, by a step of `npx tsx bench/coverage/run.ts`: `excel en` and `excel ru` for the measurement, `probes` for the probes, `probes <set>` and `probes <set> ru` for a case set, `struct <set>`, `spill <set>`, `xlsx <step>`, `entry <set>` for typed text, `cond <set>` for conditional formats; `npx tsx bench/locale/snapshot.ts ru` takes Excel's ru-RU formula text. A run keeps the stored answers of unchanged cases and computes the new and changed ones. `bench/coverage/README.md` has the layout, the steps and the rules for driving Excel. Excel for Windows's answers to the same books are in `test/excel/win/`, computed by `npx tsx bench/windows/run.ts` on a Windows machine over SSH (`bench/windows/README.md`).

Functions live in `src/functions/<category>.ts` (`math`, `statistical`, `logical`, `information`, `text`, `date`, `financial`, `lookup`, `engineering`, `database`; old names such as `NORMDIST` sit next to their modern versions), each file with its own registry, a large category assembled from `src/functions/<category>/`; `src/functions.ts` assembles them. A function declares the kinds of its arguments (`args: 'n i?'`, see `src/args.ts`), and the shared argument layer coerces them, so an implementation gets numbers, text, lists or grids and never repeats the coercion rules. Ported code keeps its attribution in the file header. A function joins its registry when `test/excel.test.ts` passes for it; the coverage measurement (`bench/coverage`, `docs/coverage-formulajs.md`) lists the cases and the classes of the remaining formula.js functions.

## How the rules were checked

The measurements behind each guide's "Checked against Excel": the harness steps, the sets, what each predicted set gave and what its misses changed.

### Arrays inside a formula

From [Arrays inside a formula](arrays-incell.md).

The `arrays` case set (`bench/coverage/sets/arrays.ts`) is computed in Excel as dynamic array formulas. The harness writes each formula the way Excel 365 saves one: `<c cm="1"><f t="array" ref="A1">…</f>` plus `xl/metadata.xml` with the `XLDAPR` metadata and `dynamicArrayProperties fDynamic="1"`. Without that metadata Excel reads a formula the old way, with implicit intersection: `SUM(IF(D!B1:B8>0,D!B1:B8))` is 114.75 as an old formula and 118.25 as a dynamic one. The `_xlfn._xlws.` prefixes only name newer functions and do not make a formula dynamic.

A case whose answer is an array is checked through a wrapper that returns one value (SUM, INDEX, ROWS, COLUMNS, COUNTA, CONCAT, TEXTJOIN), or directly, where both Excel and the engine give `#SPILL!` because every neighbour of the case is occupied.

All case sets of the library are computed as dynamic array formulas, so the whole acceptance follows Excel 365.

The set `known-held` (37 cases) was predicted before Excel computed it, to check rules taken from a few answers of the random check: `SORTBY` with an array of orders, `LOOKUP` past the end of a shorter result, `IFS` with a condition of text under `ROWS`, `TYPE` and `VSTACK`, and `SWITCH`, `IF` and `CHOOSE` beside it. Excel for Mac and Excel for Windows gave the same 37 answers, and 26 came out as predicted. The 11 misses are now the engine's rules: `SORTBY` answers by position rather than element by element (6), and `LOOKUP` reads a result of one cell as a row, reads on with a computed lookup array and reads a blank past the sheet's end (5). All 11 `IFS` cases came out as predicted.

The set `lookup-held` (19 cases) was predicted the same way for what `known-held` left: `SORTBY` with two orders of one shape, valid orders over a by-array of another size, an order of one value beside one of two, and `LOOKUP` past the sheet's last column and row and past a reference a function answers. Both Excels gave the same answers and 10 came out as predicted. The 9 misses, and 14 cases measured after them (the set `lookup-rules`), are now the engine's rules: `SORTBY` answers `#VALUE!` after the first position that sorts when it has more than one by-array, a by-array of another size gives its error at every position, the first position that sorts need not be the first, and past a shorter order the position is `#N/A`; `LOOKUP` reads on past any reference, one a spill, a function, parentheses or a range operator answers too. The set `lookup-own-held` (9 cases, predicted) puts the `LOOKUP` beside the cells it may read on to: beside a spilled column and in a row past what it may read both Excels and the engine agree (5); where the `LOOKUP` lies among the cells it may read but does not, the engine, which counts them all as its references, answers `#CYCLE!` and Excel computes it (4, a known difference).

The set `sortby-one-held` (19 cases, predicted) puts a by-array of one value into `SORTBY` beside others: in a column, a row, a table of two columns and one cell, with an order left out and with orders by position. Both Excels gave the same answers, all 19 as predicted: one value fits an array of one column or one row and nothing wider, the first by-array decides between rows and columns and a later one must fit it (`SORTBY(A1:A6,B1:B6,1,5,1)` and `SORTBY(A1:A6,5,1,B1:B6,1)` are `#VALUE!`), and an error given as a by-array is the answer.

The shared collector and array helpers were additionally compared with frozen build `69b4d1f` in both execution modes and locales: 2,600 complete results matched in values, issues and both hashes, including edits, whole-column fills, row/column selection, old-style arrays and conditional aggregates. Across 672 limit checks, the result or cell-limit error also matched at limits of 4, 20 and 100 values. These used local builds and the saved Excel fixtures, without a live Excel run.

On Node 26.7, 2026-10-09, seven alternating measured pairs after two warmup pairs (ten calls per sample) gave these medians for arrays of 34,000 elements. MAXA and AVERAGEA used mixed numbers, text and booleans; FILTER retained every third row; CHOOSECOLS selected the same column three times.

| Call | Before | Shared implementation |
|---|---:|---:|
| SUM | 0.332 ms | 0.347 ms |
| MAXA | 0.297 ms | 0.253 ms |
| AVERAGEA | 0.251 ms | 0.272 ms |
| FILTER | 0.295 ms | 0.248 ms |
| SORT | 3.564 ms | 3.429 ms |
| CHOOSECOLS | 1.064 ms | 0.940 ms |

The SUM and AVERAGEA microbenchmarks retain a small overhead. The collector stays local to the argument handler: moving it to a module-level function slowed mixed-array coercion substantially in the trial. V8's inlining trace showed the original per-element closure being inlined; the module-level collector remained a separate function.

Five alternating pairs of `synthBook({ seed: 1, rows: 34000 })`, with closures and fixed clock/seed, gave 1,561/1,547 ms for full compute and 3.71/3.53 ms for edit-A p95 (before/after). This workload showed no overall regression from the shared helpers.

### Asynchronous and threaded computation

From [Asynchronous and threaded computation](async.md).

- `test/async-inline.test.ts` and `test/async-pool.test.ts` compute the coverage measurement (in en-US and ru-RU), the semantics cases, the synthetic book, the spill book, a seeded book, iterative calculation and the 1904 date system in the same thread and in threads, with closures and eval, and compare values, issues and hashes with the synchronous engine; a session over 40 edits gives the changes and the hash of a synchronous session after each, in both execution modes and in both calculators; two sessions on one thread with their calls mixed without waiting each give their own changes and hash; a structure operation gives the synchronous session's changes and `mapCell` on eleven operations, filtered rows among them.
- In the same thread a timer runs while a book computes; deadlines, aborts (waiting and running) and limits give `CalcError`; a stopped `setCells`, `setLayout` or `editStructure` leaves the book and the hash as they were; 300 formulas that each sort 10,000 values stop at a deadline of 100 ms within 200 ms, the event loop never waiting 50 ms (measured alone: 100 to 102 ms, waits of 10 to 21 ms with slices of 10 ms).
- `test/async-inline-long.test.ts`: in the same thread, on the synthetic book of 8,500 rows, `getResult()`, `getHash()` and `getCells()` of a large rectangle give the synchronous session's values while the event loop runs, never waiting a quarter of the call's time; `insertRows`, `deleteColumns`, `renameSheet`, `moveRange` and `renameTable` over 60,000 formulas stop at a deadline of 20 ms within a quarter of the operation's time, and the session keeps its hash and then gives the synchronous session's changes. The result's steps and the checks of structure operations cost the synchronous path nothing measurable: a cold computation of the synthetic book, five runs each taken in turns with the engine without them, medians, two rounds: 246 and 253 ms against 249 and 245 ms with closures, 296 and 297 against 298 and 302 ms with eval (3,400 rows); 1,833 and 1,656 against 1,864 and 1,629 ms with closures, 1,821 and 1,820 against 1,821 and 1,802 ms with eval (34,000 rows); the values hash is the same.
- `test/async-recovery.test.ts` terminates threads by a hard deadline, by `maxHeapMb` (twenty spilled arrays of a million values in a 64 MB heap) and by a crash, and checks that the session's hash after it comes back equals that of a synchronous session given the same answered calls, from a snapshot and a journal cut at it too; a restore that fails closes that session only, and ten of them on a 16 MB book leave a 96 MB thread alive; a session of ten calls of 150 ms comes back within a deadline of 800 ms after it left a full thread and after a crash; sessions take turns on a full thread; `maxQueue`, `maxSessions`, the order of a thread's tasks, closing.
- `test/session-window.test.ts`: in a synchronous session, in the same thread and in a thread, with closures and eval, each sheet's `getUsedRange` equals the reach of the cells, spill ranges, merged areas and tables of the session's own result and book, on the books of the tests and while edits grow and shrink it (a spill taller, blocked, shorter, a cell far away patched and cleared, merged areas, inserted rows, a deleted table); `getLayout()` equals the book without its cells after layout edits, a structure operation and new styles, and a thread's is a copy; `cells` of several rectangles equals the rectangles read one by one, each cell once and in address order. `test/session-patch.test.ts`: patches on an empty cell, a value, a formula, an old-kind array's area and a spill keep the values hash and the arrays, list the cells a new format shows otherwise, leave a live session's clock as it was, change nothing when refused, and give the synchronous session's changes and book in the same thread and in a thread, after a thread is lost too, from the journal and from a snapshot taken after them. The seeded sessions of `test/session-fuzz.test.ts` patch cells after each step and compare each sheet's extent with the full computation's cells.
- `test/session-undo-fuzz.test.ts`: seeded sessions make changes of every kind (edits, a table's header typed over, patches, names, styles, layout edits, every structure operation, sheets hidden and shown among them), undo and redo some of them and make new changes, and after each move the book equals the book its place in the history had, the hash equals a full computation's of that book, the history counts the steps each way, a host's screen moved by `mapCell` and given the changes holds the session's values, and a refused change leaves the book and the history as they were; in a synchronous session, in the same thread and in a thread lost now and then with snapshots every 3 calls, with closures and eval. A history that does not put back the style table, the names, the cells a structure operation emptied, a line's entries or a sheet's fields, or a snapshot without the history, makes it fail. `test/session-undo.test.ts` takes each kind of step back and forth in both modes, and `test/session-undo-async.test.ts` compares the same thread and a thread with a synchronous session, stops an undo by its deadline and its signal, loses a thread with undos in the journal and in a snapshot, and has two sessions take turns on a full thread; in a pool of 384 MB and 6 sessions a thread a session keeps 8 MB of steps by the estimate, as a synchronous session with `maxMb: 8`, and one with its own `maxMb` keeps its own.
- `test/async-pool-edges.test.ts`: a book or an edit with a function in `meta` is refused with `InputError` and the thread goes on; a session closed while the pool closes or while its thread is lost is closed; `maxQueue: 0` takes what an idle thread takes at once; `maxHeapMb` with `spawn` is refused.
- `test/limits.test.ts`: the limits in `compute`, in a session and in `createInlineCalc`, with `SEQUENCE`, `MUNIT`, `MMULT`, lifted operators, `TOCOL` of whole columns and spills: `spill` refuses a spill of more stored cells (an old-kind array formula too) but not an array that does not spill, and `A:A*2`, `TOCOL(A:A)` and `IF(A:A>250,A:A,0)` over 500 rows compute under a `spill` of 1,000. `test/async-pool.test.ts`: a pool of 64 MB refuses an array of 1,048,577 values and a spill of 80,000, a pool of 384 MB computes `SUM(SEQUENCE(500000))` and `TOCOL(A:A)`, opens a running total over 10,000 rows and opens and edits running `AVERAGE`, `COUNTIF` and `AVERAGE` with `COUNTA` over 10,000 rows, and a host that raises `cells` keeps the bound on spills.
- `test/limits-writes.test.ts`: in a synchronous session, in the same thread and in a thread, with closures and eval, each edit `writes` bounds (a copy onto two tiles, of part of an old-kind array, from a filtered sheet and onto one, a move, `fillCellText`, `pasteText`, `replaceInCells`, `setCells`, `replaceCells`, `setCellTexts`, `applyChanges` and the previews) is refused one cell over its count with `limit: 'writes'`, `max` and `size`, the book, its revision, history and hash as before and the session going on, and at its count gives the answer and the hash of a session without limits; a copy of one cell, a value and `=ROW()*2`, onto a whole column is refused within 50 ms; a calculator's `writes` holds under a session's limits of other fields and a session's own takes its place; inserted and deleted lines, a copy of empty cells, `patchCells` and `clearCells` pass under `writes: 1`.
- `test/locale-pool.test.ts`: a pool started before ru-RU is loaded refuses a ru-RU book with the synchronous engine's `InputError`; once ru-RU is loaded its threads compute such a book and give a session's explanation and a refusal in Russian, as a synchronous session does, on both threads, with closures and eval; a pool started later and a thread started again after a crash have it; a locale outside the package computes in a thread as sent, in English there as here; a locale without its messages survives a structured clone, a locale with them does not.
- `test/worker-bundle.test.ts` bundles `xlcalc` for the browser and finds no file of the threads in it; `test/dist.test.ts` runs the pool of the built package by its name from a plain Node process, which exits with the pool open.
- `test/worker-browser.test.ts` loads the built package in headless Chromium, Firefox and WebKit (through `playwright-core`): the pool on Web Workers gives Node's hashes for the synthetic book and a session's edits, a hard deadline terminates a worker and the session comes back, and the same-thread calculator lets timers run. Each engine is skipped where it is not found.

#### Explain cancellation

`test/explain-cancel.test.ts` runs both execution modes in inline sessions and pools: already aborted, queued and running calls, soft deadlines, invalid options and successful reuse after cancellation. An expanded range of 12,000 values makes the traversal yield. `test/session-issue.test.ts` checks node locations, errors, ranges, spills and cycles in all three kinds of session. The traversal and text rendering use the same paced budget checks as the other asynchronous methods.

### Circular references

From [Circular references](cycles.md).

The workbook step of the Windows harness (`bench/windows`, `cycle` cases) opens each book through COM, takes the case's steps (a value or a formula typed, a full calculation, a recalculation, a save) and reads the cells after each; the file of a case stores the values it lists, as Excel stores what it computed (`test/excel/win/workbook/`).

| set | books | what | the engine |
|---|---:|---|---|
| `cycles` | 21 | a cycle typed into a book, the values a file stores (converged ones, others, none), a precedent changed, a formula typed again, the link broken and typed back, a cycle of one cell, through INDIRECT, OFFSET, a name and a spill, cells beside a cycle (above it, on another sheet, typed after opening, a chain closed), a file with `calcPr iterate` | all but the cycles through INDIRECT and OFFSET (not in the whitelist) and the iteration from stored values (`iterate-file`, [Limitations](limitations.md)) |
| `cycles-held` | 6 | predicted by the engine before Excel ran them: an array brought into a cycle's way, a formula typed again as it was, a cell cleared, a dependent typed out of the way, two cycles of which one is broken, a file storing values for part of a cycle | 125 of the 129 cells read as predicted; the four others are the spill an array a cycle comes to keeps in Excel |

`test/workbook-win.test.ts` compares the engine with the stored answers without Excel. `test/session-cycles.test.ts` takes the measured scenarios through each kind of session (the same thread synchronously and in portions, a thread of the pool) in both execution modes, with undo, redo, previews, typed text, a look, a row inserted and a copy, and after each step checks the session's hash against a full computation of its book. The seeded session tests do the same after each of their edits.

### Conditional formats

From [Conditional formats](conditional-formats.md).

`npx tsx bench/coverage/run.ts cond <set>` writes each book of `bench/coverage/cond/<set>.ts` with `writeXlsx`; Excel 16.113 for Mac opens it, applies the case's structure operation, computes the book, reads the display format of the listed cells (`Range.Text`, the fill and font colours, bold, italic: what conditional formats show) and saves the book. `npx tsx bench/windows/run.ts cond <set>` does the same in Excel 16.0.20430 for Windows through COM (`DisplayFormat`), and also makes rules through its object model for the set `made`. `test/cond-excel.test.ts` and `test/cond-win.test.ts` check without Excel that the engine shows what Excel showed, that the reader reads from each saved file the rules the engine holds after the case, and that the engine shows on the saved book what Excel showed.

| set | cases | what |
|---|---|---|
| `looks` | 7 | rules of every kind over numbers, text, booleans, blanks, empty text, errors and dates, priorities and `stopIfTrue`, a rule's number format; cases between readings of Excel's rules |
| `looks-ru` | 1 | text rules over Cyrillic and numbers with a decimal comma, text in the locale's order, dates, a number format, in ru-RU |
| `files` | 6 | every kind and option of rule in a file, Excel 2010's rules, many rules, whole columns and rows, books written after structure operations |
| `struct-held` | 24 | rows and columns inserted and deleted, moves and copies within a sheet and to another, sheets renamed and deleted |
| `held`, `held-2` | 9, 4 | held out after the sets above: the rules they gave on other cases |
| `both` | 3 | a sheet with conditional formats, data validation and grouped rows, written, with a row inserted and one deleted |
| `made` | 1 | 15 rules Excel for Windows made through COM (`FormatConditions.Add`, `AddTop10`, `AddAboveAverage`, `AddUniqueValues`, `AddColorScale`, `AddDatabar`, `AddIconSetCondition`), over two ranges too |
| `edge-held` | 8 | rules at the sheet's last row and column, one cell and two, under lines deleted and inserted far before them and the last line deleted |
| `formats-ru-held` | 1 | built-in number formats of rules given by number and by their codes, in ru-RU |
| `fills-held` | 3 | fills of cells and of rules as another writer may hold them (put into the file by hand): `patternType="none"` with colours, a solid fill of one colour, no pattern type, a gradient, a gradient beside a pattern |
| `fills-held-2` | 1 | held out after `fills-held`: colours that name none (`auto`, `indexed="64"` and `"65"`) in fills of cells and rules, a rule's empty pattern fill |
| `sqref-held` | 6 | a rule's cells as another writer may spell them: in lower case, with `$`, both, several ranges, a whole column, Excel 2010's `<xm:sqref>` (put into the file by hand) |

Both Excels loaded every book but one of `fills-held`, a dxf with a gradient beside a pattern. The engine's answers were committed before each run. Excel for Mac showed the looks as predicted in 594 of 676 cells of `looks`, 191 of 200 of `files`, 1,699 of 1,885 of `struct-held` and 79 of 80 of `looks-ru`, and saved the rules of 12 of the 24 cases of `struct-held` as predicted; the misses gave the rules above (how a colour scale cuts its channels, errors left out of statistics, the cell a formula is written for, IF's reading of an answer, a blank operand, blank rules, duplicates, and how operations grow, keep and order rules), and the held-out sets, made after them, showed 708 of 708 and 336 of 336 cells as predicted and the rules of 11 of 13 cases, whose misses gave the order rules of top rules and where a paste grows a rule; `both`, 132 of 132 cells and its rules and data validation. Now the engine shows what Excel for Mac showed in 4,015 of 4,017 cells, the two pinned below. Excel for Windows, predicted after the Mac's answers, showed `looks`, `struct-held`, `held`, `held-2` and `both` as predicted in every cell, and agreed with Excel for Mac in 3,988 of the 4,017 cells both read (the others `#` counted in pixels, and the order of Latin and Cyrillic); the rules Excel for Windows made itself, read from its file, showed as Excel showed them in 203 of 204 cells. Of the rules each Excel saved, the reader reads what the engine holds after the case in every case, but the bar lengths and the order of one rule pinned below. `edge-held` came out as predicted in both Excels, both refusing lines inserted that push the corner's cells off the sheet; `formats-ru-held` showed 5 of 9 cells as predicted in Excel for Windows and 4 in Excel for Mac, whose short month names differ: the rules given a built-in format by number are shown by its code (fixed), and Excel saved the two rules of each pair, by number and by code, as one. In `sqref-held` both Excels took every spelling as the range it names and saved it in upper case without `$`, where the engine predicted five of the six rules lost (it read `$` but not lower case); it now reads them as Excel does. In `fills-held` both Excels showed every cell alike, and the engine had predicted 8 of the 24 cells: a cell's fill and a dxf's are read by one function since, with Excel's differences between the two (`docs/xlsx.md`), and 8 cells stay apart, each reported: the gradient, which Excel's display format gives as black, a dxf's none, which takes the cell's own fill away, and a dxf's solid fill without its colour, white on a cell without a fill. `fills-held-2`, held out after it, held in 7 of 12 cells: a colour that names none is Excel's automatic white in a rule's fill too, and a rule's empty pattern fill changes nothing (fixed); the other miss is a rule's none taking the fill away.

Not measured: a data bar's length and an icon, which no object model reports (the rules' parameters were checked, the bar and the icon follow Excel's documented rules); borders, underline and strike of a style (written and read back, not read from the display format); rules made through Excel for Mac's AppleScript, whose `make new format condition` stopped Excel; a copy across sheets of a rule with relative references beyond the cases above.

### Data validation

From [Data validation](validation.md).

`npx tsx bench/coverage/run.ts valid <set> [ru]` writes each book of `bench/coverage/valid/<set>.ts` with `writeXlsx`. Excel for Mac opens it, makes what the case makes through its object model (`add data validation` with the range's first cell selected, groups, a collapsed group), applies the case's structure operation, and over a window of each sheet reads, with each cell selected in turn, the cell's rule (type, operator, formulas, settings, texts) and `Validation.Value`, whether the cell's value meets it; it reads the outline too and saves the book. `npx tsx bench/windows/run.ts valid <set> [ru]` does the same in Excel for Windows through COM. `test/validation-excel.test.ts` and `test/validation-win.test.ts` check, without Excel, that Excel read what the engine holds after each case, `validateEntry` with each cell's own input against `Validation.Value`, and that the reader reads the file Excel saved as Excel read it.

| set | cases | what |
|---|---|---|
| `files` | 8 | books written with every type and setting, whole columns, several ranges, another sheet, texts past Excel's limits, outlines; books written after structure operations in the engine |
| `made` | 4 | rules of every type, groups and collapsed groups made by Excel itself |
| `struct-held` | 31 | lines inserted and deleted around and inside rules and groups, moves and copies of rules and onto them, sheets renamed and deleted under lists |
| `struct-held-2` | 15 | held out after `struct-held`: lines inserted right after rules, moves onto rules and to another sheet, lines inserted into collapsed groups |
| `entry-held` | 96 values | values of every kind against every type, bounds on blank cells, lists of cells holding numbers as text |
| `lists` | 30 values | lists through names, another sheet, a spill, a formula, a whole column, literal lists with `;` and spaces |
| `entry-held-2` | 39 values | held out after `entry-held` and `lists` |
| `held-3` | 30 values, 2 moves | held out after the second sets: blanks where blanks are not ignored, literal numbers, moves and copies to another sheet |
| `ru` | 2 | in ru-RU: literal lists with commas and semicolons in a file, and lists, bounds and formulas typed into the object model in Russian |
| `edge-held` | 8 | rules at the sheet's last row and column, one cell and two, under lines deleted and inserted far before them and the last line deleted |
| `sqref-held` | 5 | a rule's cells as another writer may spell them: in lower case, with `$`, both, several ranges, Excel 2010's `<xm:sqref>` (put into the file by hand) |

The engine's answers were committed before Excel ran each set (`bench/coverage/valid/*.predicted.json`). Excel loaded every written book. Of the predictions, `files` held in 2 of 8 books (a literal list matches text in its case; texts the writer cuts), `made` in 3 of 4 (lines inserted into a collapsed group open it), `struct-held` in 26 of 31 (lines inserted right after a rule take it; a move keeps the target's rules; a formula referring to a blank cell lets anything in), `struct-held-2` in 14 of 15 (a rule moved to another sheet moves its formulas as a copy), the checks of values in 87 of 96 (`entry-held`), 26 of 30 (`lists`), 36 of 39 (`entry-held-2`) and 27 of 30 (`held-3`), and `edge-held` in 4 of 8 (a rule of one cell on the last row or column stretches over the line before it when lines above are deleted, as a range ending there does; Excel refuses lines inserted that push its cells off the sheet, as predicted). Each miss gave one of the rules above, fitted on the set that showed it. In `sqref-held` both Excels took every spelling as the range it names, where the engine predicted four of the five rules lost (it read lower case but not `$`); it now reads them as Excel does. Excel for Windows read what Excel for Mac read in every set in en-US.

In ru-RU both Excels read a file's literal lists as en-US does, split at commas. Typed through the object model in Russian, a list `x;y;z` is three items in Excel for Windows and one in Excel for Mac's AppleScript, which also made no rule of a formula with a Russian function name, and Excel for Mac's object model shows a rule's formulas now in Russian, now in English, so the ru-RU set compares Excel for Mac's values and settings and leaves the formulas to the file it saved.

What Excel does with a picked item of a dropdown, and the text its dropdown shows, was not measured: no object model gives them. Lists of cells were checked by which values they take.

An item picked from a dropdown was measured by the step `ui` on the Mac's keyboard (the set `pick-held`, 14 picks in each locale: the cell selected, its list opened with Option+Down, the item reached with Down and taken with Return): a list of cells holding text that reads as a number, a date, a boolean or a percent when typed, a number shown as a percent, plain text, a literal list, into cells in General, in the text format and a percent. The prediction, each item typed as its text with the rule checked, was committed before Excel ran the set and came out on 5 picks in each locale: Excel took no item as typed from a list of cells and showed no alert. `test/ui-excel.test.ts` compares `setCells` of each item's `input` with every pick.

### Defined names

From [Defined names](names.md).

The sets `spill/names` (16 books) and `struct/cube` (20 operations) were computed in Excel. The 24 cells of the set `spill/names-held` (10 books) were predicted before Excel computed them, and all came out as predicted.

The rules of a name were measured in Excel 16.113 for Mac and Excel 16.0.20430 for Windows, en-US and ru-RU, through the table sets ([Tables](tables.md)): names in formulas of the file and typed, names added through Excel's object model (`Names.Add`), which refuses a name Excel does not take, tables renamed and books Excel loads or not. Six sets were predicted before Excel computed them:

| set | cases | as predicted (Mac) | misses |
|---|---:|---:|---|
| `names-held` | 17 | 13 | a name typed in another case shown as defined (2, fixed); tables renamed to names Excel does not take (2, known) |
| `names-ru-held` (ru-RU) | 8 | 6 | the name's case (fixed); a file with the name `ИСТИНА`, which Excel ru-RU does not open (fixed) |
| `names-more-held` | 12 | 7 | `?` and `\` inside a name, added and in a file (4, fixed); a 255-letter name in a typed formula, cut by AppleScript (known, the Mac only) |
| `names-ru-more-held` (ru-RU) | 10 | 5 | `ИСТИНА` and `ЛОЖЬ` refused and TRUE and FALSE taken (TRUE known), `?` and `\` (fixed), files with `ИСТИНА`, `истина`, `ЛОЖЬ` not opened (3, fixed); one more the predictions held no text for: the word `ИСТИНА` shown as `'ИСТИНА'` (fixed) |
| `names-tables-held` | 4 | 3 | a table renamed `?Продажи` (known) |
| `names-ru-tables-held` (ru-RU) | 3 | 0 | tables renamed `ИСТИНА`, `ложь`, `TRUE`, where the two Excels part (known) |
| `paste-names-held` (`add-cell-shapes`) | 1 | 0 | `XFE1` and `A0` taken and read by a formula (fixed); `R1`, `C5`, `XFD1`, `r1` and `R1C1` refused, as predicted |
| `name-case-held`, `name-case-sheets-held` | 11 and 10 | 8 and 6 (Windows: 6 and 8; the rule the first set gave predicted the second) | the dotless `ı` is no `I`, the long `ſ` no `S` in Excel for Windows, which the engine follows (fixed; Excel for Mac takes `ſ` for `S`, known); `ß` is `SS` and `ﬁ` is `FI` in names and tables, as predicted; sheets compare letter by letter in both Excels, `ß` beside `SS` (known) |
| `name-case-ru-held` (ru-RU) | 10 | 8 (Windows: 10) | the long `ſ` taken for `S` by Excel for Mac (known); the rest as in en-US |

In all, 34 of the 54 cases before `name-case-held` came out as predicted on the Mac and 35 on Windows, which agreed with the Mac on every case but the 255-letter name and the tables of `names-ru-tables-held`. The two Excels agree on `add-cell-shapes`. Of the 31 cases of the `name-case` sets, 22 came out as predicted on the Mac and 24 on Windows; the two Excels part on the long `ſ` alone.

The words of a formula, TRUE and FALSE, function names and LET's names in the long s, the sharp s, the ligature `ﬁ` and the dotless and dotted i, were measured in a file and typed (the case set `name-case-held` and the entry sets `name-case-held` and `name-case-values-held`): Excel for Windows reads them letter by letter (`falſe` is a name, `gauß(1)` no `GAUSS`, `LET(ß,1,SS)` `#NAME?`), Excel for Mac reads `ſ` as `s`.

LET's names were measured as text typed into cells (the entry sets `names-held`, `let-grid-held` and `let-cells-held`, [Typed entry](entry.md)), predicted before Excel ran them: of the cells' words the engine refused all, and Excel takes every one but those that read as R1C1's row or column too (`R2`, `C100`, `RC5`), on the Mac and Windows alike; `?` and `\` inside, first and after a cell's shape (`A1?x`) are names, as predicted. The engine takes them so now (fixed).

### Editing a grid

From [Editing a grid](edits.md).

The lines harness (`bench/coverage/lines.ts`, `npx tsx bench/coverage/run.ts lines <set>`, and `bench/windows/run.ts lines <set>` for Excel for Windows) writes a book with `writeXlsx`; Excel takes the case's steps through its object model (a number format, bold, italic or a fill set on a range, a value typed into a cell, ClearContents, ClearFormats or Clear of a range), reads what each listed cell shows (`Range.Text`, the number format, the display format's fill and font) and saves the book. `test/lines-excel.test.ts` takes the same steps in a session (`patchCells`, `setCells`, `clearCells`) and compares every cell; for the clearing cases it also compares the notes, hyperlinks, conditional formats and rules of data validation of the file Excel saved, read back by `readXlsx`, with the session's book.

| set | cases | cells | what |
|---|---|---|---|
| `files`, `held` | 15 | 1,087 | measured before for the writer and the reader: columns and rows formatted across each other, values typed into them |
| `patch` | 7 | 124 | a rectangle over a filled column and across an italic row; a column over cells with their own style and format; a row's date format across a column of percents; a column filled, then bold, then a row italic across it; a row italic over a filled cell; a column filled across a row with a number format alone; a style over a book's own line styles |
| `clear` | 12 | 756 | ClearContents, ClearFormats and Clear of a rectangle of values, formulas, styles, formats, notes and hyperlinks at the top of a rule's cells; of cells of a formatted column; of a whole column crossing a bold row and of a whole row; a value typed into cells cleared of their formats |

The engine's answers to `patch` and `clear` were committed before Excel for Mac ran them. Every cell of both sets showed what was predicted, and so did the notes, conditional formats and rules of data validation of the files; the hyperlinks did not: Excel's ClearContents takes a cell's hyperlinks, which the prediction kept, and the session now takes them. Excel for Windows, run after, showed what Excel for Mac showed in every cell and every part of the files. On the cases of `files` and `held` Excel showed what a session's `patchCells` gives in every cell, the cells where Excel formats a line across a formatted one, which `setLayout` leaves to the line's style, among them.

The seeded test of sessions against full computations (`test/session-fuzz.test.ts`) replaces, clears, formats and types into cells beside its other edits, compares the session's hash with a full computation's after each, and at the end undoes every step and finds the book it opened. `test/session-edit.test.ts`, `test/session-styles.test.ts` and `test/session-text.test.ts` check each call in a synchronous session, in the same thread and in a thread, with closures and eval.

`setCellTexts` adds no rule of its own: each text is read as `setCellText` reads it, whose rules were measured in Excel ([Typed entry](entry.md)). `test/session-texts.test.ts` types a mixed set (numbers, dates and a formula in ru-RU, text that reads as a number, a column of percents, a cell with a rule, two cells that widen one column to different widths, a cell given twice, a table's header) at once and one by one, and finds the same book, values and changes; one undo takes the whole call back.

`pasteText` and `fillCellText` follow what Excel for Mac and Excel for Windows did with text pasted from the clipboard and Excel for Mac with Control+Return, measured by the step `ui` ([Typed entry](#typed-entry)); `test/ui-excel.test.ts` compares them with every case. `test/session-paste.test.ts` checks the rectangle, each kind of field, the refusals, the moved formula, the active cell's format and rule, the widths and one undo in a synchronous session, in the same thread and in a thread, with closures and eval.

### Formulas of the old kind

From [Formulas of the old kind](legacy.md).

- `legacy/params` (3,071 formulas): every whitelisted function with a range and an expression over a range put into each of its arguments, read back as Excel 365 shows them (`formula2`) and computed.
- `legacy/nesting` (3,469) and `legacy/blocks` (346): the same inside SUMPRODUCT, SUM, INDEX, ROWS and MMULT.
- `legacy/calls` (113): calls whose answer depends on their arguments, array constants, names.
- `legacy/ifs-text` (8): `IFS` with a condition of text under `ROWS`, `COLUMNS`, `TYPE`, `SUM` and `VSTACK`, predicted before Excel computed it; 7 came out as predicted, and the miss showed that `TYPE` takes a call with an array answer whole (`TYPE(IFS("x",1))` has no `@` and is 64).
- The coverage measurement computed with formulas of the old kind: every whitelisted case gives Excel's value at its row.
- `spill/legacy` and `spill/legacy-held`: intersection by row and column, on other sheets and through names. The 13 cells of `legacy-held` were predicted before Excel computed them, and all came out as predicted.
- `spill/cse`, `spill/cse-held` and `struct/arrays` for array formulas of the old kind. The 25 cells of `cse-held` were predicted before Excel computed them, all as predicted.
- `xlsx intersect`: 75 formulas of the old kind in a file Excel computed and saved. Read by `readXlsx`, each computes the value Excel stored, but two with `OFFSET`, which is not in the whitelist.

`fromLegacy` gives Excel's text for every formula of these sets, `@D!B1:B3 D!B2:C2` for the intersection `D!B1:B3 D!B2:C2` among them.

### Links to other workbooks

From [Links to other workbooks](links.md).

`npx tsx bench/coverage/run.ts links <set>` writes each case of `bench/coverage/links/<set>.ts` as `writeXlsx` writes its own cells, with the formulas over links and the link parts put in by hand as Excel writes them, and has Excel for Mac open it with `update links do not update links`; `npx tsx bench/windows/run.ts links <set>` has Excel for Windows open it with `UpdateLinks` 0. Neither showed a window about the links. Each reads the rectangles after opening, takes the case's operations (a fill down, a copy, rows inserted, a sheet renamed), reads them again after a full calculation, and saves the book. The answers are in `test/excel/links/` and `test/excel/win/links/` with the files Excel saved; the predictions in each set were committed before Excel ran it.

| set | what it covers |
|---|---|
| `first` | one cell, the shape of the step |
| `basics-held` | cells, ranges, whole columns and rows, cells not stored, text, booleans and errors stored, sheets with a space and in Cyrillic, the linked book's names, this book's names over a link, lookups, criteria, `INDIRECT`, `OFFSET`, a dynamic array, a fill down, a copy, a renamed sheet, inserted rows, failed updates, two links, six forms of a target |
| `edits-held` | which edit makes a cell not stored `#REF!`: none, fills and copies onto stored and unstored cells, a rename, inserted rows |
| `refresh-held` | the mark on stored values, on empty data, a sheet without data, and which sheets and links a fill marks |
| `functions-held` | the functions that take only a reference, a linked criterion, a linked range through a name, more functions that read values |

Both Excels gave the same values everywhere. `test/links-excel.test.ts` checks, for each platform, the values after opening against `compute` of the case's book, the values after the operations against a session that takes them, the formulas Excel showed against `toLocal` with the links' targets in the opened file's folder, and that Excel's answer was the predicted value or one of its alternatives; the predictions it did not confirm are named there with the reason (a cell not stored was predicted blank, and is, until a formula with the link is entered). Excel also built a link itself, to a book it had open: it stored only the non-empty cells of the source's used part, an empty text as `t="str"` with an empty value, and opened without the source gave the stored values, blank elsewhere. `test/links-session.test.ts` runs an edit under formulas with links, an entered link and a refused new link in a synchronous session, in the same thread and in a pool, in both execution modes, against a full computation of the session's book. `test/links.test.ts` covers the text, the check, the input check, the input hash and structure operations.

### Live workbook operations and bounded queries

From [Live workbook operations and bounded queries](live-books.md).

Find was measured in Excel 16.0.20430 for Windows through its object model (`Range.Find` and `FindNext`, step `find` of `bench/windows`), in en-US and in ru-RU, on three sets held out with their predictions committed before Excel ran them (`bench/coverage/find/`, answers in `test/excel/win/find/`, test `test/find-win.test.ts`):

| set | cases in each locale | predicted by | confirmed | what it settled |
|---|---:|---|---:|---|
| `rules-held` | 155 | the grid's own model: the lookup functions' wildcards, Look in Formulas `editText`, Values the shown text | 136 | wildcards and `~`, `?` over a character outside the BMP, the case of letters, what Look in reads, merged areas and arrays, order, hidden cells on a filtered sheet |
| `rules-held-2` | 64 | `findCells` | 45 | hidden cells on a sheet without a filter, ranges of several areas, narrow columns, more of Look in, the Greek sigma |
| `rules-held-3` | 32 | `findCells` | 32 | a filter's mode makes Look in Formulas skip hidden cells and take the blocks, even when it hides nothing |
| `rules-held-4` | 29 | `findCells` | 22 | Find from a cell (`after`): round to it, a cell outside the range refused, a hidden row without a filter; on a filtered sheet a cell of a hidden line stands in the first block, and by columns the first column of the first block is not come back to |
| `rules-held-5` | 24 | `findCells` | 15 | Find from cells of later blocks and of hidden lines further on a filtered sheet goes on as `rules-held-4` showed, by rows and by columns, in ranges that begin past the first block too |

`findCells` gives Excel's cells in Excel's order in all of them but, in each locale, the seven cases of `rules-held`, the four of `rules-held-4` and the nine of `rules-held-5` listed in [Known differences](live-books.md#known-differences). Excel for Mac's Find was not measured. `test/session-find.test.ts` checks the options in a synchronous session, in the same thread and in a pool's thread, with closures and eval.

Replace was measured the same way (`Range.Replace` with Look in Formulas and dynamic formulas, `FormulaVersion` 1, step `replace` of `bench/windows`), each case in a book of its own, its cells read and the book saved and read back, with the columns' widths from `rules-held-2` (`bench/coverage/replace/`, answers in `test/excel/win/replace/`, test `test/replace-win.test.ts`):

| set | cases in each locale | predicted by | confirmed in en-US, ru-RU | what it settled |
|---|---:|---|---:|---|
| `rules-held` | 60 | the grid's own Replace All: Find's cells in formulas, every match with `*` as long as it can be, written by `setCellTexts` | 51, 50 | wildcards in the query and in what is put, `?` over a character outside the BMP, the case, the whole cell, `*` as short as it can be, every cell found entered again, percents and dates, formulas in the book's language and ones that no longer read, arrays whole, hidden cells, a range of one cell |
| `rules-held-2` | 22 | `replaceInCells` | 20, 20 | where `*` stops, an apostrophe put, a percent begun by a sign, the text format, the list separator, a table's header, no widening |
| `rules-held-3` | 10 | `replaceInCells` | 9, 9 | a cell of the text format takes a formula and a boolean as typed elsewhere, and a date with a date format; percents, dates' and money's separators |

`replaceInCells` leaves each cell as Excel left it in all of them but the cases listed in [Known differences](live-books.md#known-differences): three of `rules-held` in en-US and four in ru-RU. `test/session-replace.test.ts` checks the rules, the undo and the options in every kind of session, and `test/find-texts.test.ts` the texts kept after a replacement.

`Range.End` was measured the same way (step `end` of `bench/windows`), each case a cell and a direction in a book Excel opened, computed, filtered and typed into, the address it came to; held out with `getEnd`'s predictions committed before Excel ran each set (`bench/coverage/end/`, answers in `test/excel/win/end/`, test `test/end-win.test.ts`):

| set | locale | cases | confirmed | what it settled |
|---|---|---:|---:|---|
| `rules-held` | en-US | 167 | 167 | blocks, one empty cell apart and further, to every edge; which cells hold something (14 kinds); spills, a blocked one, an old-kind array, a spill of a whole column; merged areas from each side and from inside; rows and columns hidden by hand and by a filter; every line hidden past a row and a column |
| `rules-held-2` | en-US | 19 | 19 | from a cell of a hidden row or column, with a value and without; a row of height 0 and a column of width 0 not marked hidden; a merged area whose top-left cell is in a hidden row |
| `ru-held` | ru-RU | 4 | 4 | blocks, an empty text from a formula, a hidden row and a merged area as in en-US |

`getEnd` gives Excel's cell in all of them, in a synchronous session, in the same thread and in a pool's thread. The keyboard on a screen was not measured: Excel answered through its object model over SSH, without a screen. `test/end.test.ts` checks `getEnd` against a walk cell by cell over the session's result and book on seeded books with spills, empty texts, looks and hidden lines, through edits of cells, of the layout and of the structure and an undo, in every kind of session, and its time on the synthetic book.

### Locales

From [Locales](locale.md).

`test/excel-locales.test.ts` computes the measurement's book with `locale: 'ru-RU'` and compares every whitelisted function's cases with Excel ru-RU's answers in `test/excel/coverage/excel-ru.json`. `test/excel-locales-sets.test.ts` does the same for the ru-RU case sets in `test/excel/cases/ru-RU/`, among them `locale` (1,109 cases) and `collation` (174). `npx tsx bench/coverage/run.ts probes <set> ru` computes a set in Excel ru-RU. Known differences are listed in `RU_KNOWN` with a line each in the README; among them are the cases where the engine follows Excel for Windows and the Mac answers otherwise (the order of text, month names).

`test/excel-win.test.ts` checks the same books against Excel for Windows (`test/excel/win/`, `bench/windows`), with the set `collation-more` (62 cases of the order inside Cyrillic and Latin) measured in Windows alone.

`test/locale-excel.test.ts` checks that both translations reproduce Excel's own text of all 6,381 formulas in ru-RU, and `test/locale-roundtrip.test.ts` that 15,266 formulas of the case sets come back unchanged from every locale.

`test/locale-modules.test.ts` loads nothing until its last part: a ru-RU book is the `InputError` that names the module in `compute`, a session, `editStructure`, `fitColumns`, the same thread and `writeXlsx`, and so is the tag in every function that takes a locale; `ruRU` as an object gives these functions their answers without loading; an en-US book gives Excel's `Январь` for `[$-419]mmmm` and Excel for Mac's widths of Cyrillic letters; `readXlsx` keeps a file's ru-RU; once loaded, a ru-RU book computes as before in a session and in the same thread, with closures and eval; `registerLocale` refuses another object with a loaded tag, en-US, and a locale without rules or with wrong messages; the files of `src/locales/`, `PACKAGE_LOCALES`, `LOCALE_MODULES` and the subpaths agree. `test/locale-args.test.ts` gives every function with a locale the tag and the object and gets the same answer, and `test/locale-pool.test.ts` checks the threads of a pool. `test/bundle-size.test.ts` bundles `compute` alone with esbuild: 253 KB minified, with en-US alone of the locales, under its budget of 300 KB.

Excel for Mac and Excel for Windows (16.0.20430) compute ru-RU alike but for their regional data. Where those part, the engine follows Windows: the order of text, month names, the times of built-in formats, `CHAR` and `CODE`. It keeps the Mac's long date `[$-F800]` (`пятница, 31 января 2025 г.`; Windows's has no weekday). Formula text in ru-RU (array constants and the like) was not compared with Windows.

### Number formats and shown text

From [Number formats and shown text](formats.md).

`test/column-width.test.ts` checks the exported pair against all 666 existing Mac fit answers, the default, every whole point through 1,535 and invalid inputs. `bench/coverage/width` saves the current model's predictions before the object-model probe: ten fixed widths from the file, the default and seven `ColumnWidth` settings, with each platform's `Width`, `ColumnWidth`, `StandardWidth` and the width Excel saved. `test/column-width-excel.test.ts` checks the stored point widths without any screen measurement. Both platforms matched 17 of 18 predictions: only `ColumnWidth = 0.5` differed (6 points on Mac, 4.5 on Windows, against 8 and 6.75 predicted). All ten fixed widths and both defaults held. Setting characters is a different conversion from setting a file width; the exported inverse takes points instead.

`npx tsx bench/coverage/run.ts display <set> [ru]` writes a book of values or formulas with their formats, one column per width, and merged areas where a case needs one. Excel reads back each cell's `string value` (its `Range.Text`) and `color of font object of display format`. Answers are in `test/excel/display/` and `test/excel/display/ru-RU/`, and `test/display-excel.test.ts` compares the engine with them.

| set | cells | what |
|---|---|---|
| `widths` | 1,523 | character widths, column room, fractional and default widths |
| `general` | 5,166 | 82 numbers in General, columns from half a character to 14 characters |
| `codes` | 1,674 per locale | 99 number formats on 12 values, 38 date formats on 15 values |
| `fit` | 678 per locale | booleans, errors and text in narrow columns, numbers near the edge, fills, merged areas, colours, date separators, calendar edges |
| `names` | 70 per locale | month names, long dates and times, large fractions, values no section takes |
| `held` | 58 per locale | predicted by the engine before Excel computed them; 57 of 58 matched in each locale, and the miss became the rule for values no section takes |
| `builtin` | 300 per locale | the built-in formats 0 to 49 by number on six values, with the codes Excel reports for them |
| `builtin-held` | 350 per locale | built-in formats on other values and widths, predicted by the engine before Excel computed them, all matched |
| `seconds-held` | 35 (en-US) | moments at the edge of a second in `d h:mm:ss`, predicted by the engine before Excel computed them; 15 missed (the rounding to the second, and the 15 digits a cell's text is written from, a known difference), Windows the same |

`TEXT` was measured as formulas in the case sets `text-format` and `text-format-more` (en-US) and `text-format-ru` and `text-format-ru-more` (ru-RU), 771 cases in all; a moment's whole seconds in the set `seconds-held` (35 values next to `SECOND` and `DAY`, the same in Excel for Mac and for Windows).

The cases write a format as a custom code, so a code equal to a built-in format does not show by the locale. The sets of built-in formats write the number alone, as Excel stores it.

The 1904 date system was checked with the workbook set `date1904` (`bench/coverage/workbook/`): the same cells and `TEXT` formulas in a book of each system.

Excel for Windows showed the same sets (`bench/windows`, `test/display-win.test.ts`). Its regional data, where they part from the Mac's, are the engine's: the built-in codes, month names in ru-RU, the times of built-in formats. The cells that show otherwise meet a column's edge, where Windows, laying text out in pixels, shows what the engine shows at another width, and the ru-RU long date `[$-F800]`, which the engine keeps as the Mac's. `test/display-excel.test.ts` covers the Mac's answers in the places where the engine follows Windows (`MAC_SHOWN`).

AppleScript returns no colour for an exact 0 in a format of fewer than three sections (`[Red]0`) and for text under a coloured text section (`[Red]@`), though Excel draws the colour on screen there. The engine gives the section's colour in these places, and the comparison accepts any colour.

Fitting columns was measured by `npx tsx bench/coverage/run.ts fit <set> [ru]`: a book of one column per case, which Excel fits (`autofit` of the entire column) or types the case's texts into in turn, through `value`, `formula2 local` or the keyboard; then it reads each column's width and what its cells show and saves the book, whose `<col>` elements give the width Excel stores. Answers are in `test/excel/fit/`, and `test/fit-excel.test.ts` compares `fitColumns`, `parseEntry` and `entryWidth` with them, width and shown text.

| set | columns, en-US and ru-RU | what | predicted before the run |
|---|---|---|---|
| `autofit` | 161 and 161 | text of every width from 3 to 60 points and wider, numbers in General and by formats, booleans, errors, several cells, empty columns, merged areas, a hidden row and column, fonts, formulas and a spill | 11 and 11 |
| `typed` | 87 and 82 | numbers, dates, money, percents and formulas typed through the object model and on the keyboard, into columns of the default width, set by hand and fitted, and in turn | 21 and 19 |
| `glyphs` | 332, en-US | each character of the table alone and in a run of about 600 points, Russian words | 253 |
| `held` | 86 and 85 | held out from the others: text near 90 points and outside ASCII, General near its own text, paddings and blanks, typing that fits exactly or drops decimals | 76 and 74 |

The first sets showed the rules: the predictions of `autofit` and `typed` came from a guess (text width plus the padding, no widening of a General number); `glyphs` showed that Excel measures text outside ASCII by the font's advances. Of the misses in `held`, five were typed through the object model, which the prediction widened though `typed` had shown it does not; the others became rules: a General number that fits exactly widens its column, a whole number's blank fraction keeps the room of its digits and slash, characters of Latin-1 have their own widths, a fitted width equal to the default stays. The engine now gives Excel's width and text on all 994 columns but the 67 in `test/fit-known.ts`, for the reasons below.

Excel 16.0.20430 for Windows fitted the same books and typed the same texts through `Value2` and `Formula2Local` (`bench/windows/run.ts fit <set>`, `test/excel/win/fit/`, `test/fit-win.test.ts`; 912 columns, the keyboard's cases being the Mac's alone). It does what the Mac does: it fits a column with something to show and marks it `bestFit`, shows a hidden column, leaves a column with nothing to show, and widens nothing for text set through the object model. Its widths are its own: it lays text out in pixels, a digit 7 at 96 dpi, and saves the width in sevenths of a digit, so of the columns it fitted only 16 of 161 (`autofit`, either locale), 34 of 332 (`glyphs`) and 15 of 71 (`held`) have the Mac's width. 1234567890 gets 11 on Windows and 11.1640625 on the Mac, a long date 21.7109375 and 19.6640625, 300 points of `b` 59.85546875 and 52.5.

Fitting rows was measured by `npx tsx bench/coverage/run.ts rows <set> [ru]`: a book of one row per case, its cells in columns of their own, which Excel reads the height of when it opens it, types the case's texts into (through `value`, `formula2 local`, or with Excel in front on the keyboard, Control+Return or Cmd+V) and fits (`autofit` of the entire row); then it reads each row's height and its cells' content and wrapping and saves the book, whose `<row>` elements give the height Excel stores. Answers are in `test/excel/rows/`, and `test/rows-excel.test.ts` compares a session with them, height, content and wrapping.

| set | rows, en-US and ru-RU | what | heights predicted before the run |
|---|---|---|---|
| `autofit` | 129 and 129 | an empty row, Calibri of 25 sizes, other fonts, values, several cells and the lines' styles, line breaks, where a wrapped line breaks, long text over six widths, heights set by hand, a hidden row, merged cells, text with a line break set through the object model | 50 and 50 |
| `held` | 99 and 99 | Calibri of 20 more sizes, wrapped lines of 15 sizes, other fonts wrapped, the row's style smaller than the default, a wrapped cell beside a larger font, a line of 16 points, spaces and a hyphen at the break, typing into wrapped and merged cells | 81 and 81 |
| `lines` | 118 and 118 | 1 to 26 wrapped lines of 11 points and lines of 6 to 24 points, half sizes, a wrapped cell beside fonts from 8 to 36 points, lines of 20 and 24 points | 79 and 79 |
| `keys` | 19, en-US | on the keyboard: line breaks typed into a cell, a selection (Control+Return), rows set by hand, a larger font, a merged cell, formulas over lines; text with line breaks pasted, inside quotes and not | 19 |
| `quotes` | 12, en-US | quotes in pasted text: a tab inside, text after the closing quote, a quote inside, a space before | 12 (the heights; the cells were predicted without quotes) |

The first set came from a guess (a line of `1.2207·s` points, the wrapped line of shown text); its answers gave the line, the room of a wrapped line and that the object model wraps text with a line break. `held` showed rows higher than their highest cell and half points, which `lines` resolved into the line cells stand on. The engine now gives Excel for Mac's height, content and wrapping on all 723 rows of both locales but the 14 in `test/rows-known.ts`: a font other than Calibri, and a formula typed over lines, whose line breaks Excel keeps in its text.

Excel 16.0.20430 for Windows ran the same books (`bench/windows/run.ts rows <set>`, `test/excel/win/rows/`, the keyboard's cases but the pastes being the Mac's alone; Cmd+V is `Worksheet.Paste` there). It does what the Mac does with the cells: text with a line break set through the object model or pasted inside quotes wraps, a paste fits no row, AutoFit shows a hidden row and drops a height set by hand, merged cells count as on the Mac. Its heights are its pixels': a line of Calibri 11 is 20 pixels, 15 points, wrapped or not, 16 points 28 pixels (21), and 49 of the 129 rows of `autofit` have the Mac's height.

The ribbon was read by `npx tsx bench/coverage/run.ts ui ribbon [ru]`, the step `ui` with Excel in front (`bench/coverage/README.md`). Excel selects each cell of a book and reads the Number Format box through Accessibility, after selecting a control cell in the text format that the box must follow; then it applies each item of the list to a cell of its own and reads the item's name, the box after and the format the saved file holds. The answers are in `test/excel/ui/ribbon.json` and `test/excel/ui/ru-RU/ribbon.json`, and `test/ui-excel.test.ts` compares the functions with them, a format by what it shows (44 and the code Excel writes for it alike).

| cells, en-US and ru-RU | what |
|---|---|
| 50 | the built-in formats 0 to 49 by number, as an xlsx file stores them |
| 50 | the same formats by their en-US codes as codes of the file's own; Excel takes a code equal to a built-in format's as that format (`0.00` saves as 2) |
| 17 | ru-RU's codes of the built-in formats |
| 193 | variants: decimals, groups and negatives, currencies of each place and form, accounting, dates and times with and without `;@` and locale tags, fractions, exponents, text, General with text, hidden formats, colours and conditions, the special formats |
| 32 | eight formats on text, a blank, a negative number and a date |
| 26 | the list's items applied, to cells in General and in other formats |

The first rule, committed before Excel ran the set, named the box right for 179 of 342 en-US cells and 110 ru-RU ones; it guessed the categories by the shape of a code. The en-US rule above came from the answers and gives Excel's box for every cell but the 7 en-US cells of `m/d/yyyy`, and Excel's items but the date and time items of the en-US run (below).

Excel 16.0 for Windows read the same books in ru-RU (`npx tsx bench/windows/run.ts ribbon ribbon ru`, the Windows step `ribbon`, `bench/windows/README.md`): it shows its ribbon only on a screen, so the step runs as a task in a session logged on through Remote Desktop and reads the box by UI Automation, selecting a control cell in the text format before each cell as on the Mac. The cases that apply the list's items are the Mac's alone. Windows named 266 of the 342 cells as the Mac; each of the 76 others is `(все форматы)` on the Mac and a category on Windows (23 Числовой, 15 Дата, 12 Денежный, 9 Финансовый, 17 others). A held-out set, `ribbon-ru-held` (50 cells: what the Format Cells window of ru-RU might list, dates with `[$-419]` and `[$-FC19]`, the ruble in more forms, fractions with the space quoted), ran after it. The answers are in `test/excel/win/ribbon/ru-RU/`, and `test/format-category-win.test.ts` compares the function with them.

| run | rule committed before it | cells named as Excel for Windows |
|---|---|---|
| 50 codes set through the object model (dropped: COM reads `$` in a code as the locale's currency) | en-US's with the locale's names and short date | 44 of 50 |
| `ribbon`, 342 cells | the first run's misses made rules | 318 by the prediction, 342 by the rule above |
| `ribbon-ru-held`, 50 cells | the same | 29 by the prediction, 34 by the rule above; the 16 others are dates and times with `;@` (below) |

### Outlines

From [Outlines](outline.md).

`npx tsx bench/windows/run.ts outline <set> [ru]` writes each book of `bench/coverage/outline/<set>.ts` with `writeXlsx`; Excel for Windows opens it, takes the case's commands through COM (`Range.Group`, `Range.Ungroup`, `ShowDetail` set True or False on a line, `Hidden` set by hand, `Outline.ShowLevels`) and, before the first command and after each one, reads every line of the window (`OutlineLevel`, `Hidden`, `Summary`, and a summary line's `ShowDetail`) and saves a copy of the book, whose `<row>`, `<col>`, `<outlinePr>` and `<sheetFormatPr>` it reads. `test/outline-win.test.ts` checks, without Excel, that the engine's book after each step gives the same lines, `outlineGroups` giving a summary line's `ShowDetail`, that `writeXlsx` writes what Excel saved, that each refusal is one of Excel's, and that the reader reads the book Excel saved at the end with the groups the engine holds.

| set | cases | what |
|---|---|---|
| `rules-held` | 87 | Group and Ungroup at every level, beside and over groups, at the deepest level and with none; Show and Hide Detail at summary lines, at lines of groups and at several lines, nested groups in every order, groups sharing a summary line; level buttons; rows hidden by hand; summaries above and left; groups at the sheet's edges; rows a file holds hidden without a mark, and a mark over rows shown |
| `ru-held` | 4 | the same commands in ru-RU |
| `rules-held-2` | 41 | held out after `rules-held`: Group beside the deepest level, Show and Hide Detail at a line of a group and at a summary line three groups share, marks under rows hidden and shown by hand, a file's mark inside a hidden group, level buttons over three levels and on both axes, columns at either edge |
| `rules-held-3` | 18 | held out after `rules-held-2`: Group and Ungroup at XFD and the last row, level buttons over columns up to XFD, marks under rows shown by hand |

The predictions were committed before Excel ran each set (`bench/coverage/outline/*.predicted.json`). For `rules-held` and `ru-held`, before the engine had the commands, they were the grid's own model (`xlcalc-sheet`'s `src/local/outline.ts`, ported into `predict.ts`): Excel confirmed 49 of 87 and 2 of 4 cases. The misses gave the rules above: the mark is Excel's own state, which the button's `ShowDetail` reads, kept on a summary row inside a hidden group and set whenever a detail is all hidden; Show and Hide Detail take one line and act on a summary line's whole detail; refusals where the button would not act; Group leaves lines at level 7. The engine predicted `rules-held-2`, 37 of 41 confirmed (columns up to XFD and one sequence by hand), and `rules-held-3`, 9 of 18 (a rule at XFD it tried, which Excel refuted, and two sequences by hand), all now known differences above. Excel for Windows 16.0.20430 ran every set, `rules-held` to `rules-held-3` in en-US and `ru-held` in ru-RU, with `XLWIN_CODEPAGE=any`, on 12 October 2026.

### Reading and writing xlsx

From [Reading and writing xlsx](xlsx.md).

`npx tsx bench/coverage/run.ts xlsx <step>` sends files to Excel 16.113 for Mac through AppleScript. Excel opens each file, computes it fully and saves it under another name. The repair window "We found a problem with some content" would mean a broken file, and none appeared. The saved files are in `test/excel/xlsx/`, and `test/xlsx-excel.test.ts` checks them without Excel.

| step | what | result |
|---|---|---|
| `roundtrip` | 181 books written by `writeXlsx`: values, escapes, formulas, spills, styles, sheet names, both locales, the synthetic book, the books of the structure, spill and table sets (names of any formula, `@`, references across sheets, array formulas of the old kind and tables with structured references among them), hidden rows and columns with their sizes, lines of size 0, filtered sheets with and without a table, the 1904 date system, iterative calculation and the built-in formats by number in en-US and ru-RU | every file opens, every book reads back as written, and Excel's values equal the engine's |
| `filters` | 8 files with a filter in each form a file holds it, `SUBTOTAL` and `AGGREGATE` over the rows it hid | read with the filter mode Excel took, every formula computes the value Excel stored |
| `saved` | 97 books of the case sets, 29,089 cases, as the harness sends them to Excel | read back, every case computes as in the book sent, and the values Excel stored equal its stored answers |
| `intersect` | 75 formulas of the old kind | read with `@` where Excel puts it, every formula computes the value Excel stored, but the two with `OFFSET` |
| `colors` | 233 fills: theme colours with Excel's tints and the legacy palette | 212 equal to Excel's colour, the rest within 3 of 255 per channel |
| `formats` | the built-in number formats 1 to 58 | the codes Excel en-US shows equal the engine's for 1 to 49 and the reader's for 50 to 58 |

Without Excel, `test/xlsx-issue-locale.test.ts` reads a file with 28 issues (26 messages of the catalog) and the 15 files Excel saved that have issues in en-US and in ru-RU, and writes a book with 8 losses in both: each Russian message is the catalog's for the issue, the codes, sheets, cells and names equal the English read's, and without the option the issues are the English ones. `test/messages-catalog.test.ts` checks that every message has its Russian text, which shows each parameter the English one shows. `test/xlsx-snapshot.test.ts` checks the three kinds of `values`, the refusals, `cached` in the shape of a result and the files written again from it, a file Excel saved among them, and `test/session-snapshot.test.ts` writes the snapshot of a synchronous session, one in the same thread and one in a thread, and reads it back with the session's values.

The 934 files of `test/excel/` were read by the reader of looks and by the one before it (`a0fa646`): every book, its looks as entries, every result of stored values and every issue were the same, and so were the files written again from them, part for part. The book of the wide shape read the same, and its windows, cells and used ranges in a session and the file written of it were those of the book of entries. `test/looks.test.ts` writes books with looks and the same books with entries from every kind of session: the same parts, and the looks read back as the session holds them.

The round trip compares books by what they mean: a style by its fields, a formula by its canonical text, spans of rows and columns line by line, a line of size 0 as hidden, a code that every locale shows as a built-in format as that format's number.

The 98 files of the workbook sets (`bench/coverage/workbook/`: hidden rows and columns, autofilters with and without criteria and the filter mark, the 1904 date system, iterative calculation), which Excel opened and computed step by step, read into the books the engine was checked with against Excel's answers, and compute the same values (`test/xlsx-workbook.test.ts`).

The answers to the filter files and to nine books of the round trip were predicted and committed before Excel ran them (`bench/coverage/xlsx/predicted.json`): the filter mode of 8 files and 118 values, all as Excel gave them. The books were predicted to come back without differences, and three did not: Excel saved the built-in formats of two by codes and escaped the literal characters of the date formats of the third.

#### Hidden sheets

`npx tsx bench/coverage/run.ts hidden <set>` writes each book of `bench/coverage/hidden/<set>.ts` with `writeXlsx`; Excel for Mac opens it, notes the active sheet, hides and shows sheets through its object model (`visible`), applies a sheet operation (rename, delete, insert, move), reads each sheet's visibility, the active sheet and the values of formulas over the sheets, and saves the book. `npx tsx bench/windows/run.ts hidden <set>` does the same in Excel for Windows (`Worksheet.Visible`). `test/hidden-sheets-excel.test.ts` checks, without Excel, that Excel read what the engine holds after each case and that the reader reads every file Excel saved with the visibility Excel read.

| set | cases | what |
|---|---|---|
| `sheets-held` | 29 | books written with hidden and very hidden sheets, the first ones too; sheets hidden, very hidden and shown by Excel; the last visible sheet hidden and deleted; hidden sheets renamed, deleted and moved, sheets inserted and moved beside them; two files changed by hand, one whose active tab is a hidden sheet and one whose every sheet is hidden |
| `sheets-held-2` | 11 | sheets moved before two hidden sheets in a row, a very hidden one and a hidden last one, a hidden sheet moved, sheets inserted before hidden ones |
| `sheets-held-3` | 10 | ends of references across sheets moved inward, outward and past each other, sheets moved past hidden ones at the start and the end |

Both Excels opened every file the writer wrote without the repair window, read each sheet's state as written, showed the first visible sheet, and computed a reference across sheets through a hidden one and `SUBTOTAL` and `AGGREGATE` over a hidden sheet's hidden row as over visible ones. A file whose active tab names a hidden sheet opened with that sheet shown, and one whose every sheet is hidden was repaired. The engine's answers were committed before Excel ran each set (`bench/coverage/hidden/<set>.predicted.json`): `sheets-held` held in 25 of 29 cases on both Excels, `sheets-held-2`, made after it, in 6 of 11, and `sheets-held-3`, made after that, in 10 of 10. The misses became the rules of [docs/structure.md](structure.md#hidden-sheets): a very hidden sheet is neither deleted nor moved, a moved sheet is placed by the visible sheets, and an end of a reference across sheets that moves stays its end until it passes the other. Both Excels answered alike in every case; Excel for Mac refused some deletes without an error, leaving the sheet.

### Reading and writing xlsx: Data validation and outlines

From [Reading and writing xlsx, Data validation and outlines](xlsx.md#data-validation-and-outlines).

The valid sets of `bench/coverage/valid/` were written by `writeXlsx` and opened by Excel for Mac and for Windows: every book loaded, and Excel read the rules, each cell's check and the outlines the engine holds but for texts cut to Excel's limits. Rules and outlines Excel made itself, and the books Excel saved after structure operations, read back as Excel read them ([Data validation](#data-validation)).

### Reading and writing xlsx: Links to other workbooks

From [Reading and writing xlsx, Links to other workbooks](xlsx.md#links-to-other-workbooks).

`npx tsx bench/coverage/run.ts links <set>` sends each case of `bench/coverage/links/<set>.ts` to Excel for Mac, which opens the file without updating its links (as a person who answers "Don't Update" does), reads the case's cells, takes its edits, computes the book fully, reads the cells again and saves the book; `npx tsx bench/windows/run.ts links <set>` does the same in Excel for Windows. The sets `first`, `basics-held`, `edits-held`, `refresh-held` and `functions-held` are files put together by hand around what `writeXlsx` writes, so they do not depend on the writer; the set `written` is files `writeXlsx` writes from a book with `externalBooks`.

| set | cases | what |
|---|---|---|
| `written` | 7 | the cases of `basics-held` but one (a link with every kind of stored value, its names and this book's names over it, lookups, a dynamic array, fill downs and a copy that mark the link; marked data and a sheet without data; two links; the forms of a target); cells stored empty and empty text, with the mark and without; #SPILL! and #CALC! stored; a link without sheets and with an empty target between two others |

The values of `written` were predicted and committed before Excel ran it. In the first run Excel for Mac opened six of the seven files and showed every predicted value, the same as for the files made by hand; it repaired the file whose link stored #SPILL! and #CALC! as their text. Written as `#VALUE!` with their rich values, the file opened and showed `#VALUE!` for both, the alternative the case gave, and Excel saved them as plain `#VALUE!`; the writer now writes that. In the last run both Excels opened all seven files and showed 72 of the 76 predicted values, the four others as their alternatives. Each case of `basics-held` read and computed as its file made by hand, before the edits and after, on both platforms, with the same formulas and shown text. Excel for Windows answered as Excel for Mac in every read.

`test/xlsx-links-excel.test.ts` checks this without Excel, and that the reader reads every link of the 48 files both Excels saved in the six sets: the targets each wrote (Excel for Mac keeps a relative target, or writes the absolute path it found beside the relative one in `alternateUrls`; Excel for Windows writes an absolute path without the drive), the marks and the stored values Excel kept, and no link reported as one the model does not have. Every book of `written` reads back from its file with the same links, names, formulas by their canonical text and input hash. `test/xlsx-links.test.ts` reads links made by hand in every form above and writes the parts as Excel writes them.

### Reading and writing xlsx: Notes, hyperlinks and frozen panes

From [Reading and writing xlsx, Notes, hyperlinks and frozen panes](xlsx.md#notes-hyperlinks-and-frozen-panes).

`npx tsx bench/coverage/run.ts parts <set>` writes each book of `bench/coverage/parts/<set>.ts` with `writeXlsx`. Excel for Mac opens it, makes what the case makes through its object model (`add comment`, `make new hyperlink`, `freeze panes`), applies the case's structure operation, reads every sheet's notes, hyperlinks and panes, and saves the book. `npx tsx bench/windows/run.ts parts <set>` does the same in Excel for Windows through COM (`AddComment`, `Hyperlinks.Add`, `FreezePanes`). `test/sheet-parts-excel.test.ts` and `test/sheet-parts-win.test.ts` check, without Excel, that Excel read what the engine holds after each case and that the reader reads every file Excel saved as Excel read it.

| set | cases | what |
|---|---|---|
| `files` | 10 | books written with every kind of note, link and frozen pane, 1,100 notes on one sheet, a long tooltip and a long address, books written after a structure operation in the engine; notes, links and frozen panes Excel made itself |
| `struct-held` | 26 | rows and columns inserted and deleted around frozen panes, notes and links; ranges moved and copied with them and onto them; sheets renamed and deleted under links to them |
| `frozen-held` | 9 | frozen rows and columns deleted past the split and inserted several at once |
| `threads` | 2 | threads of comments put into a file as Excel 365 writes them, with a row inserted above them |
| `refs-held` | 5 | a note's cell as another writer may spell it, put into the file by hand: `A01`, `A1:A1`, `a1`, `$A$1`, `A1:B2` |

Both Excels opened every book without the repair window but the file of `refs-held` with a note on `A1:B2`, which neither opens. Excel for Mac refused the first version of the book with a tooltip of 300 characters; since then the writer cuts tooltips to 255, and both open it. The engine's answers to `struct-held`, `frozen-held` and `threads` were committed before Excel ran them (`bench/coverage/parts/*.predicted.json`), and so were those to `files`, which held but for the address with `#` and the long tooltip: `struct-held` held in 24 of 26 cases, the two others showing that deleting every frozen row or column leaves one frozen; `frozen-held`, made after that, held in 9 of 9; `threads` in 1 of 2, the other being the inserted row below; `refs-held` in 1 of 5, as the engine kept a note on `A01` and `A1:A1` under those keys and left out one on `$A$1` and one on `A1:B2` without a word, where both Excels showed the first three on A1 and saved them so. Excel for Windows read what Excel for Mac read in every case, but for the author of the notes each made itself, its own user.

Neither Excel takes threads of comments into its object model: Excel 2024 for Windows refuses to make one through COM, Excel for Mac's AppleScript has none, and both show the placeholder note of a thread they open. After inserting a row both moved the placeholder note and left the thread at its cell, which is why the reader follows the placeholder. A file Excel 365 made with threads was not available, so the threads were put into the file by hand (`bench/coverage/parts/threaded.ts`).

### Reading and writing xlsx: The cells of formatted rows and columns

From [Reading and writing xlsx, The cells of formatted rows and columns](xlsx.md#the-cells-of-formatted-rows-and-columns).

`npx tsx bench/coverage/run.ts lines <set>` writes each book of `bench/coverage/lines/<set>.ts` with `writeXlsx`; Excel for Mac opens it, takes the case's steps through its object model (a number format, bold, italic or a fill set on a column, a row or a cell, a value typed into a cell), reads each listed cell's text, number format and the fill, font colour, bold and italic of its display format, and saves the book. `npx tsx bench/windows/run.ts lines <set>` does the same in Excel for Windows through COM (`NumberFormat`, `Font`, `Interior`, `Value2`, `Range.Text`, `DisplayFormat`). A bare case sends the file without the cells the writer writes where lines cross, as other writers leave it. `test/lines-excel.test.ts` checks, without Excel, that the engine shows what Excel showed after each case, on the book written or on the one the reader reads from a bare file, that it shows the same on the book read from the file Excel saved, and that a book written and saved by Excel reads back as written.

| set | cases | cells | what |
|---|---|---|---|
| `files` | 9 | 586 | formatted columns and rows over numbers, dates, text, booleans, errors, formulas and a spill without a format of their own, cells with their own style or General, empty cells with a style; rows with only a style or a format crossing columns with the other, written and bare, values typed there; Excel formatting a column, a row and lines across each other, values typed into them |
| `held` | 6 | 501 | spans of rows and columns, a row with only a style crossing columns with a format and a style, dates, a spill across formatted columns, cells with their own General, plain style or italic, written and bare, values typed where lines cross; Excel formatting a row and then columns across it |

The engine's answers were committed before each run. The first run of `files`, on the writer that did not write the cells where lines cross, held in 314 of 320 cells: four were empty or typed cells where a row with only a style or a format crosses a column with the other, which both Excels showed by the row's whole cell format, and two were the cells where Excel formatted a filled column across a bold row, which it gave both. The writer then wrote those cells and the reader read a file without them as Excel shows it, and the second run of `files` held in 584 of 586 cells, the two being those Excel formatted; `held`, made after it, in 500 of 501, the other being the same merge. Excel for Windows showed what Excel for Mac showed in every cell of both runs, opened every file without the repair window, and the engine shows what Excel showed on every book read from a file Excel saved. In the fit set `autofit`, whose cases `colformat` hold a number in a formatted column, both Excels show the column's format and fit the column for it, as the engine does.

### Reading and writing xlsx: The saved active sheet and view

From [Reading and writing xlsx, The saved active sheet and view](xlsx.md#the-saved-active-sheet-and-view).

`test/xlsx-view.test.ts` reads explicit OOXML fields, then writes, reads, writes and reads them again. It checks the first visible sheet default, input validation, both hashes in both execution modes, all three kinds of session, and active-sheet changes through rename, hide, delete and undo. No default XML behavior was changed for books without saved view fields.

### Spilling

From [Spilling](spill.md).

Each case is a small book written the way Excel 365 saves dynamic array formulas (`cm="1"` and the `XLDAPR` metadata). Reading a range of values that contains errors is unsafe in Excel (errors drop out of the list, and a two-dimensional read crashed Excel), so every cell that is read has three mirror formulas on separate sheets that never fail: a type code, the number and the text. `npx tsx bench/coverage/run.ts spill <set>` runs a set, and `test/spill-excel.test.ts` compares the engine with the stored answers without Excel.

| set | cases | what |
|---|---|---|
| `base` | 50 | directions, blocking, collisions, edges, `E2#`, sizes from data, whole columns, volatile arrays |
| `order` | 21 | two arrays over the same cells, computed four times |
| `held` | 17 | predicted by the engine before Excel computed them; all 17 matched |
| `more` | 13 | sizes from a draw (five runs), edits in books with merged areas |
| struct `spill`, `spill-more`, `spill-copy` | 22 | `E2#` and spilled cells under structure operations |

Answers are in `test/excel/spill/` and `test/excel/struct/`. Cutting a formula with a spill to another sheet crashed Excel 16.113 twice, so it is not among the cases.

### Structure operations

From [Structure operations](structure.md).

`test/structure-geometry.test.ts` checks that every kind of operation puts the cells where its `mapCell` says: each cell it keeps is at the address `mapCell` gives and nowhere else, a cell it takes away is in no other place, and the way back takes each cell to the place it came from; `test/session-geometry.test.ts` checks `moved` and `mapCell` of a filtered delete and a deleted table done, undone and redone in every kind of session, and `null` for a sheet `mapCell` does not know. The seeded undo sessions (`test/session-undo-fuzz.test.ts`) keep a host's screen, moved by `mapCell` and given the changes, equal to the session's values.

`test/copy-formula.test.ts` checks `copyFormula` against `copyRange` on every copy of the structure sets and the table sets: each formula of the source at each place the copy pastes it, 317 and 8 of them.

`npx tsx bench/coverage/run.ts struct <set>` applies each case's operation to its book in Excel through AppleScript and stores what every sheet shows in `A1:Z60`, with the defined names, in `test/excel/struct/<set>.json`. For books with styles it also reads the look of `A1:J24`: bold, fill colour and number format of each cell, widths and heights that are not standard, and the merged areas. `test/structure-excel.test.ts` compares `editStructure` with the stored answers without Excel.

| set | cases | what |
|---|---|---|
| `rows` | 26 | inserting and deleting rows and columns with every kind of reference: relative, absolute, mixed, ranges, whole columns and rows, other sheets, quoted names, global and local names |
| `edge` | 10 | references at the last row and column, the whole sheet, the two inserts Excel refuses |
| `move` | 27 | cut and copy on the same book, onto referenced cells, whole columns and rows, other sheets, fill down and right, tiles |
| `moves` | 20 | every range `Ai:Aj` and `Ai:Bj` of a 12-row window under 20 moves |
| `sheets` | 9 | renaming (quotes, case, an apostrophe) and deleting sheets |
| `history` | 6 | renaming to `History` in three cases (Excel refuses) and to similar names (Excel accepts) |
| `checks` | 26 | predicted before Excel computed them: moves outside the fitted ones, edge cases, copies with mixed `$` and local names |
| `styles` | 46 | inserts, deletes, moves and copies in a book with cell styles, styled rows and columns, widths and merged areas |
| `styles-checks` | 18 | predicted before Excel computed them: spans, tiled and moved merged areas, cells moved into styled rows and out of styled columns |
| `edge-layout`, `edge-wide` | 9, 4 | styled rows and columns, spans and merged areas at the end of the sheet |
| `spill`, `spill-more`, `spill-copy` | 8, 4, 10 | spill references and spilled cells under structure operations |
| `hidden` | 49 | copies, fills, moves, inserts and deletes over rows hidden by hand and on a filtered sheet, and next to hidden columns |
| `cube` | 20 | references across sheets and names of any formula: rows and columns, sheets renamed, deleted, inserted and moved, copies |
| `arrays` | 17 | array formulas of the old kind: lines inserted and deleted around and inside the area, the area moved, copied and pasted over |
| `zero-held` | 8 | predicted before Excel ran them: rows inserted under rows of height 0, one styled, a span of two, one also marked hidden, and under a row hidden at height 30 |

The look also lists the hidden rows and columns. The set `hidden` writes its filtered book with an autofilter and its criteria, and runs its fills through Excel's Fill Down and fill handle as well as copy and paste.

All cases match except the two known differences below. Of the predicted cases, `checks` matched 21 of 26, `styles-checks` 17 of 18 and `zero-held` 8 of 8 (Excel for Windows too); the misses found two wrong rules and a conversion error in the comparison itself, all fixed.

Not checked against Excel: a move that does not fit (it is refused like a copy), most moves to another sheet, fonts other than bold, borders and alignment (they move with the style number). For merging: borders, fonts other than bold, the vertical alignment, and the notes, hyperlinks and `meta` of the cells a merge clears (here they keep them).

Copies were applied through AppleScript's `copy range` with a destination, Excel's `Range.Copy`. On a filtered sheet it pasted values. Copying and pasting in Excel's window was not measured.

Hidden sheets were measured by the step `hidden` in Excel for Mac and for Windows (`npx tsx bench/coverage/run.ts hidden <set>`, `bench/windows/run.ts hidden <set>`, `test/hidden-sheets-excel.test.ts`): sheets hidden and shown through the object model (`Visible`), renamed, deleted, inserted and moved beside hidden ones, with references across sheets over them. Each set's answers were predicted before Excel ran it: `sheets-held` held in 25 of 29 cases, the misses being the very hidden sheet that neither Excel deletes or moves and a sheet moved before a hidden one, which went after it; `sheets-held-2` in 6 of 11, the misses being moves past hidden sheets with none visible after them and ends of references across sheets that moved without passing the other end; `sheets-held-3` in 10 of 10. Both Excels answered alike ([Hidden sheets](#hidden-sheets)).

Merging cells was measured in Excel for Windows (`npx tsx bench/windows/run.ts struct <set>`: `Range.Merge`, `Range.UnMerge`, Merge & Center as `Merge` with the alignment centred, alerts off), whose answers are in `test/excel/win/struct/` and compared by `test/structure-win.test.ts`; Excel for Mac was not asked. The look read there has the horizontal alignment as well. A case may take an operation before its own (`before`): cells merged and split again show the look the merge gave them, and a table is made over the cells first (`ListObjects.Add`). Each set was predicted before Excel ran it. `merge` (28 cases) held in 20: an empty top-left cell takes the area's first value (6 cases, the prediction kept it empty), and Excel writes a merged area of whole rows as `20:21` (2, the same area). `merge-held` (18) held in 10: the value comes as cut and paste moves it, with the references to it and its look (8). `merge-third-held` (12) held in all 12, and `merge-top-held` (3), references to the empty top-left cell, in all 3. `test/merge-cells.test.ts` checks that a session gives the book `editStructure` gives on every case, with and without building the state again, and undoes and redoes it in one step in every kind of session.

The fill handle was measured in Excel for Windows by the same step (`Range.AutoFill` over the range `to`, its type `xlFillDefault`, `xlFillSeries` or `xlFillCopy`), in en-US, three sets each predicted before Excel ran it. `fill` (46 cases) held in 38: dates at months' ends go on by months, `Q1` by quarters, `1st` by ordinals, a number before text goes on, and a number alone beside text goes on by 1; three more cases differed only in the comparison, which now writes the engine's numbers in Excel's 15 digits. `fill-held` (45) held in 37: an empty cell does not end a run, dates and text without one step are copied, months go on from one day of the month or from months' ends alone, a number first needs a space after it and comes before one at the end, ordinals end small, a number alone is copied only when the source holds numbers alone. `fill-third-held` (31) held in 28: a date beside a number alone makes it go on, steps are taken alike in 15 digits (`9:00, 10:00, 11:00` go on by the hour), a number first loses its zeros (`007 x`, `8 x`). After those corrections all 122 cases match; no fourth set was run. `test/fill-range.test.ts` checks that a session gives the book `editStructure` gives on every case, undoes and redoes it in one step in every kind of session, and keeps dates given as text.

### Tables

From [Tables](tables.md).

`npx tsx bench/coverage/run.ts tables <set>` writes each case's book with `writeXlsx`, opens it in Excel, takes the case's steps (cells typed, structure and table operations) through AppleScript, and reads the values of the case's cells, their formulas as Excel shows them, in Russian too for a ru-RU set, and the tables of every sheet. `test/tables-excel.test.ts` replays the steps in a session and compares the steps refused, the tables, the values and the formulas, read as stored text and as `toLocal` shows them.

| set | cases | what |
|---|---|---|
| `refs` | 8 | every part of a table, the formula's row from everywhere, references without a table name inside one, no header row, a totals row, escapes, names, arrays |
| `refs-held` | 5 | predicted before Excel computed them: a table away from A1, the totals row, more escapes, names, arrays beside tables |
| `ops` | 51 | rows and columns inserted and deleted around tables, renames, conversions, deletions, sheets, header cells, copies and moves |
| `ops-held` | 18 | predicted before Excel took them: inserts beside another table, a table emptied, pastes over header rows, renames Excel refuses |
| `locale` (ru-RU) | 4 | the Russian specifiers shown and typed, a totals row, the names of new columns |
| `more` | 5 | whole tables copied and moved, cells typed beside and into a table |
| `names-held` | 17 | predicted before Excel computed them: defined names in letters of any script, tables renamed to Cyrillic names and to names Excel refuses |
| `names-ru-held` (ru-RU) | 8 | predicted before Excel computed them: the same names in formulas typed in Russian |
| `names-more-held` | 12 | predicted before Excel took them: defined names added through Excel's object model where its rules refuse, files with `?` and `\` in a name |
| `names-ru-more-held` (ru-RU) | 10 | predicted before Excel took them: the same in ru-RU, and files with the name `ИСТИНА` |
| `names-tables-held` | 4 | predicted before Excel took them: tables named with `?` and `\` |
| `names-ru-tables-held` (ru-RU) | 3 | predicted before Excel took them: tables renamed `ИСТИНА`, `ложь` and `TRUE` |
| `delete-held` | 13 | predicted before Excel took them: rows and columns deleted that take one table whole and part of another, or both whole |
| `headers-ru-held` (ru-RU) | 13 | predicted before Excel took them: numbers, dates and text typed into a header cell, the formulas over the column in English and in Russian |
| `paste-names-held` | 5 | predicted before Excel took them: a table's cells pasted onto a target that is no multiple of them, a row inserted where a calculated column's formula differs from its cells', defined names shaped like cells |

The predicted sets came out as predicted on 81 of 81 cells (`refs-held`), 424 of 433 (`ops-held`: two pastes over header cells, which Excel renames, and a conversion seen through a name were found there and fixed) and 69 of 69 (`locale`); `delete-held` on 5 of 13 cases, as Excel for Mac and Excel for Windows both refuse a deletion that takes one table whole and part of another, which the engine took (fixed); `headers-ru-held` on 3 of 13 cases: the column's name and the stored formulas on 10, where a header cell's own number format (`0.00`, the short date) names the column as it shows the number and the engine named it by General (fixed), and Excel read typed text as a date (a known difference); `toLocal` in ru-RU split a column named `1,5` at its comma and put it in brackets after `@` (fixed). The sets of names came out as predicted on 34 of 54 cases in Excel for Mac; [Defined names](names.md) has the misses and what was fixed. `paste-names-held` came out as predicted on 1 of 5 cases in both Excels: a copy of formulas onto a target that is no multiple of them took the target's other cells too, and a formula there in another table's calculated column got the copied table's name (fixed); formulas pasted into part of a calculated column fill the whole column in Excel (known); defined names past the sheet (`XFE1`, `A0`) are names (fixed); and a row inserted where a calculated column's formula differs from its cells' takes the column's formula, as predicted.

`test/tables-text.test.ts` shows formulas of a table's cell with `toLocal` and the table, in en-US and ru-RU, reads them back with `toInvariant` and checks the typed text with `validateLocalFormula`, each with the table, and both give the stored text again; the canonical text `validateFormula` gives with the table is the one `compute` stores for the cell.

The xlsx round trip (`run.ts xlsx roundtrip`) sends the 29 distinct books of the sets `refs`, `refs-held`, `ops` and `ops-held`: Excel opens each file `writeXlsx` writes and saves it, the files it saved read back to the books written, and the values Excel stored in them equal the engine's.

### The @ operator, unions, references across sheets, sheet names and the reference operators

From [The @ operator, unions, references across sheets, sheet names and the reference operators](references.md).

The set `areas` (1,137 cases) puts a union and a reference across sheets into every argument of every whitelisted function that takes a range, and adds the forms above; its case ids name the function and the argument (`fn:union:SUBTOTAL#2`). The set `areas-each` (52 cases) measures how `SUBTOTAL`, `AGGREGATE`, `FREQUENCY` and `SORTBY` take them: 40 of its 47 predicted answers came out as predicted, and the misses were all `SORTBY`, which takes neither. The set `intersect` (44 cases) measures the intersection and the range operator, 40 of its 41 predicted answers as predicted (the miss: `'D'!B1:B2:C3`, whose C3 is on the formula's sheet), and `sheet-cells` (13 cases) the sheets named like cells, all 11 predictions as predicted. The set `ref-deps` (11 cases) puts formulas between the ends of ranges, on the middle sheet of a reference across sheets and in the second area of a union, all 11 predictions as predicted. The spill sets `at` (8 books) and `at-held` measure `@`. The struct set `cube` (20 operations) measures structure operations, sheets inserted and moved included. The sets `at-held` (7 cells) and `cube-held` (14 cells) were predicted before Excel computed them, and all cells came out as predicted.

Sheet names were measured in the tables step of the harness (`bench/coverage/tables.ts`, `bench/windows/run.ts tables`): a book of sheets named as above, whose sheet `F` refers to each one's A1 as the engine stores it (in quotes), opened in Excel, its formula text read back (`Range.Formula`, and `Range.FormulaLocal`). The set `sheet-quotes-held` (54 names) came out as predicted on 49; its misses (`R0`, `R1048577` unquoted, `R2D2`, `C3PO`, `RC1C` quoted) showed that Excel reads an R1C1 reference at the start of a name, and the set `sheet-quotes-r1c1-held` (54 more), predicted after them, came out on 48: a `.` after the reference (`C1.x`, `R1.`), the column past the sheet (`R1C0`, `R1C16385`) and names that are A1 cells (`RCC1`, `xR1`) made the rule above. Excel 16.113 for Mac and Excel 16.0.20430 for Windows, in en-US and in ru-RU, wrote every name alike (`test/excel/tables/` and `test/excel/win/tables/`, `test/a1-excel.test.ts`), and the formula text of every other table set is now compared with the names quoted by the same rule.

### Typed entry

From [Typed entry](entry.md).

The rules were measured in Excel 16.113 for Mac and Excel 16.0.20430 for Windows, in en-US and ru-RU, by the step `entry` of the harness (`bench/coverage/entry.ts`, `bench/windows/entry.ts`). A case types its text into a cell of its own, in General or with a format set before, through the object model as Excel 365 takes typed text in the user's language (`formula2 local` on the Mac, `Formula2Local` on Windows); then the cell's shown text and format codes are read, and the book is saved and read back with `readXlsx` for what the cell holds. On the Mac a set may be typed on the keyboard instead (System Events, an ASCII keyboard layout): the window Excel shows for a formula it does not take ("We found a typo in your formula", then why it is not taken) is answered with No and OK, and Escape leaves the cell as it was. The answers are in `test/excel/entry/` and `test/excel/win/entry/`, and `test/entry-excel.test.ts` and `test/entry-win.test.ts` compare `parseEntry` with every case, on the day of the run.

| set | cases per locale | how | predicted before the run, en-US and ru-RU |
|---|---:|---|---|
| `typed` | 552 | object model, Mac and Windows | Mac 385 and 397, Windows 385 and 378 |
| `typed-keys` | 39 | keyboard, Mac | 30 and 30 |
| `typed-held` | 257 | object model, Mac and Windows | Mac 181 and 184, Windows 186 and 184 |
| `typed-keys-held` | 44 | keyboard, Mac | 36 and 36 |
| `typed-held-2` | 102 | object model, Mac and Windows | 85 and 84 on both |
| `typed-keys-held-2` | 29 | keyboard, Mac | 25 and 26 |
| `typed-keys-held-3` | 6 | keyboard, Mac | 3 and 4: `-5%`, `+5%` and `-0.5%` typed over a percent cell are formulas with the editor's second `%` (`-5%%`), where the engine wrote `-5%` |
| `typed-pct-held` | 47 | object model, Mac and Windows | 45 and 46 on both |
| `names-held` | 28 (en-US) | object model, Mac and Windows | 21 on both: cells' words as LET's names (`XFD1`, `XFE1`, `A0`) and the words past the sheet alone, refused here and taken by Excel, and an unknown name's case |
| `let-grid-held` | 14 (en-US) | object model, Mac and Windows | 10 on both: `IV1`, `A65536` and `IV65536` are LET's names too, so the edge predicted, Excel 2003's sheet, is none, and a LET name read after its LET |
| `let-cells-held` | 25 (en-US) | object model, Mac and Windows | 11 on both: every cell is LET's name but those that read as R1C1's row or column too |

Each set's predictions (`bench/coverage/entry/<set>.predicted.json`) were committed before Excel ran it, and the misses became rules. The engine now gives Excel for Windows's answer on 1,812 of the 1,983 cases of the object model's sets, and the Mac's keyboard answer on 232 of 236; the other cases are the known differences below, pinned in `test/entry-known.ts`.

Values are compared exactly, as the saved file holds every digit. `typed-pct-held` types percents with decimals where moving the point and dividing by 100 give different numbers (`1.1%`, `0.7%`, `1.255%`, `99.9%`, with a sign, parentheses, a leading `%`, groups and an exponent), fractions and plain decimals; each cell's value is also read through the object model (`value` on the Mac, `Value2` on Windows), which equals the number the file holds in every case. Both Excels stored every decimal percent as its decimal text reads; the miss was `$1.1%`, which Excel en-US keeps as text, the other the sign difference of the object model.

Excel for Windows and Excel for Mac answer alike, but where their own built-in codes part: the Mac's en-US short date is `m/d/yy` and its ru-RU short time `hh:mm`, so it keeps the codes `m/d/yyyy` (en-US) and `h:mm` (ru-RU) as formats of the book's own, where Windows takes them as the formats 14 and 20. The engine follows Windows.

The text a cell opens with was measured by the same step on the sets `edit-held` (85 texts) and `edit-numbers-held` (38 numbers around a length of 21 characters), which read each cell's `Formula` and `FormulaLocal` and, on the Mac, its formula bar itself: each cell selected in turn with Excel in front, the bar's text read through the window's accessibility (`Formula Bar. A,1 1/31/2025  6:00:00 PM`). The bar is needed: both Excels' object model gives a date's, a time's and a percent's number there (`45688`, `0.12`), and under other formats the text the bar shows, the apostrophe being the cell's prefix character. `entryText`'s predictions were committed before Excel ran each set: `edit-held` came out as predicted on 72 of 85 cells in en-US and 74 in ru-RU, the misses being numbers of 21 characters written whole, the two spaces before a time, the apostrophe Excel keeps and ru-RU's times, and with the rules fixed `edit-numbers-held` came out as predicted on all 38 numbers in both locales, in Excel for Mac's formula bar and Excel for Windows's `FormulaLocal` alike. `test/entry-text-excel.test.ts` compares `entryText` with the formula bar and with the object model under the other formats, and reads every text back with `parseEntry` to the value Excel saved.

The object model takes some text otherwise than typing, and the keyboard sets settle these: through `Formula2Local` a text begun by `+` or `-` that is no number is text (`-C1`, `+1+2`), a number goes into a percent cell as it is (`5` is 500%), and empty text makes a cell formatted as text hold empty text. Excel for Windows has `AutoPercentEntry` on, as the answers' `meta` records, but takes `Formula2Local` as the Mac does.

Text pasted from the clipboard and a text entered into a selection with Control+Return were measured by the step `ui` (`bench/coverage/ui.ts`; the sets `paste-held`, 25 pastes on both Excels, and `fill-held`, 18 entries on the Mac's keyboard), and the text a cell of an array opens with by the set `arrays-held` (the formula bar selected and once F2 began the editing, on the Mac; `Formula2Local` on both, which is the text F2 opens). The predictions, `setCellTexts` of the texts and `editText`, were committed before Excel ran each set, and the misses became the rules of [A paste and Control+Return](entry.md#a-paste-and-controlreturn) and of the cells of an old-kind array's area; Excel for Mac and Excel for Windows pasted alike in every case. `test/ui-excel.test.ts` compares `pasteText`, `fillCellText` (from the selection's first cell, a call with a stop alert taken back) and `editText` with every case: the engine gives Excel's answer to the 100 pastes (25 in each locale on each Excel), to 33 of the 36 entries into a selection and to the 16 cases of `arrays-held`; the other three are known differences below.

A cell edited in place was measured by the same step on the Mac's keyboard (the set `edit-held`, 20 cases in each locale): the cell selected, F2, what it opened with erased or only its digit changed, the text typed and Return, the formula bar read when F2 began the editing and once more before Return. The prediction, `pasteText` of each text (no percent editor at all), was committed before Excel ran it and came out as predicted on 14 cases in en-US and 13 in ru-RU: the misses were numbers begun by a sign (`-5` is -5%), to which the editor still adds `%`, and in ru-RU a column Excel widened. `test/ui-excel.test.ts` compares `setCellText` with `inPlace` with every case, and the engine gives Excel's answer to all 40.

Line breaks were measured by the step `rows` (`bench/coverage/rows.ts`, `docs/formats.md`, "Fitting rows"): text with a line break set through the object model (the sets `autofit` and `held`, both Excels), typed on the Mac's keyboard with Control+Option+Return, entered into a selection with Control+Return and pasted with Cmd+V (`keys`), and quotes in pasted text (`quotes`, both Excels, Windows through `Worksheet.Paste`; `quotes-ru`, the same pastes in ru-RU on Windows), each cell's content and wrapping read with its row's height. The keyboard, the object model and a paste set `wrap` alike, and the two Excels pasted every quote alike; the predictions, committed before each set ran, had the paste read no quotes, which the answers made the rule above.

Refusing a value a rule does not take (`refuseInvalid`) follows the buttons of Excel's alert of each style, which the harness does not measure: the alert is a window, and the object model writes a value without one. `test/session-refuse-invalid.test.ts` checks, in each kind of session, that a refused call leaves the book, the revision and the history as they were, and that a value the rules take is written as without the option.

### Typed entry: A formula left open

From [Typed entry, A formula left open](entry.md#a-formula-left-open).

The step `entry` typed each text on the Mac's keyboard and pressed Return (`keys` sets), in Excel 16.113 and 16.114 for Mac: Excel updated itself between the first set and the rest. The harness answers Excel's windows itself, No to a correction and OK after it, and keeps their texts, so a refused case holds the correction Excel offered. The Russian names were typed in the Russian layout (`input-source.swift com.apple.keylayout.Russian`), the other sets in ABC.

| set | cases | predicted before the run |
|---|---:|---|
| `close-held` | 34 in each locale | 29 (en-US) and 32 (ru-RU): two or more functions' brackets left open (`=SUM(MAX(1,2`) get a correction, not closed without a word; a sign before the function (`=-SUM(1`) is closed; an array constant left open in a function's brackets gets no correction |
| `close-ru-held` | 11 (ru-RU) | 8: as in `close-held` |
| `close-held-2` | 22 in each locale | 19 and 20: a space before the function (`= SUM(1`) is closed; an array constant after an operator (`=1+{1,2`) gets no correction; and a date format from a reference, a known difference |
| `close-ru-held-2` | 7 (ru-RU) | 7 |
| `close-held-3` | 15 in each locale | 11 and 10: a correction drops the spaces after `=` (`= (1+2` to `=(1+2)`); a formula begun by an array constant gets one for a later array left open (`={1,2}+{3`); and the cases below |
| `close-ru-held-3` | 5 (ru-RU) | 2: as in `close-held-3` |

The rule came from the answers of Point mode's sets, which pressed Return after the arrows (`point-caret-held-2` and `-3`: 62 formulas left open), and each set's misses corrected it; the third set's corrections were taken without a fourth. Five cases differ and are known (`test/entry-known.ts`): `=MAX(C1,C2` takes C2's date format in Excel, and the engine, without a book, none; `=NOW(` in ru-RU, where `NOW` is no Russian name; and three texts with two spaces (`=  SUM(1`), which macOS types as a period and a space (`NSAutomaticPeriodSubstitutionEnabled`), so they tell nothing of Excel. A case typed with Latin letters in the Russian layout (`=ABS(-(1)`, typed as `=ФИЫ(-(1)`) was taken out of `close-ru-held`.

### Typed entry: Copying plain text

From [Typed entry, Copying plain text](entry.md#copying-plain-text).

`bench/coverage/clipboard` commits the predictions before the copies run. The copy set has ten cases in each locale on each platform: grid, controls and quotes, one cell, an empty sheet, narrow numbers and dates, hidden rows and columns, filtering, merges and a spill. `copy range` and `pbpaste` on Mac and `Range.Copy` with `Get-Clipboard -Raw` on Windows read text without bringing Excel forward. The Windows `quotes-held` set copies three rectangles, replaces the rich clipboard with their plain text and pastes them, checking the sentinel below too. `test/clipboard-excel.test.ts` checks the stored text and pins the ambiguity; `test/session-copy-text.test.ts` checks `copyText` from the books against the same stored copies in all three session kinds and both execution modes.

### Typed entry: Point mode

From [Typed entry, Point mode](entry.md#point-mode).

The step `ui` of `bench/coverage` types on the Mac's keyboard in Excel 16 for Mac, en-US and ru-RU: a text typed over C3, keys pressed, the formula bar's description read after each (the active cell and its text), then Return, and the cell read back. Three sets, each predicted before Excel ran it:

- `point-held`, 40 cases in each locale, predicted by the sheet component's own model (`bench/coverage/ui/point-model.ts`): 21 of 40 as Excel did it in each. Excel differed where an arrow enters the text (after an operand the component moved the caret), after a leading `+` or `-`, a space after an operand and `=1.`, and in every key the component left to its field (Command and an arrow, Home, Page Down). `refAt` was written from these answers.
- `point-more-held`, 27 places, predicted by that first `refAt`: 39 of 54 as Excel did it. It entered every formula ending in an operand; Excel enters one only when it reads (`=A1)`, `=1e`, `=1,5` in en-US do nothing) and takes a reference after any point (`=A1.`). `refAt` was revised by them.
- `point-third-held`, 20 places no set had, predicted by the revised `refAt`: 19 of 20 in each locale. The one left is a reference to a sheet the book does not have (`=Sheet1!A`): Excel takes it for another workbook's and the arrow does nothing, which text alone cannot tell; `test/point-excel.test.ts` keeps it as a known difference.

Then the caret elsewhere than at the end, reached by F2, Home and arrows in Edit mode and F2 again, three sets more, each predicted before Excel ran it:

- `point-caret-held`, 24 cases in each locale: 27 of 48 as predicted. With the caret before the formula's first character Excel puts in its own `=` and the reference (20 cases); `=SUM(A1,` with the caret before the comma is entered in en-US. That the whole formula decides past the caret held.
- `point-caret-held-2`, 19 cases: 16 of 19 in each locale. A group's bracket left open is not closed (`=(1+2`, `=((1`), a function's is.
- `point-caret-held-3`, 13 cases: 9 of 13 in en-US, 11 of 13 in ru-RU. Excel does nothing with `=SUM(1)+MAX(2` and `=ABS(-(1)`, and in en-US with `=ROUND(1.5` and `=ROUND(A1*2` (where `ROUND` is known), which `refAt` enters. Measuring stopped at the third set with a correction. The sets of [A formula left open](#typed-entry-a-formula-left-open) gave the rule: the arrow enters a formula left open when Return would close it without a word, so `=SUM(1)+MAX(2` and `=ABS(-(1)` do nothing; the two of `ROUND`, too few arguments, stay known differences (`test/point-excel.test.ts`).

`test/point-excel.test.ts` compares `refAt` with every case that types a text and presses Right, 138 in each locale, the caret read from the keys pressed before the arrow. A case whose formula bar does not show the text it typed is not stored (another keyboard layout, or a key pressed by someone else at the Mac during the run, spoiled 11 cases of one run and 1 of another, which were run again; macOS turns two spaces typed into a point and a space, so no case types two).

Not measured: a mouse press on a cell or on a sheet's tab while pointing (the harness presses keys only), and Excel for Windows, whose keyboard is not reachable over SSH.
