# The coverage measurement and Excel's answers

Two things live here, outside the package. The coverage measurement: how many functions of `@formulajs/formulajs` 4.6.1 give Excel's answer, bare and through an adapter following the engine's runtime rules; its report is `docs/coverage-formulajs.md`. And the harness that computes Excel's answers for the engine's tests: the measurement's cases, the probes, the case sets, structure operations and spilling.

## Where the data are

Excel's answers are stored, so the tests run without Excel. Everything the tests read is in `test/excel/`:

- `coverage/`: the measurement. `cases.json` holds the cases with their formulas (written by `cases.ts`), `excel-en.json` and `excel-ru.json` Excel's columns. Tests: `test/excel.test.ts`, `test/excel-ru.test.ts`.
- `cases/<set>.json`: the case set `sets/<set>.ts` computed in Excel en-US; `cases/ru-RU/<set>.json` the same set in Excel ru-RU. Tests: `test/excel-sets.test.ts`, `test/excel-ru-probes.test.ts`.
- `probes.json`: the probes of `probes.ts`. Test: `test/excel-probes.test.ts`.
- `struct/<set>.json`: the structure operations of `struct/<set>.ts`. Test: `test/structure-excel.test.ts`.
- `spill/<set>.json`: the spill cases of `spill/<set>.ts`; `order-runs.json` keeps every answer seen for the pairs of arrays whose order Excel varies. Test: `test/spill-excel.test.ts`.
- `locale/ru-RU.json`: Excel's ru-RU formula text and international settings, written by `bench/locale/snapshot.ts`; `bench/locale/generate.ts` writes `src/locale/ru-RU.ts` from it. Test: `test/locale-excel.test.ts`.

What only the report reads is in `bench/coverage/data/`: `formulajs.json` (formulajs through the adapter and bare), `libreoffice.json` and `summary.json` (the verdicts by case and function, written by the `report` step).

Nobody edits these files by hand: each is written by the step that names it below, and a new run keeps the stored answers of the cases that did not change.

## Running it

```sh
npx tsx bench/coverage/run.ts all
```

The step `all` does in turn: the cases and the book, formulajs, LibreOffice, Excel en-US, Excel ru-RU, the report. Each step runs alone too (`cases`, `fj`, `lo`, `excel en`, `excel ru`, `report`) and writes its JSON: the cases and Excel's columns to `test/excel/coverage/`, formulajs, LibreOffice and the summary to `bench/coverage/data/`. A trial run on a few functions: `FCOV_ONLY=SUM,TEXT npx tsx bench/coverage/run.ts cases`, then the steps needed; everything then goes to `bench/coverage/out/` (not in git). The generator's test: `npx vitest run bench/coverage`.

The step `probes` computes in Excel en-US the probes outside the measurement (`probes.ts`: numbers as text, TEXT with General, long literals, criteria over whole columns of different sheets) and writes the cases with their answers to `test/excel/probes.json`. A set `sets/<name>.ts` is computed by the same step (`probes <name>`) in Excel en-US into `test/excel/cases/<name>.json` or, with a third argument `ru` (`probes <name> ru`), in Excel ru-RU into `test/excel/cases/ru-RU/<name>.json`; `test/excel-ru-probes.test.ts` checks the ru-RU answers in a book with `locale: 'ru-RU'`. The probes of `probes.ts` are computed in en-US only. How to write a set: `sets/README.md`.

The step `struct <set>` checks operations on a book's structure (inserting and deleting rows and columns, moving and copying ranges, renaming and deleting a sheet): `struct/<set>.ts` gives the books and one operation per case, Excel en-US applies the operation to the book (`excel-struct.applescript`, alerts off during the operation), and the sheets' formulas and the names' definitions are read back into `test/excel/struct/<set>.json`. An operation Excel refuses comes back from AppleScript as an error, with no window. `FSTRUCT_ONLY=id,id` runs only those cases.

The step `spill <set>` checks spilling arrays: `spill/<set>.ts` gives small books whose formulas are written as dynamic array formulas (`E2#` as `_xlfn.ANCHORARRAY(E2)`), Excel computes the book, and the values of a case's rectangles are read through the mirror sheets `WT`, `WN`, `WX`: three formulas per cell that never give an error (type code, number, text), plus the rectangles' own formulas. The answers go to `test/excel/spill/<set>.json`. `FSPILL_ONLY=id,id` runs only those cases. The set `held` was held out: its answers were predicted and committed before the run.

Needed: Node 24+, `soffice` on `PATH` (LibreOffice), Excel for Mac (the measurement used 16.113) and permission for the terminal to control Excel and System Events (Automation and Accessibility in the macOS settings).

## The code

- `excel-functions.json`: the function list of Microsoft's page "Excel functions (alphabetical)" (name, category, the version that brought it).
- `signatures.ts`: the signature table of the set's functions: argument kinds and typical calls, hand-written edge cases, `_xlfn.` prefixes, the array-size guard.
- `data.ts`: the data sheet of the synthetic book (sheet `D`).
- `cases.ts`: the case generator; `cases.test.ts` checks the size guard over every case.
- `formulajs.ts`, `fj-worker.ts`: formulajs called through the adapter and bare in a child process; a case that hangs for over 5 s is killed and marked.
- `book.ts`: the xlsx book (ExcelJS, no cached values), its case formulas written as dynamic array formulas.
- `libreoffice.ts`: recalculation by `soffice --headless` with an en-US profile, reading the XML.
- `excel.ts`, `excel-read.applescript`: Excel through `osascript`.
- `sets/fuzz.ts`, `sets/fuzz/`: the random check, formulas over the whitelist from a seed (argument kinds and counts frozen in `sets/fuzz/spec.json`, each formula with its own stream of random numbers); `sets/fuzz-rules.ts`: the Excel rules it found, as cases.
- `compare.ts`, `report.ts`: the comparison and the report's tables. The tables stand in the report between the marks `<!-- auto:name -->` and are rebuilt by the `report` step; the text outside the marks is written by hand.

## Rules for working with Excel

Breaking them cost two modal dialogs and three Excel crashes while the measurement was prepared.

- A book goes into Excel's sandbox `~/Library/Containers/com.microsoft.Excel/Data/tmp/fcov-<pid>/`: Excel opens it from there without asking for access. The directory is removed after the part.
- Opening: the HFS path is computed before the `tell` block (`set hfs to (POSIX file p) as text`), then `open workbook workbook file name hfs`. With `POSIX file p` inside `tell` Excel answers with a parameter error (-50).
- Reading: `value of range` is read only from columns without errors. Error values drop out of a range's `value` list (positions shift), and reading a two-dimensional range with errors crashed Excel (error -609). So next to a case's formula (column A) the book has three wrapper columns that never give an error: B the type code (0 number, -1 TRUE, -2 FALSE, -3 text, above 0 the `ERROR.TYPE` number), C the number, D the text. The three columns are read in three calls and handed over through `NSJSONSerialization` (ASObjC): numbers arrive with full precision, while printing reals through `osascript` loses digits.
- Calculation: `calculate full`, then `close wb saving no`.
- Functions newer than Excel 2007 are written with a prefix (`_xlfn.XLOOKUP`, `_xlfn._xlws.SORT`), listed in `signatures.ts`. A missing prefix shows as `#NAME?` on every case of the function; the measurement has none. LET's names are written `_xlpm.x` where declared and used in their scope (`sets/let/xlpm.ts`, also by the spill harness); a word without the prefix is a defined name.
- Every case formula, of the measurement, the probes and the sets, is written as a dynamic array formula, as Excel 365 saves it: `cm="1"` on the cell, `<f t="array" ref="A1">` and the XLDAPR metadata (`fDynamic="1"`) in `xl/metadata.xml` (`book.ts`). Such a formula is computed without implicit intersection and an array answer spills; a case's neighbours are taken and a stub stands under the last case, so an answer larger than one cell is always `#SPILL!`. A `SUM(...)` wrapper does not help a formula of the old kind: the intersection happens inside it (`SUM(D!B1:B8*2)` as an old formula in row 2 gives 16). The harness writes no formulas of the old kind since 7 October 2026. The measurement's Excel columns were computed again as dynamic formulas on 6 October (the former columns `excel-en-legacy.json` and `excel-ru-legacy.json` are in git at commit `85522cf`, under `docs/coverage-formulajs/`), and so were the sets `criteria`, `findate`, `stat`, `subtotal` and the probes (one answer of 5,824 changed). The LibreOffice column and the locale snapshot were made from books of the old kind.
- The size of an array taken from an argument (MUNIT, EXPAND, TAKE, DROP): only small number literals (at most 100) in the size position, no references, text or dates. `MUNIT(45688)` and `MUNIT("1.5")` in ru-RU (the text became the date 46143) ran Excel out of resources and raised a modal dialog. For the other size and count arguments (REPT, the order of BESSEL, periods, the days of WORKDAY) the "large" value is 1000, and a date and `"1.5"` are not put in.
- Excel computes in parts of up to 400 cases, by whole functions. Before and after each part System Events checks whether Excel has a window that is not a book window (subrole other than `AXStandardWindow`). If so, the run stops, the window's text and the part's number go to `meta.stopped` of the column's file, and nothing is clicked. If System Events answers the query with error -1719 (a book's window closed during it), the query is repeated, up to three times (`excelDialogs` in `excel.ts`, shared by all steps). The parts done stay in the file, and a new run goes on from the first part not done (`FCOV_FRESH=1` starts over).
- The interface language and the calculation locale come from launch arguments; Excel's settings (`defaults write`) are not changed. For en-US Excel is started with `open -a "Microsoft Excel" --args -AppleLanguages '(en)' -AppleLocale en_US`, for ru-RU without arguments (the system's Russian language and locale). It is restarted only with no books open: `quit`, waiting for it to exit, `open -a`. AppleEvents go only to a running Excel: `tell application "Microsoft Excel"` starts a stopped Excel without arguments. The empty new book Excel opens when started ("Book1" with the English interface, "Книга1" with the Russian one; no file, no edits) is closed without saving before the first part, so neither the book count nor the window check trips on it. The step `all` ends with ru-RU, so Excel is left running without arguments.
- The interface language changes computation, not only display, so en-US needs the English interface as well as the locale. With the Russian interface the text `"TRUE"` is no boolean (`IF("TRUE",1,2)` is `#VALUE!`, while `IF("ИСТИНА",1,2)` is 1), a boolean in text comes out as `ИСТИНА` (`TRUE&""`), and `CODE("Ж")` takes the Mac Cyrillic encoding (134 against 95). The first en-US column was computed with the en_US locale and the Russian interface and differed from the right one in 32 cases; it is in git (`docs/coverage-formulajs/excel-en-ru-ui.json` at commit `85522cf`), and the report keeps its comparison with the current one as fixed text. Excel shows error texts in the interface language; the measurement does not read them, the error code comes from `ERROR.TYPE`.
- Excel's answer does not always repeat. `VLOOKUP(5,D!L1:N6,D!E2:E4,FALSE)` (a column number from a range holding the formula `1/0`) gave `#REF!` or `#VALUE!` in one book depending on the row and the run, which looks like the recalculation order. `DAY(WEEKNUM(40977,1E+3)<>2.5)` gave 1 twice in the random check's book, and `#NUM!` once and in a small book. Check a difference on a case that touches such places by a second run and a separate book first.
- If Excel shows a dialog, hangs or asks for a manual action: stop and report it; do not close or get around the dialog.
- Moving (cut and paste) a cell with a dynamic array formula to another sheet crashed Excel 16.113 twice (6 October, the struct set `spill`); Excel restarted by itself without arguments, that is in ru-RU, with a book recovery window, and the run stopped on that window. Do not add such a case.
- AppleScript variable names inside `tell application "Microsoft Excel"` must not match terms of Excel's dictionary: a variable `a1` there reads as the reference style constant `A1`, and `worksheet a1` gives "object does not exist". The same goes for `tab`: inside the block use `character id 9` instead.
