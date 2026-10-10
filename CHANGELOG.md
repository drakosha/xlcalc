# Changelog

## 0.13.0 (2026-10-10)

Internal release; not published to npm. Messages for people in the language of the book.

Changed:

- Messages for people come in the language of the book's locale: why a typed formula does not read (`parseEntry`, `toInvariant`, `toLocal` and `validateLocalFormula`, by the locale of the call), why a structure operation is refused (`restructure` and a session's `restructure` in every kind of session, by the book's `locale`), and the text of `explain`. In ru-RU they are Russian, with the formula's functions, separators, booleans and errors as ru-RU writes them: `Ожидалась ), а формула закончилась`, `Функция ОКРУГЛ принимает 2 аргумента, а указано 3`, `Нельзя изменить часть массива План!C1 (C1:C3)`, and `explain` shows `'План'!B1 = 420  ← ОКРУГЛ(СУММ(A1:A2)*1,2;0)` over `'План'!A1:A2 (2 ячейки)`; its nodes keep the stored text. Codes, positions and the English messages are as before.

Added:

- A locale file has `messages`, the catalog of the 75 messages for people, each a function of its parameters (plural forms and the locale's words need more than a template); a locale without its own text of a message takes en-US's (`docs/locale.md`). The ru-RU catalog adds 2.9 KB gzipped to a bundle.

Still in English: `validateFormula`, which has no locale, the reasons of cell errors in `issues`, the checks of names and `InputError`. In `xlcalc/unstable/internal`, `FunctionSpec.counts` is an object (`{ min, max, parity, or }`) instead of text.

## 0.12.0 (2026-10-10)

Internal release; not published to npm. What a window over a sheet needs, without the whole book.

Added:

- Contract: `Session.extent(sheet)`, where a sheet's cells end: the last row and column that an input cell (a value, a formula, or a style, a format, meta, a note or a link alone), a spill range, a merged area or a table reaches. It reads the engine's columns, spills and cells without a value instead of every cell: under 0.1 ms on the synthetic book of 34,000 rows from a thread, where finding the formulas in `book()` and reading their spills took about 700 ms.
- Contract: `Session.layout()`, the book without its cells: styles, names, settings and every field of each sheet but `cells`. It costs what the layout holds: 0.1 ms on the same book from a thread, where `book()` takes 460 ms.
- Contract: `Session.cells(sheet, ranges)` takes several rectangles in one call, one round trip in a pool; a cell in two of them is listed once.
- Contract: `Session.patchCells(patches)` changes cells' `style`, `format`, `meta`, `note` and `link` (null removes one) without their values or formulas: an empty cell, a spilled cell and a cell of an old-kind array formula's area alike, atomically and without recomputing; with the display on it returns the cells whose shown text or colour changed. A pool journals it with the other edits. Styling 2,000 cells from a thread takes 9 ms, against 15 ms by reading a window and writing it back through `setCells`.
- Contract: `CellView.array`, an old-kind array formula's area on its top-left cell.

## 0.11.0 (2026-10-10)

Internal release; not published to npm. Conditional formats, kept and computed.

Added:

- Contract: `SheetInput.conditionalFormats`, a sheet's conditional formats in priority order: cell value rules, formulas, text, blanks and errors, time periods, top and bottom, above and below average, duplicates and unique values, colour scales, data bars (Excel 2010's fields too) and icon sets (custom icons too), each over several ranges, with Stop If True and its style (font, fill, border, number format) in the rule. Layout edit `{ sheet, conditionalFormats }`.
- Contract: `CellView.look`, what the sheet's rules give a cell by the values the book holds now: the merged style of the rules that hold, a colour scale's fill, a data bar's place and colour, an icon, and whether the value is hidden. `Session.cells` gives it, a blank cell with a look too; `compute` and `result()` do not, and a look is no part of the values hash. A host reads its window again after an edit: `changes` do not list cells whose look alone changed. Statistics of the rules are cached until an edit touches their cells; on 34,000 rows with a formula rule, a colour scale and a data bar a window of 20,000 cells takes 55 to 63 ms the first time, 24 to 28 after.
- `readXlsx` and `writeXlsx` read and write conditional formats with their differential formats and Excel 2010's extensions without loss. Structure operations move them as Excel does (rows inserted right below a range join it, a move keeps the target's rules, a copy adds a rule first in order). Measured on the Mac and on Windows, predictions committed before each run: after the rules the misses taught, the engine gives Excel for Mac's look in 4,015 of 4,017 cells; Excel for Windows's in all but the cells where it shows `#` by its own text widths; rules Excel made through its object model read as Excel shows them (203 of 204 cells); `docs/conditional-formats.md`.

Fixed:

- A data validation reader took a 2007 data bar's extension for the sheet's and lost the sheet's rules of data validation.

Known differences: a data bar's shortest and longest length of 5 and 95 percent is not kept by Excel; the length of a bar and the number of an icon follow Excel's documented rules but were not measured (Excel's object model does not give them); borders, underline and strike of a rule's style are kept and written, not measured.

## 0.10.0 (2026-10-10)

Internal release; not published to npm. Data validation and outlines of rows and columns.

Added:

- Contract: `SheetInput.validations`, the sheet's rules of data validation as Excel keeps them (a range, the type: any value, whole number, decimal, date, time, text length, list or custom formula; its operator and formulas, blanks allowed, the in-cell dropdown, the input message and the error alert with its style). Its formulas are written as at the rule's first cell. Layout edit `{ sheet, validations }`.
- Contract: `outlineLevel` (1 to 7) and `collapsed` on rows and columns, as Excel's groups, and `SheetInput.outline` (`summaryBelow`, `summaryRight`); a collapsed group is its hidden rows, which `SUBTOTAL` 101 to 111 already leave out. Layout edits for both.
- Contract: `Session.validation(sheet, cell)` gives a cell's rule, its formulas as at that cell and, for a list, its items computed on the current book with their shown text (a literal list, a range, a name, a spill); `Session.validateEntry(sheet, cell, input)` tells whether a value meets the rule as Excel checks it (a literal list in its case, a list of cells in any case, numbers to 15 digits, custom formulas seeing the value). Both in synchronous, same-thread and pool sessions.
- `readXlsx` and `writeXlsx` read and write data validation (Excel 2010's extension too, where a rule refers to another sheet) and outlines. Structure operations move rules as Excel does: inserted rows right after a rule's range take it, a moved range keeps the target's rule where it brings none, a copy is Paste All. Inserted rows take the outline level of the row above. Measured on the Mac and on Windows (`struct-held` 26 of 31 predictions, `struct-held-2` 14 of 15; the values checked against Excel's `Validation.Value`); `docs/validation.md`.

Known differences: a list item holding a comma (Excel 365's `x12ac:list`) is read split at the comma, with an issue; texts longer than Excel's limits are cut when written.

## 0.9.0 (2026-10-10)

Internal release; not published to npm. Notes, hyperlinks and frozen panes in the model, and text typed into a cell read as Excel reads it.

Added:

- Contract: `CellInput.note` (author and text) and `CellInput.link` (`url`, `location`, `tooltip`), which go with their cell as a style does; `SheetInput.frozen` (`rows`, `columns`) and its layout edit `{ sheet, frozen }`; `Session.cells` gives a cell's `note` and `link`. None of them takes part in computation or in the values hash; `inputHash` takes them when a book has them, and a book without them hashes as before.
- `readXlsx` and `writeXlsx` read and write notes (the comments of the old kind, shown on hover), hyperlinks and frozen panes. Excel's threaded comments are read as notes: the first comment's author, the replies in the text, each as `Author: text`; their dates and mentions are reported as lost. A split without frozen panes is reported.
- Structure operations move notes and hyperlinks with their cells, copies copy them, and frozen panes move with rows and columns inserted or deleted before their edge, as Excel does (measured: `struct-held` 24 of 26 predictions, `frozen-held` 9 of 9, on the Mac and on Windows); deleting every frozen row leaves one frozen.
- Contract: `parseEntry(text, { locale, cell, now, dateSystem })` reads text typed into a cell as Excel reads it in the book's locale: numbers with the locale's separators, percents, currency, dates and times, booleans, errors, `'` for text, formulas through `toInvariant` (only a formula that does not parse stops the entry, at its position). It gives the `CellInput` with the format Excel gives the cell (`15.03.2026` is format 14, `5%` format 9, `1 234,5` format 4), keeping the cell's style and `meta`, and a cell's own format where Excel keeps it. Measured on 1,822 cases through Excel's object model on the Mac and on Windows, en-US and ru-RU, and on 224 typed on the Mac's keyboard; `docs/entry.md`.

Changed:

- Date text of three numbers takes a year of one digit, as Excel reads it: `VALUE("1-2-3")` is 2 January 2003 (37623). It was `#VALUE!`.

## 0.8.1 (2026-10-10)

Internal release; not published to npm. Running aggregates without rereading, a smaller graph of range dependencies, and tests that pass on every supported Node under Linux.

Changed:

- `AVERAGE`, `COUNTA` and `COUNTIF` of one range share prefixes as `SUM`, `COUNT`, `MIN` and `MAX` do, read cell by cell without making the range whole: a running `AVERAGE` over 34,000 rows opens in a thread of 384 MB in 0.28 s instead of 6.0 s, a running `COUNTA` in 0.28 s instead of 1.7 s. `AVERAGE` keeps its compensated sum and its exact fallback, so values and hashes are the same; one such range larger than `limits.cells` is computed rather than refused.
- Formulas over overlapping ranges share interval vertices in the dependency graph instead of an edge from every formula they read: 1,024 running ranges keep under 32,768 edges where they had 524,800. The order of evaluation and every value are the same.
- A large batch of formulas drops the running aggregates over them once, not once per formula.

Fixed:

- The tests passed only on the Mac: on Linux, and on Node 20 and 22, the round trip and filter files of the xlsx subpath failed their check against the files Excel saw (the platform's deflate writes other bytes, so a file is now identified by its parts), two known differences of `RATE` failed their pins by their last digits (a pin is now compared within Excel's tolerance), and Node 20's `TextDecoder` read windows-1252 otherwise than the standard. They pass on Node 20.19, 22, 24 and 26 under Linux. A tag's CI pipeline only packs the package, its commit having been tested when it was pushed.

## 0.8.0 (2026-10-10)

Internal release; not published to npm. A locale is one file, and lookups and statistics measured at their edges in Excel.

Added:

- Contract: `LOCALES`, the package's locales by tag (`LOCALES.get('ru-RU') === ruRU`). The contract's comments name the locales by it instead of listing their tags.
- A locale is one file of `src/locales/` and a line of its list (`src/locales/index.ts`): its formula text, which `bench/locale/generate.ts <tag>` writes between marks, its regional data and the widths of its script's letters. The tests against Excel and the harnesses of the Mac and of Windows take any locale of the list by its tag; a locale without Excel's answers skips its tests against Excel. `docs/locale.md` describes adding one.
- Case sets measured in Excel for Mac and Windows with the engine's answers committed before: `stat-edges-held` (111 cases), `lookup-held` (19), `lookup-own-held` (9), and `lookup-rules` (14), measured after the misses.
- A seeded test of sessions against full computations aimed at the range caches, in both execution modes.

Changed:

- The order of day, month and year in date text, the side the currency is read on, the published table of the code page and the widths of Cyrillic letters are data of a locale; names are compared by one locale-independent key. ru-RU and en-US compute as before.
- `SORTBY` with an array of sort orders answers position by position, as Excel does: the orders are expanded as an operator's operands, each position gives the sort's error or its first value, the next ones a blank (`{2;-1}` gives `{#VALUE!;"thirteen"}`).
- `LOOKUP` reads on past the end of a result that a spill, a function, parentheses or an operator answers, as it did past a reference written in the formula: a column down, a row or a cell to the right. It was `#REF!`.

Known differences:

- Variance, standard deviation, covariance, correlation and regression on data where doubles lose the scale (a small spread around 10^12, values near 1E-200, three values of 1E+308): the engine gives the exact answer, Excel for Mac and Windows lose digits, up to all of them, or answer `#NUM!` and `#DIV/0!` (40 cases).
- `LINEST` over data with a large offset keeps fewer digits than Excel (2 against 14 over 10,000 rows near 10^12), while `SLOPE` is exact there.
- `LOOKUP` in a cell it may read past its result's end is `#CYCLE!`, where Excel computes the cell it read: the engine counts every cell the lookup may read among the formula's references.

## 0.7.0 (2026-10-10)

Internal release; not published to npm. Names checked in a book's locale and without the book, a budget for the range caches, and what Excel had not measured.

Added:

- Contract: `validateName(name, { locale })` checks a name by the rules of a book of that locale: a ru-RU book refuses `ИСТИНА` and `ЛОЖЬ` in any case. Without a locale it checks en-US's rules, as before.
- Contract: `validateNames(names, { sheets, tables?, externalFunctions?, locale? })` gives the name issues a book with these sheets, tables and locale would give, by the book's own check, without the book; `engine` has it too.
- Contract: `NameIssue.dropped` is true for a name the book leaves out (one Excel would not take, a duplicate, an unknown scope sheet, a formula that does not parse or a reference without its sheet) and false for one it keeps that gives an error where it is used (an unknown sheet, table or function in its formula).
- Case sets measured in Excel, with the engine's answers committed before each run: `areas-each`, `intersect`, `sheet-cells`, `single-values`, `findate-edges` and `ref-deps` in Excel for Mac ru-RU (258 of 258 held); `tables/delete-held` (13 cases) and `tables/ru-RU/headers-ru-held` (13) in Excel for Mac and Windows; `struct/zero-held` (8, all held) in both. The Windows harness reads the built-in number formats (`xlsx formats`).

Changed:

- A session's range caches keep a range's values and a criteria index from the range's second read, and the caches of all books in one heap share a budget: a quarter of `maxHeapMb` in a pool's thread (96 MiB for 384 MB), 512 MiB elsewhere. Past it the ranges read least recently are dropped between formulas. A book of running `AVERAGE`, `COUNTA` or `COUNTIF` over growing ranges kept every range: over 10,000 rows it ran a thread of 384 MB out of memory, and over 34,000 rows a synchronous computation brought down a Node process with a heap of 4 GB. Now such books open and are edited (10,000 rows of `AVERAGE` in that pool: 0.56 s to open, 0.34 s an edit). Values and hashes are the same, and the synthetic book computes as fast as before.
- Rows or columns deleted across two tables are refused unless both tables go whole, as Excel for Mac and Windows refuse them. They were refused only when both tables kept some of their lines.
- A number typed into a table's header cell names its column as the cell's own number format shows it (`1,50` by `0.00`, `31.01.2025` by format 14 in ru-RU), as Excel does; it was named as General shows it.

Fixed:

- `toLocal` in ru-RU keeps a column named with a comma whole (`Sales[1,5]`; it turned the comma into `;` and changed the reference) and writes `[@1,5]`, as Excel ru-RU does.

Known differences: Excel for Windows shows the built-in formats 50 to 58 (East Asian dates) as its short date, where the reader reads them as `m/d/yy`; Excel reads the text `2025-01-31` typed into a header cell as a date, where `setCells` keeps the text.

## 0.6.5 (2026-10-09)

Internal release; not published to npm. Sums of one range in a thread pool, and how to add a locale.

Fixed:

- `SUM`, `COUNT`, `MIN` and `MAX` of one range read it cell by cell under a finite `limits.cells` too, as they do without limits. A pool with `maxHeapMb` always has such a limit, so there these sums made an array of every range and a session kept each one: a running total over 10,000 rows (`SUM($A$1:A<row>)`) ran a thread of 384 MB out of memory, and one over 5,000 rows held 103 MiB where it now holds 8. One range larger than `cells` under these functions is computed rather than refused; values and hashes are the same.

Changed:

- ru-RU's code page of `CHAR` and `CODE` and the word of new table columns' names (`Столбец`) are data of the locale, beside its other regional data. Nothing computes differently.

Added:

- `docs/locale.md` describes adding a locale: its two layers, where it is registered, where its data come from, how it is checked against Excel, and what still needs code (the order of dates, a currency before the amount, the widths of characters, the case of names, the harnesses).

## 0.6.4 (2026-10-09)

Internal release; not published to npm. The rules of names measured in Excel.

Changed:

- A defined name in a formula's text is written as the name is defined, as Excel shows it: `ставка*2` is `Ставка*2` in `result()` and changes, with a name `Ставка` (contract: the doc of `CellResult.formula`).
- A name may hold `?` and `\` after its first character, as Excel takes them (`Став?ка`, `Rate?`, `Став\ка`; not `?Ставка`), defined names and tables alike. `?` can no longer be a separator of a locale.
- In a ru-RU book `ИСТИНА` and `ЛОЖЬ` are no names, in any case: Excel ru-RU does not open a file with such a name. The name is left out with an issue, and `writeXlsx` does not write it for a ru-RU book. In en-US they are names, as there.
- `toLocal` in ru-RU writes a name `ИСТИНА` or `ЛОЖЬ` of an en-US book in quotes (`'ИСТИНА'+1`), as Excel ru-RU shows it, and `toInvariant` reads it back as the name.
- Contract: the docs of `NameInput.name` and `validateName` give these rules; `validateName` checks those of en-US.

Added:

- Six case sets of names and table names in en-US and ru-RU (`tables/names-held`, `names-more-held`, `names-tables-held` and their ru-RU sets), measured in Excel for Mac 16.113 and Windows 16.0.20430, the engine's answers committed before: 34 (Mac) and 35 (Windows) of 54 predictions held. The harness defines names through Excel's object model, so a name Excel refuses is an answer of the script rather than a window to close.
- Known differences: a table renamed to a name Excel does not take is renamed by Excel's object model (`_1Ставка`, `СТАВКА_1`) and refused here; `TRUE` and `FALSE` as names in ru-RU, which Excel ru-RU takes and stored formulas read as booleans; tables named `ИСТИНА`, `ложь` or `TRUE` in ru-RU, where Excel for Mac and Windows part (the engine refuses them, as Windows refuses names).

## 0.6.3 (2026-10-09)

Internal release; not published to npm. Three known differences from Excel in lookups closed, measured in Excel for Mac and Windows.

Changed:

- `SORTBY` with an array or a range of several cells as a sort order answers an array of the orders' size, as Excel does: the first value of the sort by the orders' first elements, then blanks where every order is 1 or -1, `#VALUE!` where an order is another number or has no value at that position, an order's error where it holds one (`SORTBY(A1:A6,B1:B6,{1;-1})` is `{"one";}`). It was `#VALUE!`.
- `LOOKUP` reads on past the end of a result range shorter than its lookup range, as Excel does: a column or a row along itself, one cell as a row (`LOOKUP(4.19,L1:L6,B2)` reads D2), past the sheet's edge a blank; those cells are among the formula's references, so a session recomputes it when they change. It was `#REF!`; a result range that a function returns (`IF(TRUE,B2:B3)`) still is.
- `IFS` whose condition is text that is no boolean answers its `#VALUE!` as an array of one error, as Excel does: `ROWS` and `COLUMNS` of it are 1, `TYPE` 64; an error given as a condition or a value is the answer as it is. `fromLegacy` writes `TYPE` of such an array without `@`.

Added:

- The case sets `known-held` (37 cases) and `legacy/ifs-text` (8), with the engine's answers committed before Excel computed them: 33 of 45 predictions held, and the engine was fixed for the 12 that did not; Excel for Mac and Windows agree on all 45.
- The Windows harness takes `XLWIN_CODEPAGE=any` to run a set that reads no code page without restarting the machine for one.

## 0.6.2 (2026-10-09)

Internal release; not published to npm. Caches of ranges, aggregates and exact lookups, stable statistical moments, and a fix of cached ranges after a sheet grew.

Fixed:

- A session read a cached range with the shape of the sheet at the time it was cached: after a write outside the range grew the sheet (`Data!E1000` beside a formula over `Data!A1:D256`), a formula over the range computed again gave a wrong value (1 where 3 was right), and a whole-column spill freed of an obstruction gave `#N/A` in sums over it. A cached range now keeps the shape it was read with.

Changed:

- Variance, standard deviation, covariance, correlation, regression and the database statistics use compensated sums and a corrected two-pass calculation, and exact binary sums where doubles lose the scale (a small spread around a large mean, values near 1e-200 or near the largest double, a small intercept or residual after cancellation). Ordinary data give the same values; such data give the mathematically exact answer, which can differ from Excel's in the last digits. Excel has not computed these cases; the stored Excel answers all still match ([docs/statistics.md](docs/statistics.md)).
- Faster sessions over large ranges: cached ranges and criteria indexes are found by sheet and column when an edit drops them; `SUM`, `COUNT`, `MIN` and `MAX` of one range keep their answers until the range changes; a repeated exact lookup in an unchanged range uses an index built after enough scans ([docs/cache.md](docs/cache.md)). On the synthetic book an edit of a criterion cell takes 0.77 to 0.80 ms at p95 instead of 0.98 to 1.1 (3,400 rows); a cold compute and `hash()` take what they took, and every hash is the same.

## 0.6.1 (2026-10-09)

Internal release; not published to npm. Defined names in any script.

Fixed:

- A defined name takes letters of any script, as in Excel: `Ставка`, `НДС_2025`, `АБ12` (Cyrillic letters are no cell address) are names in `BookInput.names`, `validateName`, `setNames` and formulas, in any case (`ставка` is the name `Ставка`). They answered `a name must start with a latin letter`. Defined names and tables now share one check of Excel's rules: a letter of any script, `_` or `\` first, then letters, digits, `_` and `.`, at most 255 characters, not a cell address in Latin letters, not R1C1, not TRUE or FALSE.

Changed:

- `R` and `C` alone are no defined names (Excel reads them as R1C1 references), and a defined name may start with `\`, as a table name already did.
- Contract: the doc of `NameInput.name` gives these rules.

Added:

- The case sets `tables/names-held` (en-US) and `tables/names-ru-held` (ru-RU): names in Cyrillic in formulas, in any case, beside tables, next to addresses and R1C1, with the engine's answers committed as the prediction; Excel has not computed them yet.

## 0.6.0 (2026-10-09)

Internal release; not published to npm. A spill has a limit of its own, so a thread pool bounds spills by its heap strictly and arrays loosely; a call's limits are taken field by field.

Added:

- Contract: `CalcLimits.spill`, the cells one spill may store (the cells past a large array's last value are not stored, so `A:A*2` over the data rows is not refused); without it `cells` bounds spills too, as before. `CalcError.limit` can be `spill`. `cells` is the values of one range read whole or one array.

Changed:

- A pool with `maxHeapMb` takes, where its `limits` give none, `cells` of a value per 64 bytes of heap (6,291,456 values for 384 MB; 0.5.0 took a value per KB) and `spill` of the smaller of `cells` and a value per KB (393,216 for 384 MB), so `SUM(SEQUENCE(500000))` and `TOCOL(A:A)` compute in such a pool again while a spill stays bounded, also where the host raises `cells`. Measured on Node 24 and 26 with heaps of 256 to 1,024 MB: an array brought the process down from 30 bytes of heap per value, a spill read by `compute`, `result()`, `hash()` or `setCells` from 286; in 1,380 runs at the defaults and at twice them the process never went down.
- Contract: a calculator's `limits` (`createInlineCalc`, `createCalcPool`) hold for every call, and a call's own `limits` take their place field by field. A call that gave `limits` replaced them whole: one that gave `formulas` alone lost the pool's `cells` and `spill` taken from its heap, and an array the heap could not hold ran the thread out of memory.

## 0.5.0 (2026-10-09)

Internal release; not published to npm. A thread pool's default limit on one array measured, and in the same thread a result that gives way to the event loop and a structure operation that stops at its deadline.

Changed:

- A pool with `maxHeapMb` and without `limits.cells` takes `cells` of `maxHeapMb · 2^20 / 1024`, a value per KB of heap (262,144 values for 256 MB; it was a 32nd of the heap's bytes, 8 million values). Measured on Node 24 and 26 with heaps of 256 to 1024 MB: a large spill read by `compute`, `result()`, `hash()` or `setCells` brought the whole process down from 300 bytes of heap per value (1.57 million values in a heap of 384 MB), which the old default let through; other arrays from 30 bytes. Formulas over whole columns are computed on their occupied rows and stay within the limit; an array over it, such as `TOCOL(A:A)` or `SEQUENCE(500000)` in a pool of 384 MB, is `CalcError` with `limit`, and `limits.cells` given to the pool allows more.
- Contract: the doc of `CalcPoolOptions.limits` names the new default and why.
- In the same thread (`createInlineCalc`) `result()`, `hash()` and `cells()` give way to the event loop between portions, as a computation does: on the synthetic book of 34,000 rows the event loop waits about 50 ms at most, where `result()` held it for its whole second.
- A structure operation checks the deadline and the signal as it rewrites and moves cells: a session's `restructure` with `timeoutMs` stops with `soft_timeout` and leaves the session as it was (47 ms with `timeoutMs: 20` on the synthetic book of 34,000 rows, where it ran 1.9 s). In the same thread a signal that fires while it runs is seen when it ends. Structure operations walk a sheet's cells by their keys, which makes them faster on large sheets.

Added:

- The thread pool is checked in headless Firefox and WebKit besides Chromium, with Node's hashes in each; WebKit stands for Safari.

## 0.4.4 (2026-10-09)

Internal release; not published to npm. Fixes from a second review (calculated columns under copies, moves and sheet operations, long lists of dependents, the thread pool's fallback to objects), text with a line break read as a number, and a faster `Session.hash()`.

Fixed:

- The formula of a table's calculated column, the one a row inserted into the table gets, is copied and moved with the table as the formulas of its cells are, and follows a renamed, moved or deleted sheet and a moved range it refers to. It stayed as it was: a row inserted after `copyRange` or `moveRange` of a table got a formula of the table's old place (`A3*2` where its row is 7), and one inserted after `renameSheet` a reference to the old name (`#REF!`).
- A volatile cell read by more formulas than a JavaScript call takes arguments (139,999 in Node) threw `RangeError` in `compute` and in a session; `FREQUENCY` over more than about 120,000 numbers answered `#VALUE!`. Both take any number now.
- Text with a line break where a number is read (`A1*1` over a cell holding `"x\ny"`, in en-US and ru-RU) threw `TypeError` and stopped the whole `compute` or session call; it is no number, `#VALUE!`, as other text.
- The thread pool: a large book that JSON cannot hold as it is (`-0`, a date in `meta`) took the fields of the book and its sheets (a sheet's name, the names) as they were when its cells had been written, after the call had returned; a sheet renamed meanwhile reached the thread and the session's snapshot (since 0.4.2). They are taken when the call is made, as for a book sent as text.

Changed:

- `Session.hash()` and `result().hash` write the canonical lines to SHA-256 as the cells are read, one canonicalization for both, without the canonical text or a map of the values: on the synthetic book `hash()` takes 30 to 32 ms instead of 36 to 38 at 3,400 rows and 340 ms instead of 443 to 474 at 34,000 rows, a cold compute is no slower (2 to 9 % faster at 34,000 rows), and every hash is the same.

## 0.4.3 (2026-10-09)

Internal release; not published to npm. Excel for Windows answers in en-US and ru-RU, and Excel for Windows as the reference where the locale shows or orders text differently from Excel for Mac.

Changed:

- Where the regional data of Excel for Mac and Excel for Windows part, the engine computes and shows as Excel for Windows does (it followed the Mac); values change, and the values hash of ru-RU books that compare or sort text or write month names changes with them. The synthetic book's hash does not change (`101578a9c6b600b1…`): its values are those of en-US and shown text is no part of the hash.
  - Text order in ru-RU puts Latin before Cyrillic, as in en-US: `"a"<"а"` is TRUE (it was FALSE), in comparisons, criteria (`COUNTIF(A:A,"<м")`), lookups and `SORT`. Inside each script nothing changes: `ё` and accents weigh only where the letters are otherwise equal (`"ёа"<"еб"` is TRUE), `й` follows every word that starts with `и`.
  - Month names in ru-RU have a capital in `TEXT` and in cells (`TEXT(45688,"ДД ММММ ГГГГ")` is `31 Январь 2025`, it was `31 январь 2025`; `МММММ` gives `Я`), also under `[$-419]` in en-US books; the short names have three letters (`фев`, `мар`, `июн`, `июл`, `сен`, `ноя`; they were `февр`, `март`, `июнь`, `июль`, `сент`, `нояб`). Date text is read in any case, as before.
  - Built-in formats by number show as Excel for Windows shows them: in en-US the short date 14, and 22, 27 to 31 and 36, is `m/d/yyyy` (`1/31/2025`, it was `1/31/25`); in ru-RU the times of 20 to 22 and 32 to 35 and the long time `[$-F400]` have no leading zero (`6:00`, it was `06:00`). `readXlsx` reads such a format as its number, as before, and the writer writes the number.
  - Contract: the doc of `NumberFormat` names `m/d/yyyy` as en-US's short date.
  - The ru-RU long date `[$-F800]` stays as Excel for Mac writes it (`пятница, 31 января 2025 г.`); Windows's has no weekday.

Added:

- The case set `collation-more`: 62 cases of the order of text inside Cyrillic and Latin (`ё`, `й`, case and accents before a later letter), measured in Excel for Windows ru-RU.
- Excel for Windows answers for every set the Mac has answers for, in both locales (Excel 16.0.20430: 34,349 formula cases in en-US, 7,297 in ru-RU, 6,999 formulas of the old kind, 12,949 cells shown, and the tables, spill, workbook, structure and xlsx filter sets), with tests against them; where the two Excels part in regional data the engine follows Windows (above), elsewhere the Mac, and `bench/windows/differences.md` lists the cases with their reasons.

## 0.4.2 (2026-10-09)

Internal release; not published to npm. Fixes from a review, books sent to pool threads as text in chunks, `Session.hash()`.

Changed:

- The thread pool sends a book to a thread and back as JSON text, its cells written and read in chunks between turns of the event loop: on the synthetic book of 34,000 rows `openSession` stops the event loop for 62 ms (324 to 395 ms before), a snapshot of a session for 3 ms (234 to 290 ms) and `book()` for 20 ms (230 ms); a session's snapshot takes 29 MB on the calling side (68 MB). The book is taken when the call is made, as before; a book JSON cannot hold as it is goes by structured clone, as before.

Fixed:

- A range made by the range operator found the cells between its ends in more forms: an end under unary plus (`+A1:C1`), an end bound by LET (`SUM(A1:LET(x,C1,x))`), both ends spilled (the corner between them included). Such formulas could be computed before the cells inside the range and miss their edits in a session.
- A session took an edit that turns a serial number into a typed date as no change in a book of the 1904 system, and ignored an edit that changes only the area of an old-kind array formula.
- With `execution: 'eval'`, a formula nested deeper than the JS parser allows (a long flat formula) failed to compile; it now runs as closures, in the same order, with the same values.
- The calculated columns of a table follow the table when rows are inserted or deleted around it, and each new row gets the formula copied to its own row; deleting the first data row keeps the column's formula.
- xlsx: a calculated column formula that cannot be written, and the properties of row and column styles the model does not have, are reported in `issues`; a shared formula group larger than the JS argument limit is read.
- A spill that is settled is no longer recomputed by an edit that does not touch it; results hash without sorting the cells again. Cold compute of the synthetic book is a little faster with closures; the values and hashes are the same.

Added:

- `hash()` of a session, in every mode: the values hash without the result. From a thread on the synthetic book of 34,000 rows it takes 0.5 s and only the hash comes back, where `result()` takes 1.6 s.
- Contract: `Session.hash()` and `AsyncSession.hash()`, the values hash of `result()` without building the result.

## 0.4.1 (2026-10-08)

Internal release; not published to npm. The xlsx reader and writer carry hidden rows, filters, the 1904 date system, iteration and built-in formats; the intersection and range operators; `#NULL!` joins `ErrorCode`.

Changed:

- Contract: `ErrorCode` gains `NULL`, Excel's `#NULL!`: the intersection of ranges that do not meet. The literal `#NULL!` was read as `ERROR` (shown as `#ERROR!`) before; `ERROR` remains the error of a formula that does not compute.
- `S1:S3!A1`, with sheets named like cells, is the range operator between the cell S1 and `'S3'!A1`, as Excel reads it: `#VALUE!` on another sheet, A1:S1 on S3 (it was a reference across the sheets S1 to S3, which `'S1:S3'!A1` still is).
- `SUBTOTAL`, `AGGREGATE` and `FREQUENCY` take a union area by area, `FREQUENCY` a reference across sheets too, as Excel does (they answered `#VALUE!` or `#REF!`); a reference across sheets that `SWITCH` compares is `#REF!`.
- One value where a function takes a table, an array or dates, as Excel takes it: the table of `VLOOKUP`, `HLOOKUP` and `MATCH` (`#N/A`, `#VALUE!` for text that is no number), `XLOOKUP`'s lookup array (the return array's one row or column), `WRAPROWS` and `WRAPCOLS` of fewer values than the count (one line, not padded), text as the array of `F.TEST`, `FTEST` and `MIRR` (no numbers, `#DIV/0!`), a boolean or text as `XIRR`'s and `XNPV`'s values or dates (`#VALUE!` before the sizes count).
- Financial and date functions at edges, as Excel computes them: `ACCRINT` on basis 1 over several coupon periods, and with a settlement on a coupon date; a coupon period before 1900 (`ACCRINT` 0, `PRICE` `#NUM!`); serial 0 as `ACCRINT`'s issue or first interest date and in `XNPV`'s dates (`#NUM!`); the first period of `AMORDEGRC` and `AMORLINC` capped at cost less salvage, and `AMORDEGRC`'s period that passes the salvage; date text among the holidays of `WORKDAY` and `NETWORKDAYS` (it was `#VALUE!`).
- A row of height 0 is hidden for `SUBTOTAL` 101–111 and `AGGREGATE`, and on a filtered sheet for `SUBTOTAL` 1–11 and the structure operations, in a book, in a session's `setLayout` and in the thread pool, as Excel takes it (an xlsx round trip through Excel); it was counted unless `hidden` said so.
- Contract: a row of `height` 0 is hidden (`RowInput.hidden`) and a column of `width` 0 too (`ColumnInput.hidden`), as Excel takes them; `SheetInput.filtered` says that a table's filter with criteria makes its sheet filtered as well as the sheet's autofilter.
- `readXlsx` reads what the book now holds instead of approximating it: hidden rows and columns get `hidden: true` and keep their sizes (they were read with a size of 0), and a row of height 0 or a column of width 0 is hidden as Excel takes it; a file of the 1904 date system gives `dateSystem: 1904` with its numbers as stored (they were moved to the 1900 system); `calcPr iterate` gives `iteration`; a built-in number format is read by its number, 1 to 49 (it was the code Excel for Mac en-US shows). The issue code `date1904` is gone, `hidden` is about hidden sheets alone, `calculation` no longer reports iterative calculation, and `numberFormat` reports only the East Asian formats 50 to 58 and unknown ones.
- `writeXlsx` writes the codes that ru-RU shows otherwise than built-in formats 37 to 40 (`#,##0_);(#,##0)` and the like) as codes of the file's own, so Excel ru-RU shows what `format` says, and writes `calcPr` in every file, as Excel otherwise keeps iterative calculation on from a book opened before.

Added:

- The intersection operator, a space between references (`SUM(A1:C3 B2:D4)`, `#NULL!` where they do not meet), and the range operator between references that are not one range (`A1:INDEX(B1:B9,3)`, `Sales[Qty]:Sales[Price]`), read and computed as in Excel, in xlsx files and `fromLegacy` too; a formula with the range operator reads every cell between its ends, a spill reference as an end (`A1:E2#`) as far as the array spills (`docs/references.md`).
- xlsx: a sheet's filter mode both ways. `readXlsx` gives `filtered: true` to a sheet whose autofilter has criteria (with Excel's mark `filterMode` or without it), whose table's filter has criteria, or that has the mark without an autofilter (an advanced filter), as Excel takes them; the autofilter's range and criteria are reported (`autoFilter`). `writeXlsx` writes a filtered sheet with criteria every row meets, on its first table with a header row or on an autofilter of its own with the mark, so Excel counts its hidden rows as filtered. Hidden rows and columns, `dateSystem` (`workbookPr date1904`, dates typed in as 1904 serial numbers) and `iteration` (`calcPr iterate`, `iterateCount`, `iterateDelta`) are written too. Checked through Excel 16.113: 8 files with filters in each form and 9 more books of the round trip, their answers predicted before Excel ran them.

## 0.4.0 (2026-10-08)

Internal release; not published to npm. Adds asynchronous and threaded computation (`createInlineCalc`, `xlcalc/worker`), Excel tables with structured references, `Session.cells` and `copyFormula`.

A synchronous session call that throws while it recomputes (over its limits, say) leaves the session as it was before the call; the session computes its book again on its next call.

Fixed:

- The thread pool: a book or an edit that a structured clone cannot carry to a thread (a function in a cell's `meta`) is refused with `InputError` and the thread stays free; it used to leave the thread busy for good, and an answer arriving later threw an uncaught `DataCloneError`.
- The thread pool: a session that leaves a full thread (`maxSessionsPerThread`) takes its book along as a new snapshot, so it comes back without replaying its journal, and a restore after a lost thread has the hard deadline of the opening and of each call it replays; such a session was closed with `restore_failed` when its journal took longer to replay than one task's `hardTimeoutMs`.
- Computation with a deadline or slices looks at the clock before each formula: 300 formulas that each sort 10,000 values ran 698 ms past a soft deadline of 100 ms without a yield, now they stop at 100 ms with slices of 10 to 21 ms.
- Closing: a session's `close()` waiting when its pool or its calculator closes resolves (it hung in a pool and was refused in the same thread), and so does one whose thread is lost while it closes.
- `maxQueue` counts only the tasks no idle thread takes: `maxQueue: 0` refused every task.
- `maxHeapMb` with `spawn` throws `InputError`: such a worker has no memory limit, and the option was ignored.
- A restore that fails leaves nothing of the session in the thread; each one kept its book there until the thread died.

Changed:

- The call after a stopped one computes the session's book again outside its own `timeoutMs` (its `signal` and a pool's hard deadline still bound it), so a session whose cold computation is longer than its calls' deadline no longer fails every call after one has stopped; such a call can take longer than its `timeoutMs` by that computation.
- `CODE` of a character outside the code page of the book's locale is 63 (`?`), as Excel for Windows answers (it was 95, `_`, Excel for Mac's answer): `CODE("中")`, and `CODE("Ж")` in en-US.

Added:

- `cells(sheet, range)` of a session, in every mode: the cells of a rectangle as `result()` gives them, for a window over a sheet. From a thread only the rectangle comes back: 20,000 cells of the synthetic book of 34,000 rows in 27 ms, where the whole result takes 1.7 s.
- `copyFormula(text, from, to, book?)`: the text of a formula copied from one cell to another, as `copyRange` pastes it (`docs/structure.md`).
- Contract: `Session.cells(sheet, range)` and `AsyncSession.cells(sheet, range)`, the cells of a rectangle as `result()` gives them, with each cell's own `style` (`CellView`), for a window over a sheet.
- Excel for Windows is a second oracle: `bench/windows` sends the Mac harness's books to Excel for Windows over SSH and COM, stores its answers in `test/excel/win/`, and `test/excel-win.test.ts` checks the engine against them (`bench/windows/README.md`).
- Computation in the same thread in portions, `createInlineCalc`: it gives way to the event loop between formulas (slices of `sliceMs`), and a call stops at its soft deadline (`timeoutMs`) or when its `signal` fires, with `CalcError`. A session's calls are taken in order, and a call that fails changes nothing (`docs/async.md`).
- Computation in threads, `createCalcPool` from the new subpath `xlcalc/worker`: Node `worker_threads` or Web Workers on one protocol. A session's book lives in its thread; edits go there and changes come back. A task past `hardTimeoutMs` terminates its thread, `maxHeapMb` bounds a Node thread's heap, `maxQueue` and `maxSessions` refuse more work, `maxSessionsPerThread` sends the session called least recently off a full thread. The pool keeps a snapshot and a journal of each session: a session whose thread was lost, or that left a full thread, comes back on its next call with the hash it would have had. Checked against the synchronous engine on the case books in both execution modes, and in headless Chromium.
- `limits` in `compute` and `openSession`: a book with more formulas than `formulas`, or a range, an array or a spill of more values than `cells`, throws `CalcError` instead of computing; an array too large is refused before it is made (`SUM(SEQUENCE(20000,20000))` used to bring the Node process down).
- `CalcError`, exported from `xlcalc` and `xlcalc/worker`, with `code` (`soft_timeout`, `aborted`, `hard_timeout`, `out_of_memory`, `thread_failed`, `restore_failed`, `limit`, `queue_full`, `too_many_sessions`, `closed`).
- `playwright-core` 1.63.0 joins the development dependencies, for the browser test.
- Contract: `CalcLimits` (`formulas`, `cells`) in `ComputeOptions.limits` and `SessionOptions.limits`; a book over them throws `CalcError` with code `limit` instead of computing. The types of asynchronous and threaded computation: `Calc`, `AsyncSession`, `CalcOptions` (`signal`, `timeoutMs`), `AsyncComputeOptions`, `AsyncSessionOptions`, `AsyncRestructureResult` (as `SessionRestructureResult` without the book), `InlineCalcOptions`, `CalcPoolOptions`, `WorkerLike`, `AbortSignalLike` and `CalcErrorCode`.
- Contract: tables (Excel's list objects). `SheetInput.tables` lists a sheet's tables (`TableInput`: `name`, `range`, `headers`, `totals`, `columns` with each column's `name`, the `formula` of a calculated column and its `totals` function or label; type `TotalsFunction`). Tables and defined names share one namespace. Structured references (`Sales[Amount]`, `Sales[[#Headers],[#Data],[Qty]:[Price]]`, `Sales[[#This Row],[Qty]]`, `[@Qty]` inside the table) are part of formula text. `ValidateContext.tables` checks them in `validateFormula`, whose result gains `tables`; the check code `unknown_name` also covers an unknown table or column. `StructureOp` gains `addTable`, `renameTable`, `deleteTable` and `convertTable`, `StructureRefusal` gains `table` and `table_name`, and `Session.setCells` renames a column when its header cell is edited.
- Tables (`docs/tables.md`): a sheet's tables are computed with structured references of every form Excel takes (`Sales[Qty]`, `Sales[[Qty]:[Price]]`, `Sales[#All]`, `[#Data]`, `[#Headers]`, `[#Totals]`, their combinations, `Sales[@Qty]` and `[@Qty]` inside the table, Excel's escapes, any case, spaces where Excel takes them), stored as an xlsx file stores them (`Sales[[#This Row],[Qty]]`, `Sales[]`). A part the table lacks is `#REF!`, the formula's row outside the data `#VALUE!`; names can hold structured references; no array spills inside a table or into one. Structure operations do to tables what Excel does: rows and columns inserted and deleted grow, shrink, move and delete them, fill calculated columns, name new columns `Column1` (`Столбец1` in ru-RU) and lose references to deleted columns, with Excel's refusals (code `table`); a table emptied of data keeps one row; header cells typed over rename their columns (`Session.setCells`); copies and moves name the table in references that leave it, copy whole tables as new ones and move them; `addTable`, `renameTable`, `deleteTable` and `convertTable`. Checked against Excel 16.113 on 91 cases in en-US and ru-RU, two sets of them predicted before Excel ran them.
- `readXlsx` reads tables (table parts) and `writeXlsx` writes them; tables are no longer an issue, and the code `table` now reports a table left out, a header cell given its name or a structured reference without a table name. A formula in a table's totals row is written as a plain formula and recorded in the table part, as Excel needs to open the file.
- Locales: structured references translate (`Locale.tableItems`, ru-RU's `#Все`, `#Данные`, `#Заголовки`, `#Итоги`, `#Эта строка`, the list separator between items), and `toLocal` shows them as Excel does (`Sales[@Qty]`, `Sales`), with `opts.table` as inside that table (`[@Qty]`).

## 0.3.1 (2026-10-07)

Internal release; not published to npm.

- `xlcalc/unstable/internal` imported before the root entry threw `ReferenceError: Cannot access 'collator' before initialization` (a module cycle between the values and the criteria). Every subpath now loads on its own, and a test imports each one first in a fresh process.

## 0.3.0 (2026-10-07)

Internal release as `xlcalc`, used by our own applications before a public release; not published to npm.

Breaking changes:

- The package is renamed `xlcalc` (it was `formula-engine`): `npm install xlcalc`, imports from `xlcalc`, `xlcalc/contract`, `xlcalc/unstable/internal` and `xlcalc/unstable/synth`; the repository is `github.com/drakosha/xlcalc`.
- Every values `hash` changes: the configuration version is now `xlcalc/1;canon15`. It was `formula-engine/1;compiled-js;formulajs@4.6.1;canon15` in 0.2.0 and `formula-engine/2;compiled-js;canon15` in the development builds after it; `compiled-js` is gone since formulas run as closures by default. The synthetic book (seed 1, 3,400 rows, clock `2026-10-04T12:00:00Z` with seed 7) keeps its values and hashes to `101578a9c6b600b1…`; with the earlier version strings the same values still hash to `8d63bad2c247ae26…` (development builds), `18be025cb4581543…` (0.2.0) and `9d7414efbda34a63…` (0.1.0), pinned by a test.
- No runtime dependencies: `@formulajs/formulajs` is no longer a dependency. Its code, and jStat's, is ported into `src/functions/` with attribution (`THIRD_PARTY.md`).
- Numbers become text as Excel converts them (`&`, `CONCAT`, text arguments): 15 significant digits, positional notation up to 20 characters (sign aside), else scientific with a two-digit exponent: `"a"&0.0000000745` is `a0.0000000745` (was `a7.45e-8`), `"a"&(1/134228)` is `a7.4500104300146E-06`, `"a"&1E+21` is `a1E+21`. `TEXT(x,"General")` follows Excel's 11-character General format (`TEXT(1/3,"General")` is `0.333333333`).
- Number literals and numeric text longer than 15 significant digits keep 15 and drop the rest (`123456789012345678` is 123456789012345000, as in Excel; it was rounded); text that is not a decimal number (`"0x10"`) no longer reads as a number. A result below the smallest normal number (2.2250738585072E-308) is 0.
- `SUMIF`, `SUMIFS` and `COUNTIFS` compare the referenced sizes of their ranges: ranges of different sizes are `#VALUE!` (a whole column next to a shorter range was computed over the occupied rows), whole columns of sheets with different numbers of rows work (`SUMIFS('S2'!A:A,'S3'!B:B,">0")` was `#VALUE!`), and `COUNTIFS` counts the blank cells of whole columns when every criterion matches a blank (`COUNTIFS(A:A,"")` is 1,048,576 minus the filled cells).
- An error as a criterion, or its text in any case (`"#N/A"`, `"<>#DIV/0!"`), matches the cells holding that error in every function with criteria (`SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`, the database functions), and `<`, `>` order errors as `ERROR.TYPE` does, as in Excel: `COUNTIF(A1:A9,NA())` counts the `#N/A` cells (it was `#N/A`), `COUNTIFS(A1:A9,"#DIV/0!")` the `#DIV/0!` cells (it was 0); text that reads like an error is not matched. A number or date in a criteria header of a database function names the field with that header (it was read as a column number).
- Values that change because they now match Excel en-US: a number criterion matches text that reads as that number (`COUNTIFS(A:A,2)` counts `"2"`); `DATE(1900,2,29)` is 60, `EOMONTH` reads serial 0 as January 1900 and 60 as February 1900 (`EOMONTH(0,1)` is 59, `EOMONTH(60,1)` is 91) and reads date text (`EOMONTH("2025-01-31",1)`); a month out of range in `DATE` and `EOMONTH` is `#NUM!` (was `#VALUE!`); `ROUND` with a huge digit count returns the number (was `#NUM!`); `TEXT` formats fractions (`TEXT(4.34,"# ?/?")` is `"4 1/3"`) and gives `""` for an empty format; `0^0` is `#NUM!` and `0^-1` `#DIV/0!`.
- Functions of the former Analysis ToolPak answer `#VALUE!` for a boolean where they expect a number or a date, as Excel does: `EOMONTH(TRUE,1)` and `EOMONTH(40544,TRUE)` are `#VALUE!` (were 59 and 40602); the same holds for the bond, coupon and discount functions, `EDATE`, `YEARFRAC`, `WEEKNUM`, `WORKDAY`, `NETWORKDAYS` and their `.INTL` forms.
- `TEXT` reads every format code that a cell's display reads, with the same engine (`docs/formats.md`), so values that were `#VALUE!` change: exponents (`TEXT(1234.567,"0.00E+00")` is `1.23E+03`, `##0.0E+0` engineering notation), elapsed time (`[h]:mm`, `[mm]:ss`, `[ss]`) and fractions of a second (`mm:ss.00`), colours (`[Red]`, `[Color10]`; in ru-RU the Russian names `[Красный]`, `[Цвет10]`, where `[Red]` stays `#VALUE!`), conditions (`[>100]0.0;[<=-100]-0;0`, a section that takes only negative numbers writes no minus), padding (`_)` is a space), fills (`*x` is dropped), currency and locale tags (`[$₽-419]`, `[$-409]mmmm`, `[$-F800]` the long date, `[$-F400]` the long time); a value that no conditional section of two or three takes is `#VALUE!`; a number of more than 255 characters is `#VALUE!` (`TEXT(1E+300,"0")`) and up to that it is written out (`TEXT(1E+250,"0")` was `1e+250`); a fraction whose numerator passes 2,147,483,646 is `#VALUE!`. Checked against Excel on 771 cases (`text-format`, `text-format-more`, `text-format-ru`, `text-format-ru-more`); the ru-RU cases of exponents and elapsed time leave `RU_KNOWN`.
- Contract: the error code `CALC` (Excel's `#CALC!`: an empty array, a text over 32,767 characters from `CONCAT` or `TEXTJOIN`); code that matches every `ErrorCode` meets a new member. The check code `unsupported` now means a result that is an array of several values (or an external function with non-literal arguments).
- Arrays inside a formula, as Excel 365's dynamic arrays (`docs/arrays-incell.md`): a range of several cells or an array where a function takes one value, or under an operator, is computed element by element (`SUM(B1:B9*C1:C9)`, `SUM(IF(A1:A9>0,A1:A9))`, `SUMIF(A:A,{"a","b"},B:B)`); such formulas had the check code `unsupported` and `#ERROR!`. A formula whose result is an array of several values gives `#SPILL!` (was `#ERROR!`), with an issue that names the size and has the check code `unsupported`; `validateFormula` answers `unsupported` with a new message where the result is an array whatever the data. The functions of the former Analysis ToolPak and the database functions' field answer `#VALUE!` for a range of several cells in a one-value argument, as Excel 365 does (`BESSELJ(A1:A2,1)`, was `#ERROR!`).
- `INDEX` with row or column 0 answers the row, column or range (was `#VALUE!`: `SUM(INDEX(A1:C3,0,2))` sums the column), and `INDEX` and `XLOOKUP` answer references: `SUM(INDEX(A1:A3,2))` skips text in A2, as Excel does (it was read as a number). `ROW` and `COLUMN` of several rows or columns answer all of them (`SUM(ROW(A1:A3))` is 6, was 1; `ROW(A1:A3)` alone is `#SPILL!`, was 1). `TYPE` of a range of several cells is 64 (was 16). `XLOOKUP` with a return array wider (taller) than its lookup line answers the row (column) found (was `#VALUE!`). `VLOOKUP` and `HLOOKUP` take an array or a range of several cells as the column (row) number, as Excel 365 does: the answer is the found row (column) at each number, shaped as the index, and the first invalid number answers for all, `#VALUE!` below 1 and `#REF!` past the table or for an error or non-numeric text (`VLOOKUP(5,A1:C6,{2,3},FALSE)` is an array, `VLOOKUP(5,A1:C6,{1;4},FALSE)` `#REF!`; a range was `#VALUE!`); `MATCH` takes the top-left cell of a range as its match type. `VALUETOTEXT` of a range of several cells is computed for each cell (was `#VALUE!`). `CHOOSE` answers the chosen argument as it is, a range too.
- Values that change because the random check (below) found Excel's rule: text in numeric arguments, in operators and in `TEXT` reads as `VALUE` reads it (`ABS("2025-01-31")` is 45688, `"50%"*2` is 1, `SIGN("$1,000")` is 1; they were `#VALUE!`); where several arguments are wrong, the error is the one Excel finds first (values before references and arrays, the ToolPak functions' basis, frequency, places and suffix as Excel reads them, a one-cell error and then sizes before errors inside paired arrays: `LARGE(A1:A9,"x")` is `#VALUE!` with an error in A1:A9, `PRICEDISC(NA(),39508,0.0525,100,7)` is `#NUM!`); `LINEST`, `LOGEST`, `TREND` and `GROWTH` give a collinear variable coefficient 0 (were `#NUM!`) and check sizes before values; `ISPMT` no longer truncates its period and count; a database without records, or a field taken from a cell holding an error, is `#VALUE!`; an empty delimiter in `TEXTSPLIT` is `#VALUE!`; `PERCENTOF` answers an error given as a value; an odd coupon period that begins before 1900 is `#NUM!` in `ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`; the complex-number functions read number text (`IMABS("50%")` is 0.5, was `#NUM!`); an error among `TREND`'s or `GROWTH`'s new x is `#VALUE!`; `XIRR` answers `#VALUE!` for a boolean or other text among its dates before `#NUM!` for a blank one.
- Text order for `<`, `>`, criteria and lookups follows the en-US collation, Latin before Cyrillic, as Excel en-US sorts (`"x"<"а"` is TRUE; it was FALSE with the ru-RU order).
- Time text in date and time arguments is read as Excel en-US reads it: without AM/PM up to 9999 hours and any minutes and seconds (`"25:00"` is 25 hours, `"10:60"` is 11:00, `"1:99"` is 2:39), with AM/PM hours 0 to 12 and minutes and seconds below 60 (`"0 PM"` is noon, `"13:00 PM"` is not a time). Such text was `#VALUE!` before.
- Sheet names follow Excel's rules: a book with a sheet name over 31 characters, with one of `: \ / ? * [ ]` or with an apostrophe at either end throws `InputError`, and `restructure` refuses to rename a sheet so (`sheet_name`).
- Deleting columns under a span of columns that reaches XFD keeps it reaching XFD (the columns that come in at the end take its entry; its style and format only where XFD itself is deleted), and deleting the sheet's last rows or columns keeps a merged area that ends there, as Excel does (the struct sets `edge-layout`, `edge-wide`); the span was cut short and the area shrank.
- The sheet name `History`, in any case, is reserved as in Excel, which refuses to rename a sheet so (the struct set `history`): a book with such a sheet throws `InputError` and `renameSheet` to it is refused with `sheet_name`.
- Unknown fields of a book and of a sheet throw `InputError`; they were ignored.
- Values that change because they now match Excel en-US (the `core` and `near-integer` sets): an operator converts its left operand before the right one's error counts (`"abc"-NA()` is `#VALUE!`, was `#N/A`); an array of one value stays an array under operators and lifted calls (`TYPE({1}+0)` is 64, `SUM({TRUE}=TRUE)` 0, `DEVSQ({1}>=TRUE)` `#NUM!`), `INDEX` of a line selected by 0 is an array, `TRANSPOSE` of a value is that value; `IF`, `IFS`, `SWITCH`, `CHOOSE` and `XLOOKUP`'s `if_not_found` pass a one-cell reference on (`SUM(IF(TRUE,A2))` skips text in A2, it was read as a number; `COLUMN(IF(TRUE,C1))` is 3, was `#VALUE!`); `HSTACK` and `VSTACK` answer an error given as a value (`VSTACK(1,1/0)` is `#DIV/0!`, it was stacked); `VLOOKUP` and `HLOOKUP` lifted by the key read the index's top-left number in each call; comparison operators and criteria take numbers equal to 15 significant digits as equal (`0.1+0.2=0.3` is TRUE, was FALSE), and `INT`, `TRUNC`, `ROUNDDOWN` and `ROUNDUP` read their number to 15 digits (`INT(1-2^-52)` is 1, was 0); whole-number arguments of `LEFT`, `INDEX`, `ROUND`, `DATE` and the others in the README take a number less than 2^-22 short of an integer as that integer, `ISEVEN` and `ISODD` one less than 10^-10 short; `DAY`, `MONTH`, `YEAR` and `WEEKDAY` read a date and time to the nearest second; `DATE` rounds a day below zero down; `EVEN` rounds exactly (`EVEN(2+2^-51)` is 4, was 2); `AGGREGATE` 14–19 take the top-left cell of a range given as `k` (was `#VALUE!`).
- Criteria read numbers, dates and times in their text and in the text of cells as `VALUE` does, and a number criterion's `<>` holds for all text, as Excel en-US does (the `criteria-text` set): `COUNTIF(A1,">=1/1/2025")` and `COUNTIF(A1,"Jan 31, 2025")` compare dates (they compared text), `COUNTIF(A1,"$1,000")` matches 1000, `COUNTIF(A1,0.5)` counts a cell holding the text `50%`, `COUNTIF(A1,"2025-01-31")` a cell holding that text, and `COUNTIF(A1,"<>2")` counts the text `2` (it did not).
- `DATE` keeps Excel's limits on its arguments (the `date-day` set): a day outside −32768…32767 is taken as 32767 (`DATE(2025,1,40000)` is 78424, it was 85657), a month of 32767 or more is `#NUM!`.
- `AGGREGATE` 14–19 take an array as well as a reference, as Excel does: `AGGREGATE(14,6,10/A1:A9,1)` is the largest quotient (it was `#VALUE!`).
- `AND`, `OR` and `XOR` take a direct text argument `"TRUE"` or `"FALSE"` (any case) as that boolean, as Excel en-US does (the `logical-text` set): `OR("TRUE")` is TRUE (it was `#VALUE!`); text in cells and arrays is still skipped.
- Values that change in en-US books because they now match Excel en-US (the `date-noyear` set): date text without a year is a date of the clock's year, as Excel completes it with the current year (`VALUE("1/31")`, `DAY("Jan 31")`, `COUNTIF(A:A,"1/31")`, `TEXT("31-Jan","yyyy")`; such cells are volatile), a number that is no day of its month is a year (`2/29` in a year that is not leap is February 2029, `1/32` January 1932, `1/2025` January 2025); a month is three letters or more of its name (`Febr 3`); spaces may stand around `/` and `-` (`1 / 31 / 2025`), and `1-31-2025` is a date; `Jan 31 2025` without a comma is no date (it was read).
- Time text is read by one rule in `VALUE`, operators and date and time arguments, as Excel en-US reads it (the `time-text` set): AM or PM only after a space (`HOUR("6pm")` is `#VALUE!`, it was 18), `A` and `P` alone (`6 P`), spaces around `:` (`6 :30`), a `:` at the end without AM or PM (`6:`, `6:30:15:`), minutes and seconds with decimals (`6:30.5` is 6 minutes 30.5 seconds), hours of three or four digits without a leading zero (`VALUE("00123:00")` is `#VALUE!`, it was 5.125), a time before the date (`18:00 1/31/2025`); `2025-01-31T18:00` is no date (it was read).
- Date and time text after a space or a tab is no date or time in date arguments, numeric arguments, operators, criteria and `TEXT`, in both locales, as in Excel (the `lead-spaces` set): `DAY(" 1/31/2025")`, `" 6:30"+0` and `ABS(" 6:30")` are `#VALUE!`, `COUNTIF(A1," 2025-01-31")` compares text (they read the date); `VALUE`, `DATEVALUE`, `TIMEVALUE` and `DAYS` drop the spaces, and numbers keep taking them (`" 5"+0` is 5). A tab is no space in date and time text (`VALUE(CHAR(9)&"6:30")` is `#VALUE!`).
- The parts of complex text are read as `VALUE` reads a number in en-US too, as in ru-RU (the `complex-text` set): `IMREAL("50%+i")` is 0.5 and `IMSUM("$1+2i","1+i")` is `2+3i` (they were `#NUM!`), the first part with its own sign (`IMREAL("-(1)+i")` is 1); spaces count as Excel counts them (`IMREAL("1 000+2i")` and `IMREAL("1+ i")` are `#NUM!` in en-US, they were 1000 and 1).
- `TEXT` rounds a number that `%` or a thousands scale multiplies as Excel shows the product, to 15 significant digits, before the format's own rounding (the `format-scale` set): `TEXT(0.285,"0%")` is `29%` (it was `28%`), `TEXT(0.145,"0%")` `15%`.
- Contract: arrays spill, as in Excel 365 (`docs/spill.md`). `CellResult` and `CellChange` gain `spill` (the range a formula's array fills), `spillOf` (the formula a spilled cell belongs to) and `spillFill` (the one value of a large array's cells past its listed part, such as the 0s of `A:A*2` below the data); `ExplainNode` gains `spillOf`. Spilled cells are values of the book: they are in `sheets[].cells` and in the values `hash` (a book without spills hashes as before). The check code `unsupported` now means only an external function called with arguments that are not literals; `validateFormula` accepts a formula whose result is an array. `Session.setLayout` returns the changed values, since merged areas stop arrays spilling over them.
- `validateFormula` (and `compute`, with `#ERROR!`) refuses the argument counts Excel does not load: `SORTBY` with 4, 6… arguments (each further by-array needs its order), `SWITCH` with more than 126 pairs; the check code is `syntax`.
- A function newer than Excel 2007 (stored with the `_xlfn.` prefix in a file) takes at most 254 arguments, as Excel loads it (the `arity` set): `CONCAT`, `TEXTJOIN`, `STDEV.S`, `STDEV.P`, `VAR.S`, `VAR.P`, `MODE.SNGL`, `MODE.MULT`, `SKEW.P`, `XOR`, `CHOOSECOLS`, `CHOOSEROWS`, `HSTACK`, `VSTACK`, `AGGREGATE` with 255 arguments, and `SORTBY`, `MAXIFS`, `MINIFS` with 255, are the check code `syntax` (they were computed); the old names (`STDEV`, `VAR`, `MODE`) take 255.
- Values that change with arguments left out (the `omit` sets): `COUNT` counts a blank given directly (`COUNT(IF(TRUE,))` is 1, it was 0) and `AND`, `OR` and `XOR` take it as FALSE; `SORTBY` with one value as a by-array keeps the order (it was `#VALUE!`); `LOOKUP` past the end of a result that is a value or an array is `#N/A` (it was `#REF!`; past the end of a range it still is); `RATE` whose balance only vanishes as the rate goes to −1 (`RATE(48,0,8000)`) is `#NUM!`, as in Excel (it was −0.58).

Added:

- Hidden rows and filters as Excel computes them. `SUBTOTAL` 101–111 and `AGGREGATE` options 1, 3, 5 and 7 skip hidden rows, `SUBTOTAL` 1–11 skips the hidden rows of a filtered sheet, and hidden columns count. A session recomputes these formulas when rows are hidden or shown or the filter mode changes. On a filtered sheet structure operations act on the visible rows: deleting rows deletes the visible ones, inserting rows inserts as many as are visible, a copy leaves out hidden rows and pastes values, a paste goes into the visible rows. Inserted rows and columns are shown, even inside a hidden block. Checked against Excel on 15 workbook books with 4,758 cells (3 of them predicted before Excel ran, all as predicted) and 49 structure books.
- The 1904 date system as Excel computes it: dates typed in, date text, `DATE`, `TODAY`, `NOW` and the date functions count from 1904-01-01, the former Analysis ToolPak's functions take date text given as a value as Excel does there, and cells and `TEXT` show a negative date or time with a minus. Checked against Excel on 5 books with 384 cells, one of them with 46 cells predicted before Excel ran and 45 as predicted.
- Built-in number formats by number, shown by the locale as Excel shows them: en-US's codes of Excel for Mac, ru-RU's own for dates, times, money and negatives. Checked against Excel on 1,300 cells in both locales, 700 of them predicted before Excel ran, all as predicted.
- Iterative calculation as Excel computes a book it opens: the cells of every cycle and their dependents from 0, sheet by sheet by name and row by row, until the iteration count or until no number changes by the maximum change. Checked against Excel on 78 books with 546 cells, 8 of them predicted before Excel ran and 7 as predicted. Excel goes on from the values it holds at a recalculation, the engine starts from 0 at every computation (a known difference).
- The harness's `workbook` step (`bench/coverage/workbook.ts`): books whose file holds what the cells alone do not (hidden lines, filters, the date system, iterative calculation, stored values), steps Excel takes after opening them, and their cells read at each step.
- Contract: hidden rows and columns and filtered sheets. `RowInput.hidden` and `ColumnInput.hidden` hide a line, as xlsx `hidden` does. `SheetInput.filtered` says the sheet's autofilter has criteria: Excel then counts every hidden row of the sheet as one a filter hid. `LayoutEdit` gains `{ sheet, filtered }`, and `setLayout` takes `hidden` in its row and column patches. All of them go into `inputHash` only when present.
- Contract: `BookInput.dateSystem`, 1900 by default or 1904, as xlsx `date1904`. A 1904 book counts days from 1904-01-01, serial 0. `DisplayOptions.dateSystem` shows a value by it. A 1904 book goes into `inputHash` and the values `hash` (`;date1904` after the configuration version), books without it hash as before.
- Contract: a number format can be the number of one of Excel's built-in formats, 0 to 49, as an xlsx file stores a cell's format with no code (`NumberFormat`, in `CellInput.format`, `ColumnInput.format`, `RowInput.format`, `DisplayOptions.format` and `CellResult.format`). Excel shows some of them by the locale: `format: 14` shows `1/31/25` in en-US and `31.01.2025` in ru-RU. Code that reads `format` as a string now meets a number too.
- Contract: `BookInput.iteration`, iterative calculation of circular references with `maxIterations` and `maxChange`, as xlsx `calcPr iterate`. It goes into `inputHash` when present.
- Contract: the refusal `filtered` of `restructure`. Excel cannot paste a copy that left out rows a filter hid over hidden rows of a filtered sheet, and the engine refuses it too.
- Names of any formula (`docs/names.md`): a constant, a calculation, a reference from `INDEX`, a union, references without `$` that move with the cell using the name and wrap round the sheet's edges, `ROW()` and `COLUMN()` of that cell. A name's formula sees the global names. A name qualified by a sheet (`'T'!rate`) is that sheet's own name, else the global one. A cycle through names is `#CYCLE!`. A name over an unknown sheet or function gives `#REF!` or `#NAME?` with a name issue. Checked against Excel on 26 books, 10 of them with 24 cells predicted before Excel ran, all as predicted.
- The `@` operator, unions and references across sheets (`docs/references.md`). `@A1:A10` takes the cell in the formula's row or column, an array's top-left value. `(A1:A3,C1:C3)` is a union of ranges, `'Jan:Mar'!A1:B2` the same rectangle on every sheet from Jan to Mar in the book's order. Functions over lists take each area, `INDEX` takes an area number, `VSTACK`, `HSTACK`, `TOCOL` and `TOROW` stack the areas; the other functions answer the error Excel answers. `AREAS` and `ISREF` join the whitelist (469 functions). Checked against Excel on 989 cases of every function and 8 books of `@`, and on 21 cells of `@` and of references across sheets predicted before Excel ran, all as predicted. Structure operations keep references across sheets under rows and columns and change their ends with renamed, deleted and moved sheets (20 cases).
- `fromLegacy(formula, { names, sheet })`: a formula of the old kind, as a file without dynamic array metadata stores it, as the dynamic formula Excel 365 shows for it, with `@` where Excel puts it (`docs/legacy.md`). It gives Excel's text on 7,000 formulas: every whitelisted function with a range and an expression in each argument, at the top and nested where arrays are computed, calls that may answer several values, names. The converted formulas compute Excel's values of the old ones at their rows, the coverage measurement computed as formulas of the old kind too; 13 cells predicted before Excel ran came out as predicted.
- `readXlsx` reads formulas of the old kind as the dynamic formulas Excel 365 shows for them, with `@` where Excel puts it, and keeps the text as written where none is needed. It reads array formulas of the old kind into `CellInput.array`, `_xlfn.SINGLE(x)` as `@x`, references across sheets, and names of any formula. `writeXlsx` writes all of them back as Excel stores them. The issue codes `legacyArray`, `implicitIntersection`, `singleOperator` and `reference3d` are gone, and `name` no longer reports constants and formulas. The 75 formulas of the old kind Excel computed and saved compute Excel's values when read, but two with `OFFSET` and one with the intersection operator.
- Array formulas of the old kind, entered with Ctrl+Shift+Enter (`CellInput.array`): the answer fills the area as Excel fits it, `ROW()` and `COLUMN()` give the area's rows and columns, nothing spills over the area, a session refuses edits of part of it, structure operations move it whole and refuse to change part of it. Checked against Excel on 17 books and 17 structure cases, and on 9 books with 25 cells predicted before Excel ran, all as predicted.
- Contract: defined names of any formula. `NameInput.ref` is the name's formula, as Excel's "Refers to" and an xlsx `<definedName>` hold it: a reference with a sheet (as before), a constant (`0.2`, `"text"`, `{1,2,3}`), a formula (`'Data'!$A$1*2`, `INDEX('Data'!$A$1:$C$9,2,0)`), a union of references (`'Data'!$A$1:$A$3,'Data'!$C$1:$C$3`). A reference without `$` is relative to A1 and moves with the cell that uses the name. Books whose names are references compute as before.
- Contract: `StructureOp` gains `insertSheet` (a new empty sheet, before a sheet or at the end) and `moveSheet` (a sheet moved before another or to the end), for books with references across sheets, which cover the sheets between their ends in the book's order.
- Contract: array formulas of the old kind (Ctrl+Shift+Enter). `CellInput.array` is the area of such a formula, on its top-left cell; `CellResult.spill` and `spillOf` (and the same fields of `CellChange`) describe the area and its cells; `Session.setCells` throws `InputError` for edits of part of an area, and `restructure` refuses an operation that changes part of one with the new `StructureRefusal` `array`.
- Contract: what a cell shows, as Excel shows it (`docs/formats.md`). `CellResult` and `CellChange` gain `text`, the cell's value by its number format (its own, else its row's, else its column's) in the book's locale and its column's width (Excel's `Range.Text`: General narrowed to the width, `#` across the column for a number that does not fit, fills `*x` repeated across it, text as it is), and `color`, the colour of the format's section (`[Red]`, `[Color10]`) as `#RRGGBB`; `value` stays the value. They are on by default; `ComputeOptions.display` and `SessionOptions.display` set to false leave them out (values alone, nothing spent on formatting). They are not part of the values hash. A session returns the cells whose shown text or colour changed: a new format through `setCells`, a width, a format or a merged area through `setLayout`, a moved column through `restructure`. `displayValue(value, { format, locale, width })` (also `FormulaEngine.displayValue`, types `CellDisplay` and `DisplayOptions`) shows a value without a book. `CellInput.format` is documented as an xlsx format code (en-US syntax whatever the locale) and `ColumnInput.width` as the xlsx width with Excel's default when absent.
- Contract: `BookInput.locale`, the locale the book is computed in, a BCP 47 tag of one of the package's locales (`'en-US'`, the default, or `'ru-RU'`); an unknown tag throws `InputError`. A locale other than en-US goes into `inputHash` and into the values `hash` (`;locale=<tag>` after the configuration version), so books without it and en-US books hash as before and `CONFIG_VERSION` does not change.
- Computing by a book's locale: a book with `locale: 'ru-RU'` is computed as Excel 16.113 for Mac with the Russian interface and locale computes it (`docs/locale.md`): number text with a decimal comma, a space between groups and the ruble after the amount (`VALUE("1 234,5 ₽")`), day-month-year date text with Russian month names (`DAY("31.01.2025")`, `DATEVALUE("15 янв 2025")`), date text without a year in the clock's year (`VALUE("1.5")` is 1 May; such cells are volatile and a session recomputes them on `setClock`), `ИСТИНА` and `ЛОЖЬ` as booleans in text and as the text of booleans, numbers as text with a decimal comma (`"a"&1.5` is `a1,5`, complex numbers too), `TEXT` with Russian format codes (`ДД.ММ.ГГГГ`, `ч:мм`, `# ##0,00`, `Основной`), `DOLLAR` and `FIXED` (`1 234,57 ₽`), `NUMBERVALUE`'s default separators, error texts (`VALUETOTEXT(NA())` is `#Н/Д`) and `ARRAYTOTEXT`'s separators, criteria words and numbers (`">1,5"`, `">=01.01.2026"`, `"ИСТИНА"`, `"#Н/Д"`), and text sorted Cyrillic before Latin. Checked against Excel ru-RU on the coverage measurement's cases (`test/excel-ru.test.ts`) and 1,371 cases of the ru-RU sets (`test/excel/cases/ru-RU/`, `test/excel-ru-probes.test.ts`); `run.ts probes <set> ru` computes a set in Excel ru-RU.
- Spilling, as in Excel 365 (`docs/spill.md`): a formula whose result is an array fills the cells to its right and below (it gave `#SPILL!` with the check code `unsupported` everywhere); `#SPILL!`, with the cell in its way in the issue, where a value (an empty text too), a formula, a merged area or the sheet's edge stops it, while a format or a fill does not; spill references `E2#` (in names too, passed through by `toLocal` and `toInvariant`, moved by structure operations, `#REF!#` and `#REF!E2#` when lost, as Excel writes them); formulas reading spilled cells come after the array, and a session moves them when its size changes. An array over a whole column keeps its data rows and one value for the rest (`spillFill`). Checked against Excel en-US on 88 cases (`test/excel/spill/`; 17 of them held out, predicted before Excel ran, all as predicted) and 12 structure cases. Of two arrays over one cell that do not depend on each other Excel lets either spill, varying between runs; the engine lets the lower one. The tests' case books keep a value right of each case and under the last, as Excel's books do, so array answers stay `#SPILL!` there.
- The option `execution` of `compute` and `openSession` (type `Execution`): `'closures'`, the default, runs each formula as a tree of JavaScript closures and generates no code, so the engine computes under a Content Security Policy without `'unsafe-eval'`; `'eval'` compiles the book's formulas to JavaScript text with one `new Function`, as before. Values, issues and hashes are the same in both, and a session compiles edits in the mode it was opened with. On the synthetic book (3,400 rows, 3,907 formulas; medians of 8 interleaved runs, Apple M5 Pro, Node 26) closures and eval take 285 and 297 ms for a cold compute, 9.3 and 9.1 ms for a criterion edit (p95), 22 and 22 ms for a data edit, 30 and 35 MB of heap, and 15 and 29 ms to compile the parsed formulas alone.
- Array constants `{1,2;3,4}` (rows by `;`, values by `,`: numbers, text, booleans, errors).
- 28 functions with an array answer, usable inside a formula (`SUM(UNIQUE(...))`, `INDEX(SORT(...),1,1)`, `COUNTA(FILTER(...))`) and matching Excel en-US on all of their cases: `CHOOSECOLS`, `CHOOSEROWS`, `DROP`, `EXPAND`, `FILTER`, `FREQUENCY`, `GROWTH`, `HSTACK`, `LINEST`, `LOGEST`, `MINVERSE`, `MMULT`, `MODE.MULT`, `MUNIT`, `RANDARRAY` (volatile, drawn from the clock seed, the cell and the element), `SEQUENCE`, `SORT`, `SORTBY`, `TAKE`, `TEXTSPLIT`, `TOCOL`, `TOROW`, `TRANSPOSE`, `TREND`, `UNIQUE`, `VSTACK`, `WRAPCOLS`, `WRAPROWS`. An empty answer is `#CALC!`, as in Excel. The whitelist has 465 functions.
- `TEXTBEFORE` and `TEXTAFTER` take several delimiters (an array or a range).
- `CHAR` and `CODE`, by the Windows code page of the book's locale (Windows-1252 in en-US, Windows-1251 in ru-RU), as Excel for Windows computes them; Excel for Mac uses Mac Roman and Mac Cyrillic for codes 128–255 (a known difference). Checked against Excel en-US and ru-RU on codes 1–127 and the edges (the `char-code` set) and against Unicode's tables. The whitelist has 468 functions with `LET`.
- The case set `arrays` (`bench/coverage/sets/arrays.ts`, 2,352 cases with Excel's answers in `test/excel/cases/arrays.json`), computed as dynamic array formulas: the harness writes them as Excel 365 saves such formulas (`cm` and the `XLDAPR` metadata). Among them, every whitelisted function's one-value arguments given a range and an array.
- The random check: `bench/coverage/sets/fuzz.ts` generates formulas over the whitelist from a seed (arguments by kind, nested calls and operators, ranges and arrays of the data sheet, number and date text, errors as values; sizes only as small literals), two batches of 2,000 with Excel's answers in `test/excel/cases/fuzz.json`. The set `fuzz-rules` (302 cases) pins the rules it found; differences whose cause lies outside the functions are listed in `test/excel-open.ts`.
- The coverage measurement's Excel columns and the `text` and `matheng` sets are computed as dynamic array formulas.
- Contract: cell styles and the layout of sheets, as an xlsx file keeps them. `BookInput.styles` is a table of `CellStyle` (font: name, size, bold, italic, underline, strike, colour; a solid fill; borders: the line and colour of each side; alignment: horizontal, vertical, wrap, indent), `CellInput.style` an index into it, and a sheet has `columns` (width, style, format, by column `C` or span `B:D`), `rows` (height, style, format, by row `5` or span `5:9`) and `merges` (merged areas `B2:D3`). A cell without its own style or format shows its row's, else its column's. They are checked like the rest of the input, take no part in computation or in the values `hash`, and go into `inputHash` only when present, so a book without them hashes as before; `compute` does not return them.
- Contract: session methods for a host that keeps the book. `restructure(op)` applies a structure operation to the session's book and recomputes; its result is `restructure`'s with `changes`, the cells whose value differs from the one `mapCell` carried there, by new address. `setStyles(styles)` replaces the style table and `setLayout(edits)` changes column widths, row heights, row and column styles and formats (a patch over `C` or `B:D`, `5` or `5:9`, spans split where an edit cuts them) and merged areas; neither recomputes. `book()` is the session's current book, what a host saves as a version. `setCells` with an edit that keeps a cell's value and formula (a new `style`, `format` or `meta`) recomputes nothing and returns no changes. `FormulaEngine` gains `restructure`; `StructureOp`, `RangeRef`, `RestructureResult` and `StructureRefusal` move to the contract, with `SessionRestructureResult`, `LayoutEdit` and `LinePatch`.
- The case sets `core` (258 cases: the order of operators, arrays of one value, references passed on, stacks, lifted lookups, argument counts) and `near-integer` (4,220 cases: exact values near whole numbers, from one unit in the last place to 10^-6 away, in every kind of whole-number argument, in comparisons, criteria, lookups and date functions), with Excel's answers in `test/excel/cases/`.
- `restructure(book, op)`: inserting and deleting rows and columns, moving (cut and paste) and copying (copy and paste, fill) ranges, renaming and deleting sheets, with references in formulas and names adjusted as Excel adjusts them, and cell styles, columns, rows and merged areas changed as Excel changes them (inserted rows take the format of the row above, columns of the column on the left; moved and copied cells keep their look; whole rows and columns take their sizes and styles along; merged areas grow, shrink, move, and are copied in tiles). It returns the new book and `mapCell` (where an old cell is now), or `{ ok: false, code, message }` where Excel refuses (`off_sheet`, `sheet_name`, `last_sheet`, `merged`). Copies have no `meta` unless `withMeta: true`. Checked against Excel en-US on 182 cases, 64 of them with styles (`test/excel/struct/`, `test/structure-excel.test.ts`; `npx tsx bench/coverage/run.ts struct <set>` refreshes them), one known difference (whole columns copied onto part of a merged area: Excel pastes, the engine refuses); the rules are in `docs/structure.md`. As a session method it rebuilds the state: 190 ms for a row inserted into the 3,400-row synthetic book, the cost of a cold computation.
- 90 functions, ported from formula.js and matching Excel en-US on all of the coverage measurement's cases: `ABS`, `ACOS`, `ACOSH`, `ACOTH`, `AND`, `ASIN`, `ASINH`, `ATAN`, `ATANH`, `AVEDEV`, `AVERAGE`, `BINOM.DIST.RANGE`, `BINOMDIST`, `CEILING`, `CEILING.MATH`, `CEILING.PRECISE`, `CHOOSE`, `CLEAN`, `COLUMNS`, `CONCAT`, `CONCATENATE`, `COSH`, `COUNTA`, `COUNTBLANK`, `CSCH`, `DEGREES`, `ERROR.TYPE`, `EVEN`, `EXACT`, `EXP`, `F.DIST.RT`, `FIND`, `FISHER`, `FIXED`, `FLOOR`, `FLOOR.MATH`, `FLOOR.PRECISE`, `GAMMALN`, `GAMMALN.PRECISE`, `GAUSS`, `IFS`, `INT`, `ISBLANK`, `ISERR`, `ISERROR`, `ISLOGICAL`, `ISNA`, `ISNONTEXT`, `ISNUMBER`, `ISPMT`, `ISTEXT`, `LEFT`, `LEN`, `LN`, `LOG`, `LOG10`, `LOWER`, `MAX`, `MEDIAN`, `MIN`, `N`, `NA`, `NORM.DIST`, `NORM.S.DIST`, `NORMDIST`, `NOT`, `NPV`, `ODD`, `OR`, `PDURATION`, `PHI`, `PI`, `POWER`, `PROB`, `PRODUCT`, `RADIANS`, `REPT`, `RIGHT`, `ROUNDDOWN`, `ROUNDUP`, `ROWS`, `SECH`, `SIGN`, `SINH`, `SQRT`, `T`, `TRIM`, `TRUNC`, `UNICODE`, `XOR`. The whitelist has 107 functions.
- `test/excel.test.ts`: every whitelisted function is computed on the coverage measurement's cases and compared with Excel en-US (answers stored in `test/excel/coverage/excel-en.json`); a function is in the whitelist only if all of its cases match.
- `test/excel-probes.test.ts`: 959 probes of rules outside single functions (numbers as text, TEXT General, long literals, subnormal results, criteria over whole columns) against Excel en-US's answers in `test/excel/probes.json` (`npx tsx bench/coverage/run.ts probes` refreshes them).
- Locales of formula text: `toLocal(formula, locale)` shows a stored formula and `toInvariant(text, locale)` reads typed text back the way Excel does in another language (function names, TRUE and FALSE, error literals, the decimal separator, the separators of arguments and array constants), and `validateLocalFormula(text, locale, ctx?)` checks typed text with positions in it. Locales are plain objects (`Locale`): `enUS` and `ruRU`. The ru-RU data were read from Excel 16.113 for Mac with the Russian interface (`test/excel/locale/ru-RU.json`, `bench/locale/`), and both translations reproduce Excel's own text of 6,381 formulas. Stored formulas and computed values do not change.
- Arguments left out between commas (`IF(A1,,1)`, `TEXTSPLIT("a,b",",",,,1)`, `IF(A1;;1)` in ru-RU), as Excel 16.113 takes them: a blank read by the argument's kind as 0, `""` or FALSE, `#VALUE!` where the function takes an array or a reference, the function's default where it treats the argument as not given (`src/functions/omitted.ts`), not given or `#N/A` in the former Analysis ToolPak's functions; a formula leaving out a reference Excel requires is the check code `syntax`. They were syntax errors (`#ERROR!`). Measured on every whitelisted function's arguments beside 0, `""` and FALSE (the sets `omit`, 4,139 cases, and `omit-held`, 55 cases held out: 58 of 61 predictions, the refused calls among them, as Excel answered); the differences the replacements found in other functions are in the README's known differences.
- `LET`, in the whitelist (466 functions): names over values, arrays and references, scoped as in Excel (an inner `LET` shadows, a name shadows a defined name, the case of names aside), each value computed once on its first use, references kept as references, a calculation left off answering the last value, at most 254 arguments; the canonical text spells a name as it is declared. Checked against Excel on the sets `let` (82 cases) and `let-held` (33, held out: 34 of 35 predictions as Excel answered) and the spill set `let` (defined names, spills, formula text; 3 of 3 as predicted). `unstable/internal`: the AST gains the kinds `missing`, `let` and `local`.
- Spilling measured where it was not (`docs/spill.md`): arrays whose size comes from a draw (the spill set `more`, five runs; Excel gives `#SPILL!` when the size changes between its passes, the engine draws once, a known difference), values typed into and cleared from a spill range after the computation in books with merged areas (Excel's answers through a session, as the engine had them), copies and moves of spilled cells and anchors within a sheet (the struct set `spill-copy`, as the engine had them).
- Reading and writing xlsx, the subpath `xlcalc/xlsx` (`docs/xlsx.md`), with no dependencies. `readXlsx(data)` gives the book, the values Excel stored for its formulas and spilled cells (`cached`) and `issues`, everything the book cannot hold with its sheet and cells. `writeXlsx(book, { result, mode })` writes formulas as dynamic array formulas with their values and spills, or values only, with styles, number formats, widths, heights, merged areas, names and the locale. Checked through Excel 16.113: 143 written books open without a repair window and read back as written, and 91 case books Excel saved read back to the same values.

Changed:

- The harness writes stored answers in the order of their cases, so a run that adds cases changes only their lines. The files were put in that order once, with no answer changed.
- Excel's answers that the tests read move from `docs/` to `test/excel/` (`coverage/`, `cases/`, `probes.json`, `struct/`, `spill/`, `locale/`), and the coverage report's own data to `bench/coverage/data/`; `docs/` keeps the documents. The coverage measurement's three old Excel columns (answers to formulas of the old kind) are removed, they stay in git at `85522cf`. The harness writes only dynamic array formulas; its comments, the coverage report and the measurement's case tags are in English.
- A name whose reference Excel lost (`'S'!#REF!`, `#REF!A1`) stays a name and gives `#REF!` in the formulas that use it, as in Excel; it was reported in `nameIssues` and the formulas gave `#NAME?`.
- Known differences: the 40 cases of a range in a one-value argument (`rangeAsScalar`) are gone, 37 of them now match Excel. The coverage measurement's Excel columns (en-US, ru-RU) and the `text` and `matheng` sets are computed again as dynamic array formulas, as Excel 365 computes them (9 answers changed in each column, 10 in the sets), so no answer of the old kind is left to differ from. New: `VLOOKUP` and `HLOOKUP` lifted by an array key with an array index (`lookupIndexLifted`). `CONCAT` and `TEXTJOIN` longer than 32,767 characters are `#CALC!` (were `#VALUE!`), as in Excel.
- The lost references Excel writes, `'S'!#REF!`, `#REF!A1` (after deleting a sheet) and `#REF!#REF!`, are read as `#REF!` and printed back as written; they were syntax errors (`#ERROR!`). A sheet in `'S'!#REF!` that the book lacks is `unknown_sheet`.
- Functions are organised by category in `src/functions/`, on a shared argument layer (`src/args.ts`) that coerces arguments by kind for every function of a group: numbers, text, booleans, dates with en-US date text, lists by the SUM, AVERAGEA and strict rules, values with their origin, grids, paired arrays, holidays.
- `XLOOKUP` compares the referenced sizes of its ranges: whole columns of sheets with different numbers of rows work (they gave `#VALUE!`).
- Known differences: `lookupIndexLifted` and `nearInteger` are gone, the random check's open differences (`test/excel-open.ts`) are all fixed; new are `lastBits` (`FTEST`'s last bits, seen when magnified) and `omittedArgument` (an argument left out between commas is a syntax error). The sets `criteria`, `findate`, `stat`, `subtotal` and the probes are computed again as dynamic array formulas (one of 5,824 answers changed), so all acceptance is on Excel 365 formulas.

## 0.2.0 (2026-10-05)

Prepared as the first public release, as `formula-engine`, and never published. Version 0.1.0 was an internal release named `@mike/formula`.

Breaking changes:

- Every values `hash` changes: the configuration version, which is part of the hash, is now `formula-engine/1;compiled-js;formulajs@4.6.1;canon15` (it was `app-formula/1;compiled-js;formulajs@4.6.1;canon15`). Values and `inputHash` do not change. The synthetic book (seed 1, 3,400 rows, clock `2026-10-04T12:00:00Z` with seed 7) hashes to `18be025cb4581543…`; with the 0.1.0 version string the same values still hash to the 0.1.0 `9d7414efbda34a63…` (pinned by a test).
- The subpaths `./internal` and `./synth` move to `./unstable/internal` and `./unstable/synth` and are outside semver.
- Input of the wrong shape throws `InputError` instead of being accepted silently or failing with an unrelated error: cells given as bare values, unknown fields in cells, names and options, cell addresses outside `A1:XFD1048576` or not in upper-case A1 style, non-string formulas, malformed dates and clocks.
- Node 20.19 or newer (0.1.0 asked for Node 24). Installing from git no longer builds `dist` (there is no `prepare` script).

Added:

- `compute(book)` and `openSession(book)` work without a clock: the current moment and a random seed, recorded in `result.clock` / `session.clock`; passing that clock again reproduces the result. Calls with a clock work as before.
- `InputError` (a `TypeError` with `path`) with messages that name the place and the problem.
- When `new Function` is blocked (a Content Security Policy without `'unsafe-eval'`), `compute` and sessions throw an `EvalError` that says so.
- The library runs outside Node: SHA-256 in plain JS (digests identical to `node:crypto`), no `node:*` imports, `performance` and `crypto` used only when present. A test bundles it for the browser and runs it in an empty JavaScript context.

Changed:

- Documentation, comments, the contract's JSDoc, messages and test names are in English. `debugGeneratedCode` marks cells in a cycle with `(cycle)`.
- `exports` use the `types` and `default` conditions, so `require()` works on Node versions that can require ES modules.
- The LibreOffice oracle's answers are stored in `test/oracle/libreoffice.json`; `npm test` no longer needs LibreOffice, `npm run test:oracle` recalculates them.
- The README's code examples run as a test. CI for GitLab and GitHub. MIT license; `THIRD_PARTY.md` carries the full license texts of formula.js, jStat and bessel.

## 0.1.0 (2026-10-05)

Internal release as `@mike/formula`, extracted from a private application together with its history. Contract: `compute`, `validateFormula`, `validateName`, `openSession` (`setCells`, `setNames`, `setClock`, `explain`, `result`, `close`), check codes, cell errors with reasons, volatile cells, values and input hashes, external functions through `externals`. Whitelist: SUM, SUMIF, SUMIFS, COUNTIFS, IF, IFERROR, XLOOKUP, ROUND, EOMONTH, DATE, TEXT, TRUE, FALSE, TODAY, NOW, RAND, RANDBETWEEN; ROUND, EOMONTH, DATE and SUM from formulajs 4.6.1. Configuration version `app-formula/1;compiled-js;formulajs@4.6.1;canon15`; built with `tsc` to `dist` with subpaths `.`, `./contract`, `./synth`, `./internal`.
