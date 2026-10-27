# Changelog

## 0.60.0 (2026-10-27)

Paste Special for a grid's Paste Special box, and the function an issue reports as a field of its own. Values and hashes of the synthetic workbook are as in 0.59.1.

Added:

- Paste Special: `copyRange` takes `paste` (`PasteKind`: `all`, `values`, `formulas`, `formats`, `validation`) and `transpose`, in every kind of session, one step of undo each, as Excel for Mac pastes them (49 cases, three sets, the last two held out). `values` pastes each cell's value as calculated; `values` and `formulas` keep the destination's looks, merged areas, notes, hyperlinks, conditional formats and data validation, and are refused with `merged` onto a merged area the copy does not bring; `formats` pastes styles, number formats, merged areas and conditional formats and keeps the contents. With `paste`, a cell of a merged area stands for the area, as Excel's selection does. A transposed copy turns a reference into its source with the cell it points to, its `$` swapped between the axes, and one out of it without a `$` about the source's corner ([Paste Special](docs/structure.md#paste-special)). Not measured in Excel for Windows nor with the Format Painter; `formats` of a table's cells does not paste the table style's look.
- `fn`, the function upper-case, on what reports a function the engine cannot take, so a host counts them without reading messages: on an `XlsxIssue` of code `userFunction`, which `readXlsx` now gives one per function on a sheet or a name, its `cells` the calls (one issue held all of a sheet's calls); on a `validateFormula` rejection, and a cell's issue in `diagnostics` and `Session.getCellIssue`, for `unknown_function` and `unsupported`.

Changed:

- A copy tiled over its destination gives a conditional format the one rectangle its tiles fill, as Excel does, where it gave a rectangle a tile.

## 0.59.1 (2026-10-26)

Fixed:

- `writeXlsx` leaves out the hidden names `openKept` gives the references of the parts the workbook does not hold (`_xlcalc.keep.*`), so a workbook opened by `openKept` can be saved as a new file; they were written into it. `writeKept` is unchanged.

## 0.59.0 (2026-10-26)

What a grid's Home tab needs from the library: Increase and Decrease Decimal, AutoSum, rows that follow their fonts, the workbook's default font, names refused as the workbook would drop them. Values and hashes of the synthetic workbook are as in 0.58.0.

Added:

- `stepDecimals(format, value, step, opts?)`, in `xlcalc/text` and the root: the number format Excel's Increase Decimal and Decrease Decimal give the active cell, by its format, value and column (a General number by the decimals its column shows; a code by a copy of its last decimal in each section; dates, times, fractions, text and accounting's zero stay). Comma Style and Accounting are the built-in formats 43 and 44 ([Number formats](docs/formats.md#increase-decimal-decrease-decimal-comma-style-and-accounting)).
- `Session.getAutoSum(sheet, cell)`, in every kind of session: the range Excel's AutoSum offers and its `SUM` formula, in the workbook's locale too (`AutoSumGuess`): the numbers above, past the blanks and labels right above the cell, else to the left; a `SUM` or `SUBTOTAL` subtotal ends the run, and below a subtotal the subtotals alone are summed; a date is no number ([A viewport for a grid](docs/grid.md)).
- `BookInput.font` (`{ name, size }`): the workbook's default font, its Normal style's. `readXlsx` reads one other than Calibri 11 from the file's first font, `writeXlsx` writes it, `openKept` keeps it and `writeKept` refuses a changed one. It is in `inputHash` when given. The engine still measures widths, fitted heights and `####` as for Calibri 11.
- `EditResult.fitted` (`FittedRows`): the rows a call fitted, with their heights, so a host reads those rows again rather than the sheet.

Changed:

- `patchCells` that changes a style fits the rows of its cells as Excel keeps them fitted, in the same undo step: a larger font makes a row taller, a smaller one lower again; a row of its own height or hidden stays, and over entire columns or the whole sheet only the rows with cells are fitted. A number format a number no longer fits widens its column, as typing does (Excel widens it for Accounting or a decimal more).
- The issue `defaultFont` of `readXlsx` says the font is kept and measured as Calibri 11.

Fixed:

- `setNames`, and the operation `setNames` of `applyChanges`, refuse with an `InputError` at the name a name the workbook would leave out, as `validateNames` finds it (`1abc`, `Q1`, `TRUE`, a duplicate, an unknown scope sheet, a formula that does not parse, a reference without its sheet), where they took it and dropped it; a name the workbook already has as it stands is taken back. `compute` and `createSession` still report such names in `diagnostics`.

## 0.58.0 (2026-10-25)

`writeKept` follows structure operations; arrays larger than the sheet, iterations that never settled and text comparison as Excel. Values and hashes of the synthetic workbook are as in 0.57.0.

Added:

- `writeKept` writes the references of parts outside the model back where an insert, a delete, a moved range or a renamed sheet took their cells: series, categories and titles of charts, the anchors of pictures and shapes by how they move with cells (one whose rows are all deleted goes, with its parts), a pivot table's place and its cache's source (a renamed table too), sparklines, the autofilter and its sort, print areas and titles, page breaks and protected ranges. A moved cell with what only the file holds (rich text runs, phonetic text, a value by `vm`) keeps it, up to 5,000 such cells a sheet; past that the move is reported. Still refused: a range moved over a picture, a pivot table cut, a deleted source of a pivot cache, a sparkline's or autofilter's cells deleted ([xlsx over its original file](docs/xlsx-keep.md)).

Fixed:

- An array larger than the sheet no longer crashes the process with an out-of-memory error: `SEQUENCE` and `RANDARRAY` of more than 1,048,576 rows or 16,384 columns return `#VALUE!` and `VSTACK`, `HSTACK`, `TOCOL` and `TOROW` `#NUM!`, as Excel does (measured), and any array of more than 2^27 values returns `#VALUE!` in its cell. `getValues` and `copyText` read as before.
- A workbook under iterative calculation whose spilled arrays come back to the same state each round finishes its calculation, where it ran 100,000 rounds and threw; calculations that finished before give the same answers.
- Text is compared as Excel compares it: `=` tells a soft hyphen apart, exact `MATCH`, `XMATCH` and `VLOOKUP` tell composed and decomposed spellings apart (é and e with a combining accent), criteria compare text in its composed form.
- `clearCells` of formats over a rectangle in columns with a style of their own, and formatting rows across formatted lines past 1,048,576 cells, are no longer refused; `copyRange` and `moveRange` of whole formatted columns no longer throw `RangeError`.
- A legacy array formula moved or copied over a table's header row is refused, as one over its data already was.
- Arrays waiting for cells their own formula reads back find what they wait for by what their readers read: 1,000 arrays below a chain of 20,000 formulas in 146 ms instead of 2.6 s.

## 0.57.0 (2026-10-25)

Writing an xlsx over its original file, so that what the workbook model does not hold survives an edit; hidden names; fixes in reading and writing xlsx. Values and hashes of workbooks without hidden names are as in 0.56.0.

Added:

- `xlcalc/xlsx/keep`: `openKept(data)` reads a file as `readXlsx` does, and `writeKept(original, { book, result })` writes the edited workbook over the original file instead of rebuilding it. Charts, pivot tables, pictures, shapes, sparklines, the theme, page setup, rich text in unedited cells, macros and parts the model does not know are kept as the file holds them; without edits the original bytes come back. An edit that would leave a reference of such a part out of date is refused with a `KeptRefusal`; writing such references back comes next. [xlsx over its original file](docs/xlsx-keep.md).
- `NameInput.hidden`: a hidden name, as Excel stores it. It works in formulas as any name; `readXlsx` reads `hidden="1"`, `writeXlsx` writes it, every operation keeps it, and `inputHash` counts it.

Changed:

- `readXlsx` reads a hidden name as `hidden: true`; it read it as a visible name with an issue.
- `xlcalc/text` is 45 KB instead of 50 and `xlcalc/worker` 25 KB instead of 30 (minified), as the criteria rules and the en-US locale are no longer built at load for them.

Fixed:

- `readXlsx` no longer throws a `TypeError` on a row with `customFormat` that repeats a cell, and each repeated cell keeps its own content.
- A row height or column width that is empty, blank or not a decimal number sets no size and is reported (`part`), where it read as 0 and hid the row or column.
- Hyperlinks to a sheet whose name `readXlsx` repairs follow the repaired name.
- A formula that read a sheet since deleted is written to a file as `#REF!` (a spill reference as `_xlfn.ANCHORARRAY(#REF!)`), as Excel writes it: `#REF!E2*2`, the form the model holds, made a file Excel for Mac would not open.

## 0.56.0 (2026-10-25)

Speed: reading xlsx, a full calculation, edits of lookup keys and summed amounts, outline and layout commands, merged areas, and deletes, pastes and fills on filtered sheets. Values, texts, workbooks, issues and hashes are the same on every input; the synthetic workbook's values hash and `inputHash` are as in 0.55.0.

Faster:

- `readXlsx` no longer walks every cell of a sheet once per table, merged area, data-validation rule, large array formula or formatted empty cell: a workbook of ten tables over 102,000 cells reads in 110 ms instead of 1.5 s, a sheet of 20,000 merges or 4,000 validation rules in tens of milliseconds, and a thin shared-formula group in time linear in its copies. The synthetic 34,000-row workbook reads 15 % faster.
- `compute` takes the workbook's cells as they are instead of copying each sheet (`createInlineCalc`, whose host runs between the steps, still copies), lists and sorts each sheet's cells once, reads addresses without a regular expression, and hashes with Node's SHA-256 where there is one, with the same digests (a browser and a Node without `node:crypto` keep the plain-JS hasher, and no browser bundle holds a Node module). A cold calculation of the synthetic 34,000-row workbook takes 0.83 s instead of 1.48 s.
- A session finds the formulas a one-cell edit concerns in a bucket of 64 rows, keeps parsed criteria by the value itself, and writes General text and rounds a number without its decimal text where it is already short or whole: editing a lookup key and an amount summed by SUMIFS are faster.
- Outline commands (group, ungroup, show and hide detail, the level buttons) and layout edits patch each axis as one sorted list: 8,000 outline entries in 11 ms instead of 4.7 s, 5,000 wrapped texts typed in 45 ms instead of 2.8 s. AutoFit and typing into a wrapped cell no longer walk every cell; `applyChanges` and `setLayout` skip sheets whose parts they were given unchanged.
- The looks of cells (`getCell`, `getCells`, the viewport, find, a result's changes) and typing read the parsed merged areas the session keeps, instead of parsing every merge of the sheet per call or per cell.
- `deleteRows` and a paste over a filter's visible rows move the cells once instead of once per run of visible rows (a sheet with tables keeps the run-by-run delete): 800 runs in 0.3 s instead of 12 s. A paste or fill reads each copied formula once for all its copies, finds the lines of each tile by halving, and joins pasted validation rules in linear time; a merge across reads its range by row once.

## 0.55.0 (2026-10-23)

Corrections across the engine, sessions, structure operations, xlsx and the pool, with the values that changed checked in Excel. The synthetic workbook's values hash is as in 0.54.2; the `inputHash` of a workbook whose sheets are not in name order changes, since the order of the sheets now counts.

Breaking changes:

- A cell value that is `NaN` or an infinity throws an `InputError` at its path, in `compute`, `createSession` and every cell edit. It used to read as `#NUM!` through references but as a number in `ISNUMBER`, `COUNT` and its own cell.
- `clock.now` must be ISO 8601 with `Z` or an offset, or a date alone (midnight UTC), on a day the calendar has. A date and time without a time zone, free text or `2026-02-30` throws an `InputError` in `compute`, `createSession`, `setClock` and the entry options; such a clock was read in the host's time zone, so `NOW()`, `TODAY()` and the hash moved between machines.
- A number literal from `1E308` on, or one that is not 0 below `1E-309`, makes its formula a `syntax` error, as Excel rejects it typed; it used to calculate as an infinity and print as `Infinity`. A literal up to the smallest normal number is 0 and is written `0` (`=1E-309+1` is stored as `0+1`).
- `dateToSerial` and a cell's `{ date }` throw an `InputError` for a date or time that does not exist (`2025-02-30`, `9999-12-32`, an hour of 25) or text after the date, instead of rolling it into another date. Excel's day 0, 1900-02-29 and a `Z` or offset after the time are taken as before; `parseEntry` is unchanged.
- `externals.values` is checked as cell values: an unknown error code, an array, `undefined`, a `Date` or a `BigInt` throws an `InputError` with its path, where it calculated as 0 or crashed with a `TypeError`. A non-finite number is still `#NUM!`.
- A cell's `meta` that holds a cycle or a `BigInt` throws an `InputError` at its path, in `compute`, `createSession` and the pool, where it overflowed the stack or broke `inputHash`. What JSON turns into `null` or `{}` still passes.
- A rectangle of more than 65,536 cells without an entry of their own whose formatting changed is no longer listed cell by cell in `EditResult.changes`: it comes as a range in `EditResult.repaint`, after `patchCells`, `clearCells`, `applyChanges`, a structure operation (a merge of whole columns, say) and their undo and redo. A host that repaints from `changes` alone must also read again where `repaint` meets its viewport ([Editing a grid](docs/edits.md)).
- `setLayout` rejects a merged area over a table's cells with an `InputError`; it was accepted, and the session's workbook was then rejected by `compute` and `writeXlsx`.
- A legacy array formula of several cells over a table is rejected: by `setCells`, `replaceCells` and `validateEntry` with the workbook check's reason, and by `moveRange` and `copyRange` with the code `table`.
- A copy of a whole table numbers its name past defined names too, and a numbered name Excel would not accept (`Tab3`, a cell address) rejects the copy with `table_name`, as `renameTable` does.
- Rows or columns inserted that would push a table, or a legacy array formula's range, past the sheet's edge return `off_sheet`.
- `mergedArea`, `rectsMeet` and `rectContains` of `xlcalc/a1` check a rectangle as `rangeAddress` does: corners in either order, and a rectangle off the sheet throws an `InputError`.
- `inputHash` counts the order of the sheets when it is not their order by name, since a 3-D reference covers the sheets between its ends in workbook order. Such workbooks, the synthetic one among them, get another `inputHash`; their values hash is unchanged.

Added:

- `TableInput.style` (`TableStyle`: `name`, `firstColumn`, `lastColumn`, `rowStripes`, `columnStripes`): `readXlsx` reads a table's style, operations keep it and `writeXlsx` writes it back. A table without it is written as Excel's default, `TableStyleMedium2` with banded rows, as before; a custom style is reported and read as the default.
- `EntryOptions.inPlace` for `parseEntry`: text edited within the cell, as `entryText` writes it, so a percentage cell's `12.5%` is parsed as 0.125 rather than as the text `12.5%%`.
- `xlcalc/worker/thread`: the pool's thread module, for bundlers that keep `new URL('./thread.js', import.meta.url)` as written (esbuild); [Asynchronous and threaded computation](docs/async.md) gives the entry point and the `spawn` form.
- `xlcalc/contract` exports `Translation`, `TranslateOptions` (the options of `toLocal` and `toInvariant`, exported nowhere before), `CopyContext` and `DateSystemOptions`, so it names every type of the root's signatures.
- `CellTextRefusal` includes `external_ref`, which `setCellText`, `setCellTexts`, `fillCellText`, their previews and `replaceInCells` already returned.

Changed:

- Digit formats write 15 significant digits and then zeros below 1E+21 too, as Excel does: `TEXT(2^60,"0")` is `1152921504606850000`, not the double's binary expansion.
- `FIXED` and `DOLLAR` write the digits of numbers from 1E+21 on (they wrote `1e+21`) and return `#VALUE!` past 255 characters, `DOLLAR` counting the padding of its currency format.
- In ru-RU a color name with a date letter (`[Красный]`, `[Синий]`, `[Черный]`, `[Голубой]`) no longer turns a number section into a date: `TEXT(1234.5,"[Красный]0.00")` is `#VALUE!`, as with `0.00` and the other colors, where it was `12.35`.
- A text section that ends a format of two or three sections takes no numbers: under `h:mm;@` a negative displays `#` (`#VALUE!` in `TEXT`) and under `0;@` it is `-5`, where they showed `0.25` and `5`.
- An elapsed total of 1E+21 and more (`[h]`, `[mm]`, `[ss]`) is written in digits, not as `2.4e+21`; the percent operator returns 0 for a hundredth below the smallest normal number, as `/` does.
- `SUMIF` and `AVERAGEIF` whose sum range, resized to the criteria range, runs past the sheet's edge add the cells that exist instead of returning `#VALUE!`. A formula inside its own resized sum range is a circular reference, as in Excel.
- `UNIQUE` takes spellings Unicode makes equivalent (é in NFC and NFD, `ﬁ` and `fi`) as one row, and keeps a soft hyphen apart; ASCII text is compared as before.
- Operators read number text padded with tabs or no-break spaces as text (`"\t5"+0` is `#VALUE!`), as `VALUE` does; spaces are trimmed as before.
- A data validation rule's formulas are relative to the top-left cell of the rectangle around its ranges, as Excel reads and saves them: a rule over `H3:H6 J1:J3` with `$A1>2` checks `$A3>2` for H3. It was the top-left cell of the first range.
- `2/29` typed in a year that is not a leap year is February 2029 in the format `mmm-yy`, as Excel has it, not 1 February in `d-mmm`.
- A legacy array formula that cannot be prepared (a sheet not in the workbook, an unknown function) displays its error in every cell of its range, rejects values typed into it and gives a dynamic array spilling into it `#SPILL!`, as in Excel; the error was in the top-left cell alone.
- A defined name whose area mixes `$` and relative bounds keeps the relative bounds as offsets when lines are inserted or deleted (inserting row 1 turned `S!$A1` into `S!$A2`), and becomes `#REF!` when the line of its `$` bound is deleted, as in Excel.
- A range typed with its corners reversed is stored in order, each axis on its own with its `$` (`SUM(B2:A1)` is `SUM(A1:B2)`, `$B1:A$2` is `A1:$B$2`), so a formula's text no longer changes at its first structure operation.

Fixed, calculation and spilled arrays:

- Spilled arrays placed again unchanged in every round, or under iterative calculation, settle instead of throwing that they did not settle; an array that gave way to another spill takes its cells once it ranks first.
- Legacy array formulas whose ranges cross each other are rejected; they were accepted and then threw from the calculation.
- `SUMIF` and `AVERAGEIF` depend on their sum range resized to the criteria range, so a session no longer keeps a stale value after an edit of a cell that the written range leaves out.
- A criteria function with an array where a range belongs, inside an operator (`COUNTIF({1},">0")*1`), returns its array of `#VALUE!` instead of aborting `compute`, `createSession` and `setCells` with a `TypeError`.
- `diagnostics` find the source of an error without reading a range over `limits.cells`, which threw a limit error from a workbook that had calculated within it.

Fixed, sessions, undo and stopped calls:

- A merge, an undo or redo of a merge, a preview and `validateEntry` stopped during their recalculation (`soft_timeout`, `aborted`, a limit) leave the session as it was. A stopped merge lost the merged cells' contents with no undo step, and a stopped preview left its trial values in the workbook, so every later call hit the same limit.
- Clear All through a table header: over whole lines the workbook and the calculated state stay together, a cell holding only formatting is cleared, undo and redo keep a valid workbook and give back the rules over the table, and `applyChanges` clears a header cell as `clearCells` does instead of throwing.
- In an `applyChanges` batch, later operations see the formatting earlier ones set, so a batch leaves the workbook the same calls made one by one leave.
- An undone or redone `clearCells` of formats reports `clearCells` in `UndoResult.call`, not `patchCells`.
- A closed synchronous session throws `CalcError` with `closed`, and `setClock` on a live session throws an `InputError`, as the other kinds of session fail.
- `findCells` with `in: 'value'` reads a logical as `TRUE` or `FALSE` and an error by its name (`#DIV/0!`); it matched JavaScript's `true` and an internal code.
- `getCell`, `getCells` and `getViewport` of a synchronous or inline session return copies of a cell's note and link, as a pool does; a host editing them changed the workbook with no undo step or notification.
- A fill over `{ date }` values past the date system writes serial numbers instead of throwing about a `serial` the caller never passed; `limits.writes` bounds a fill's copies together, where a fill wrote up to twice the limit.
- `setCellText`, `setCellTexts` and `fillCellText` wrapping text in more than about 125,000 rows of a sheet no longer throw `RangeError`.

Fixed, structure operations and formatting:

- Formatting alone at the sheet's edge no longer rejects an insert: what is pushed off is dropped, as in Excel, and formatting over a whole column stays over it when rows are inserted. Cells with content, a note, a link or `meta` at the edge still return `off_sheet`.
- A copy or fill that moves a conditional format's first cell rewrites its formulas for the new one, wrapping round the sheet's edge as the rule's evaluation does; the rule kept reading shifted cells.
- `renameSheet` rewrites references to the sheet's own names (`'S'!name`), which turned `#REF!`.
- `renameTable`, `deleteTable`, `convertTable`, a deleted table column and a deleted table sheet rewrite structured references in conditional formats and validation rules, which stopped working; `convertTable` turns a reference to a column the table lacks into `#REF!` instead of its first column.
- A fill series over a table's header renames its columns, where the workbook was then rejected.
- `deleteSheet` and `moveSheet` keep a 3-D reference whose end is not in the workbook instead of throwing `TypeError`.
- A copy of part of a legacy array formula keeps each cell's own note, link, `meta` and formatting, not the top cell's; a copy or fill keeps the spelling of the references it does not change (`$a$1` stays as written).
- A copy from a filtered sheet drops the legacy array range from pasted values and pastes the values of a large spill's cells past its data, which were blank.
- A data validation rule whose own ranges overlap is accepted (by `readXlsx` too, which rejected the file); a list through a defined name sees the blank cells past its sheet's last value.
- A shift that leaves a sheet no rows, columns or merged areas drops `rows: {}`, `columns: {}` and `merges: []`.
- `editStructure` with an op named as an `Object.prototype` member (`constructor`, `__proto__`) and `addTable` with a column without a name throw an `InputError`, not a `TypeError`.

Fixed, functions and number formats:

- `FILTER` checks its result against `limits.cells`; `FREQUENCY` and `AMORDEGRC` count their work, so `timeoutMs` and the signal stop them; `TAKE` and `DROP` of a whole column build no index arrays.
- `LINEST`, `LOGEST`, `TREND` and `GROWTH` fit columns of more than about 124,000 rows, where they returned `#VALUE!`.
- A tab or line break before `(` keeps a function call (`SUM\t(A1)`); `TEXTBEFORE`, `TEXTAFTER` and `TEXTSPLIT` with match mode 1 cut in the right place after an `İ`.
- A number format with more than 100 decimals displays them instead of throwing `RangeError` from `displayValue`, `getCells`, `getViewport` and `setCellText`.
- `entryText` of a date that rounds past 9999-12-31 is the number, not 255 `#`, and of a percentage past the largest number moves the exponent (`1E+309%`).
- ru-RU `explain` writes a value as a formula writes it (a quote doubled, `1E+21`); a locale object not registered has its regional rules checked where it is taken, with an `InputError` naming the field.
- `functionSignature`, `callAt` and `functionDescription` read a function's name as the engine does, so `gauß(` gets no `GAUSS` tip; `nameAt` returns the name being typed after `:`.

Fixed, xlsx:

- A damaged deflate entry rejects `readXlsx` alone instead of ending the Node process with an unhandled rejection; an entry past the end of the file, two entries of one name and bytes that do not match their CRC-32 are rejected as not a readable zip archive.
- XML comments are not read as content, CDATA sections give their content, and a character reference past U+10FFFF rejects its part as not valid XML.
- Reported as issues, where the whole file was rejected or data was lost silently: a sheet name Excel would not take (repaired, and references to it rewritten), a Strict file's date that is not a date, a width or height that is not a size, a threshold of a type the model does not have (the rule left out), a cell, value, formula or `<sheetData>` left open, shared strings left open, a row number that is not a whole number, a style or shared string index out of range, and a number or boolean value that is not one.
- An `extLst` inside `autoFilter`, `sortState` and the like no longer hides Excel 2010's conditional formats or breaks the validations, and SpreadsheetML with a prefix (`<x:conditionalFormatting>`) keeps its conditional formats.
- Frozen panes with a scroll position round-trip without the frozen rows growing at each write and read.
- `writeXlsx` keeps `numberFormat: 0` apart from the code `'0'` (and `14` from `'14'`), which shared one cell format; it reports a formula it leaves out and `#ERROR!` written as `#NULL!`; text cut to an Excel limit does not split a surrogate pair.
- A name scoped to a sheet whose part is missing is left out with the sheet; a cell given twice in a row keeps its last style.

Fixed, the pool:

- A task whose workbook text cannot be written fails alone instead of blocking its queue, and a task rejected before it waits leaves no unhandled rejection.
- A thread that cannot start fails one task with `thread_failed` and the next task tries again; a restore that cannot be sent closes its session with `restore_failed` instead of retrying forever on the event loop.
- `getCells` in a thread keeps its `signal` and `timeoutMs`, and reads without options (`getCell`, `getViewport`, `getUsedRange`, `getEnd`, `getCellIssue`, `getValidation`) are bounded by the pool's `timeoutMs`, as under `createInlineCalc`.
- A pool session keeps its own copy of its options (`externals`, `limits`, `undo`), so changes the host makes to them later no longer reach a restored session; options that cannot be cloned throw an `InputError` at `opts` when the session opens.
- The snapshot taken at `snapshotEvery` goes ahead of calls sent without waiting, so the journal stays within `snapshotEvery`.

Fixed, the package:

- `validateFormula`, `validateLocalFormula`, `parseEntry`, `entryText`, `needsApostrophe` and `xlcalc/entry` work in a tree-shaken bundle without `compute` or `createSession`, where they threw `no arities to check`.

Fixed, speed:

- Wildcards are matched part by part rather than by one backtracking regular expression: `COUNTIF(A1,"*a*a*c")` over 8,000 characters took about a minute in one call `timeoutMs` could not stop. This serves criteria, `MATCH`, `VLOOKUP`, `XLOOKUP`, `SEARCH` and Find and Replace; `XLOOKUP` and `XMATCH` match mode 3 keeps the regular expression as one call.
- Large changes of formatting go by rectangles: `clearCells` of formats over `A2:H1048576` took 43 s and 8.4 million changes and now under a millisecond, `patchCells` over `A:H` 48 s and now 2 ms, a merge of whole columns seconds and now milliseconds. Inserted lines under styled cells, copies, fills and moves of formatted ranges no longer throw `RangeError` or take seconds, and a conditional format over 200,000 ranges no longer breaks structure operations and opening a session.
- A fraction format finds its fraction on the value's continued fraction: 9 placeholders took 0.7 s per value. xlsx parts with elements left open read in linear time, where a 6.6 KB file held a CPU for 44 s.

Fixed, arrays in a cycle through their spill:

- Arrays in a cycle through their own spill ranges settle where they threw that spilled arrays did not settle, and which of two arrays spills follows their formulas and references, not which spilled first: an array that reads a cell within the largest range its formula can take gives way to one that does not, so the answer is the workbook's as Excel opens it from its file and a session gives what `compute(getBook())` gives. Workbooks without such cycles calculate as before. Without iterative calculation, where an array's size depends on the values a cycle holds, the outcome can still depend on the order the cells were entered in ([Spilled arrays](docs/spill.md)).

Documentation:

- `inputHash` is not a cache key on its own for a workbook with external functions, and a workbook read back from xlsx usually gets another `inputHash`: compare `hash` to tell whether its values changed.
- Cell objects passed to a session belong to it and must not be changed in place; `meta` returned on the session's thread is the workbook's own object; `readCells` and `findCells` return `text` under `display: false` when their fields ask for it.
- New known differences: text that Excel's `=`, `MATCH` and `COUNTIF` compare differently (a soft hyphen, NFC and NFD) and a `SUMIF` inside its own resized sum range ([Limitations](docs/limitations.md)); a deleted sheet's names, which Excel writes with the file's name, and arrays in a cycle where Excel keeps the outcome of the order they were typed in ([How the rules were checked](docs/development.md#how-the-rules-were-checked)).
- Smaller corrections: the formula limits `xlcalc/contract` exports, `fromLegacy` from `xlcalc/xlsx` only, `numberFormat` in `displayValue` and `entryWidth`, fitted row heights read with `bestFit`, `createInlineCalc` from `xlcalc` in the 0.54.2 entry, and what a large change of formatting costs.

## 0.54.2 (2026-10-20)

The first public release. The versions below it were internal builds used by our own applications and were not published to npm; their entries are kept as written, and the README describes the package as it is now.

What it includes:

- `compute` and `createSession`: a workbook's formulas are compiled once, to closures or with `execution: 'eval'`, and recalculated incrementally, with deterministic values, a SHA-256 `hash` of them (configuration version `xlcalc/1;canon15`) and an `inputHash` of the input.
- A whitelist of functions, each admitted only after it matched Excel's results on all of its cases, calculated in Excel 16 for Mac and checked again in Excel for Windows; arrays inside formulas and spilled arrays as in Excel 365; tables and structured references; defined names; links to other workbooks through their stored values.
- Sessions: edits, undo and redo, subscribers notified of every change (`subscribe`), structure operations (inserting, deleting and moving rows, columns and sheets; merging and unmerging cells as Excel's Merge Cells, Merge Across and Merge & Center do; the fill handle with the series Excel continues), typed entry parsed as Excel parses it, fill and paste, Find and Replace, outlines, AutoFit, conditional formats, data validation, and the text each cell displays in its number format.
- `liveBook(session)`: a synchronous session's sheets and cells as objects to read and assign (`sheet.cells.A1.value = 100`), batches as one edit, and changes followed by cell, sheet or workbook.
- `xlcalc/xlsx`: reading and writing xlsx without dependencies, with what a workbook cannot keep listed as issues.
- `createInlineCalc` from `xlcalc` and `createCalcPool` from `xlcalc/worker`: calculating without blocking, in slices on the same thread or in worker threads, with deadlines, cancellation and limits.
- Localization: formulas, number formats and messages in ru-RU through `xlcalc/locales/ru-RU`.
- Formula bar tooltips: a function's arguments (`functionSignature`) with argument names of the library's own in English and Russian, its category as Excel's Insert Function dialog lists it (`FUNCTION_CATEGORIES`), and a one-line description (`functionDescription` from `xlcalc/function-text`, in Russian from `xlcalc/function-text/ru-RU`), without the engine.
- `xlcalc/a1` and `xlcalc/text`: addresses, a sheet's outline groups and their commands (`outlineEdit`), the references in formula text, what an arrow key does in it as Excel's Point mode decides (`refAt`), the call and argument the caret is in (`callAt`), and number format categories, without the engine.

The code is that of 0.53.1. The README says that the package is under heavy development and AI-engineered.

## 0.54.1 (2026-10-20)

The internal build of the first public release, without the README's line on its status.

## 0.53.1 (2026-10-20)

Internal release; not published to npm. A sparse sheet's memory, and a comparison with other libraries in the README. The synthetic book's values and hashes are as in 0.53.0.

Fixed:

- Since 0.50.4 a sheet whose values lie far apart kept the empty rows between them: 1,000 values 97 rows apart, each read by a formula, took 8.6 MiB, where 0.50.3 took 2.1. A column's chunk of 1,024 rows with fewer than 20 values now keeps them as pairs of a row and its value and becomes an array by row at the 20th, so the sheet takes 2.2 MiB, and dense columns read and take memory as before.

Documentation:

- [Compared with other libraries](README.md#compared-with-other-libraries): agreement with Excel on the coverage run's formulas next to Univer, IronCalc and HyperFormula, and the speed of opening, edits and structure operations, memory and xlsx next to HyperFormula, IronCalc, SheetJS and ExcelJS.

## 0.53.0 (2026-10-20)

Internal release; not published to npm. The engine's English messages in plain English. The synthetic book's values and hashes are as in 0.52.1.

Changed:

- 166 English messages are rewritten in plain English with the words of the guides (workbook, returns, formatting, color): the en-US catalog of formula checks and entries, the reasons in `diagnostics` and `getCellIssue`, `explain`, the outline commands, the input checks behind `InputError`, `CalcError`, and the issues of `readXlsx` and `writeXlsx`. Codes, fields and behavior are the same, and so are the ends of a reason that name its source (`, from 'S'!A2`). Code that matches a message's text rather than its `code` needs the new text: `is in no group` is now `is not in a group`, `shown already` is `already shown`, and so on.

## 0.52.1 (2026-10-19)

Internal release; not published to npm. Workbooks open faster. The synthetic book's values and hashes are as in 0.52.0.

Changed:

- Formulas that differ only in their references and numbers, such as a formula copied down a column, are parsed, checked and printed once per shape while a workbook is built; each takes the shape's template with its own references and numbers, so formula texts, values and hashes are the same. Opening is faster, with closures and with eval: the chain of 10,000 formulas 1.18 and 1.15 times, the batch workbook 1.37 and 1.31 times, the cumulative sums 1.29 and 1.14 times, the 3,400-row SUMIFS workbook 1.15 and 1.09 times. Structure operations and `fillRange`, which build the workbook again, gain the same way (`fillRange` of 10,000 formulas 1.23 times). Against IronCalc 0.9.0 the chain now opens 1.35 and 1.55 times slower, down from 1.63 and 1.73 ([the comparison](bench/compare/README.md)).

## 0.52.0 (2026-10-19)

Internal release; not published to npm. `liveWorkbook` is `liveBook`, and the API reference in idiomatic English. The synthetic book's values and hashes are as in 0.51.1.

Breaking changes:

- `liveWorkbook` is `liveBook` and its type `LiveWorkbook` is `LiveBook`, so that the API calls the workbook a book throughout (`BookInput`, `getBook`, `{ book }`). There is no alias: the root no longer exports `liveWorkbook`, the type `LiveWorkbook` does not compile, and the errors of a call begin with `liveBook:` ([Live objects](docs/live-objects.md)).

Documentation:

- The comments of the public types and functions, from which [the API reference](docs/api.md) is made, rewritten in idiomatic English, as the guides were in 0.51.1; the code is as it was. `ReadXlsxOptions.limits` gave one part's default bound as 256 MiB; it is 512 MiB, as `limits.entryBytes` says.

## 0.51.1 (2026-10-19)

Internal release; not published to npm. The code in modules by what they do, and the guides in idiomatic English. The synthetic book's values and hashes are as in 0.51.0, and so are the answers, the refusals of wrong input and the public types.

Changed:

- Input checks, the public types (`xlcalc/contract` re-exports ten modules by area), structure's transforms, the formulas xlsx writes and the criteria's selections moved to modules of their own; five cycles of imports are gone. No call answers otherwise: 7,861 refusals of wrong input, the API's 669 types and 1,099 xlsx readings and writings are as in 0.51.0.
- README and every guide in `docs/` rewritten in idiomatic English with the terms of Excel and JavaScript (workbook, returns, displayed text); their code is as it was, but for the quick start's section on locales, which now shows the same cells under each locale's rules and `toLocal` and `toInvariant`.

## 0.51.0 (2026-10-19)

Internal release; not published to npm. `readXlsx` within bounds of memory, the system's date and time tags and the calendar prefixes of number formats as Excel shows them, the fill handle over hidden rows, measured in Excel for Mac and Windows. The synthetic book's values and hashes are as in 0.50.5.

Added:

- `readXlsx(data, { limits })`: `entryBytes`, the bound of one part once inflated, 512 MiB by default; `totalBytes`, of the parts read together, 1 GiB; `ratio`, how many times its size in the file a part over 16 MiB may inflate to, 200 ([Reading and writing xlsx](docs/xlsx.md)).

Fixed:

- `readXlsx` inflated a part whole before it checked its size: a file of 400 KB holding 400 MB of spaces took over 1 GB of memory. It stops a part as soon as it inflates past the size it declares, refuses a part past the bounds before inflating it, and refuses records of the archive over the same bytes.
- `[$-x-sysdate]` and `[$-x-systime]`, which Excel for Mac writes for Long Date and Time, showed `-x` before the date; they show the system's long date and time as `[$-F800]` and `[$-F400]` do. A currency tag with a culture (`[$€-x-euro2]`, `[$$-en-US]`) shows its symbol alone.
- The calendar prefixes at the start of a date code: `B1` the Gregorian calendar with English names, `B2` the Hijri (tabular) calendar with Arabic names, as Excel shows them in en-US and ru-RU (`B2dd/mm/yyyy` of 1 January 2000 is `25/09/1420`). They showed as text.
- `fillRange`'s copies go into hidden rows, on a filtered sheet too, as Excel's fill handle lays them, and a source over hidden rows is copied whole; its series did so already. `copyRange` (copy and Fill Down) fills the visible rows, as before.
- `quoteSheetIfNeeded` with a locale of the wrong shape, a tag such as `'ru-RU'` among them, is an `InputError`; it was ignored, and the name came back without the quotes ru-RU's booleans need.

## 0.50.5 (2026-10-19)

Internal release; not published to npm. The database functions faster on edits, and a quick start. The synthetic book's values and hashes are as in 0.50.4.

Changed:

- The database functions (DSUM, DCOUNT, DCOUNTA, DAVERAGE, DMAX, DMIN, DGET, DPRODUCT, DSTDEV, DSTDEVP, DVAR, DVARP) over a reference select their records through the criteria selection SUMIFS uses, kept while the database's columns hold the same values; records and the order of adding them are as before. 100 DSUM over 34,000 records: an edit of an amount 158 to 0.77 ms, of a product 158 to 2.2 ms, opening 223 to 64 ms.

Documentation:

- [Quick start](docs/quick-start.md): the engine's main uses, from a calculator on a server to a grid, xlsx files, external values and threads, each with an example that runs.
- [Structure operations](docs/structure.md) lists every operation and refusal code in its API.

## 0.50.4 (2026-10-18)

Internal release; not published to npm. A sheet's values in groups of rows. The synthetic book's values and hashes are as in 0.50.3.

Changed:

- A sheet's column of values is kept in groups of 1,024 rows, each made when a cell of it is first written, as IronCalc keeps its rows. A cell far below the others (row 100,000 under 34,000 rows of data) had turned the whole column into a dictionary, read by hashing: an edit of a summed cell on such a book 14.3 ms, now 8.9, as without the far cell. Opening and edits of dense books as before, a chain of formulas 1.15 to 1.22 times faster.

## 0.50.3 (2026-10-18)

Internal release; not published to npm. Conditional formats' code within memory, and SUMIFS and its family faster on edits. The synthetic book's values and hashes are as in 0.50.2.

Changed:

- A conditional format's formula whose references move with its cell has one code for every cell, which reads the cell it runs at. A formula with a name, `@`, a table or a moving reference to another sheet keeps a code per cell, within 8 MB a session, the code read least recently going first. One rule over 200,000 cells, scrolled through by windows of 50 rows, held 205 MiB of code and holds 0.4; its first scroll 307 to 107 ms. A pool's thread of 256 MB ran out of memory after 390 such windows and no longer does.
- SUMIF, SUMIFS, AVERAGEIF, AVERAGEIFS, MAXIFS, MINIFS, COUNTIF and COUNTIFS keep the rows their criteria select while the criteria's ranges hold the same values, keep a criteria range's texts in lower case, and take the index of the criterion that selects fewer rows. On the synthetic book of 34,000 rows, an edit of a summed cell 105 to 9.9 ms, a cell edit at the median 2.65 to 0.48 ms; 1,000 SUMIFS whose first criterion selects half the rows open in 80 ms rather than 790.
- The indexes of MATCH, XLOOKUP and XMATCH count in the cache budget with the array they index, as VLOOKUP's did.

## 0.50.2 (2026-10-18)

Internal release; not published to npm. Structure operations and their undo with less work. The synthetic book's values and hashes are as in 0.50.1.

Changed:

- A structure operation and its undo with `display` make a moved cell's shown text only where its format or room differ on its new place, and no longer read the sheet's new cells again for their looks: on the synthetic book of 34,000 rows, inserting a row 985 to 697 ms and its undo 936 to 669 ms; without `display` 6 to 9 % faster. The answers are as they were, checked against ones recorded on 0.50.0 for eight operations with a spill, an old-kind array formula, held values of a cycle, a merged area, a table, conditional formats, data validation and a hidden row, with every undo and redo, in each kind of session.
- A test that a pool's closed session leaves no record on the calling side and none in its thread.

## 0.50.1 (2026-10-18)

Internal release; not published to npm. A closed session's memory. The synthetic book's values and hashes are as in 0.50.0.

Fixed:

- `createInlineCalc` kept every session it had opened until the calculator itself was closed: about 47 KB for each closed session, more with its undo history, without a bound in a calculator that lives long. A session leaves its calculator once it is closed. A test opens and closes sessions in a fresh process, synchronous, in the same thread and in a pool's thread, and waits for the heap and the session object to be given back.

A session closed at once after it opened can look held for a few milliseconds: the JavaScript engine's background compilation of its formulas keeps it until it finishes (8 MiB for a chain of 10,000 formulas, gone within 10 ms). The library holds nothing then.

## 0.50.0 (2026-10-18)

Internal release; not published to npm. Edits, opening, memory and xlsx writing nearer the fastest of HyperFormula 3.5.0, IronCalc 0.9.0 and SheetJS 0.18.5, and a rectangle's values without an object for each cell ([the comparison](bench/compare/README.md), "Parity on the branch `perf/parity`"). The synthetic book's values and hashes are as in 0.49.1.

Added:

- `Session.getValues(sheet, range, opts?)` and `AsyncSession.getValues`: a rectangle's values as rows, each cell's value as `readCells` gives it, null where a cell holds none; more than `maxCells` cells (1,000,000 by default) is refused (`CalcError` `limit`), never cut ([Live books](docs/live-books.md), "Values of a rectangle"). The synthetic book's data sheet of 506,594 values reads in 10 ms, HyperFormula's arrays of values in 13.

Changed:

- An edit's recalculation keeps a changed cell by a number of its place and orders the formulas to recompute by the book's order: an edit at the root of a chain of 10,000 formulas 4.9 to 2.2 ms, of 50,000 formulas that read one cell 36 to 16 ms (IronCalc 1.9 and 15.4), 2,000 inputs at once 3.5 to 2.6 ms.
- Opening finds a formula by its cell's number and an open book holds less: references and successors of their own length, a range's cache key made when a cache asks for it. The chain retains 13 to 7.2 MiB, the 50,000 formulas 63 to 37 MiB; the synthetic book of 34,000 rows computes cold 11 to 15 % faster.
- In closures, an operator over two cells or a cell and a constant is one function, and a call of up to three arguments takes them without an array.
- Writing a sheet's xlsx keeps its cells' XML by row and column, the bytes the same: 34,000 rows 751 to 680 ms (SheetJS uncompressed 363, compressed 463).
- `npm test` and `npm run test:dist` run the tests that time themselves after the others, one file at a time.

## 0.49.1 (2026-10-17)

Internal release; not published to npm. The Sort command's order of errors, measured in Excel. The synthetic book's values and hashes are as in 0.49.0.

Fixed:

- `compareValues` and `valueComparator` take all errors as equal again, as Excel's Sort command does: measured in Excel 16.114 for Mac and 16.0.20430 for Windows, errors keep their order ascending and descending. 0.49.0 had given them the order of the function SORT, by ERROR.TYPE's number, which SORT and SORTBY keep.

## 0.49.0 (2026-10-17)

Internal release; not published to npm. Excel's answers in the cells of IronCalc's xlsx tests where the library answered otherwise: 459 of 461 cells now answer as Excel for Mac, each rule checked by a set of cases in Excel for Mac and Windows (25 sets, 570 cases). The synthetic book's values and hashes are as in 0.48.0.

Added:

- `XMATCH` and `XLOOKUP` take match mode 3, regular expressions, as Excel for Mac 16.114 does. The patterns are read by JavaScript's syntax, where Excel reads PCRE2's ([Limitations](docs/limitations.md)); Excel for Windows 16.0.20430 has no such mode and answers `#VALUE!`.

Changed:

- `compareValues` and `valueComparator` order errors by their `ERROR.TYPE` number, as `SORT` now does: `#DIV/0!` sorts before `#N/A`, where before errors were equal and kept their order.
- A date's format code reads a dot and one to three zeros anywhere after its first date code as the decimals of a second, with `m` or `mm` before them minutes, and any other digit placeholder, `%` or `E` makes the code invalid, as Microsoft's grammar of format codes has it: `TEXT(x,"mm.00")` of 12:17:42.72 is `17.72`, `TEXT(x,"mm###")` is `#VALUE!`.
- en-US number text reads the euro sign before or after the number, and a fraction after a currency sign or in parentheses (`VALUE("€1,234.5")` is 1234.5). Typed entry reads them as before.
- The README counts the cases checked in Excel by locale, and the guides' titles use the accepted terms (Legacy array formulas and implicit intersection, Cell entry, External references and others).

Fixed, as Excel answers:

- `ACCRINT` steps its coupon dates one from the other, and with `calc_method` FALSE counts whole periods as 0.
- Criteria (`COUNTIF`, `COUNTIFS`, `SUMIF` and the others) tell a formula's empty text from a blank: `"<>"` counts it, `"="` takes blanks alone.
- `XLOOKUP` and `XMATCH` search by halves as Excel does, at the middle with blanks greater than any key, and find the next value across types.
- The database functions name a column by its header in any case, a boolean header too.
- `FLOOR` of 0 is 0, any mode but 0 turns `FLOOR.MATH` and `CEILING.MATH`, and `ROUND` reads 15 digits.
- `TEXTAFTER` and `TEXTBEFORE` round the instance down and answer `#VALUE!` for a delimiter longer than the text once an instance is given; `TEXTSPLIT` needs a delimiter.
- `DAYS360` takes the days of moments, `WEEKDAY` reads both arguments before its range, `WORKDAY` and `WORKDAY.INTL` round their days down, and time text refuses five-digit minutes and seconds.
- `FV` and `PV` raise by Excel's power (a rate of -1 over 0 periods is `#NUM!`), `IRR` with a guess of -1 or less is `#VALUE!`, `MIRR` without inflows -1, `XNPV`'s first blank `#VALUE!` and later ones `#NUM!`.
- `GCD` and `LCM` of blanks alone are `#VALUE!`, `CSCH` and `IMCSCH` keep a tiny argument's digits, `POISSON.DIST` checks x before truncating, `PROB` passes over text, and `LOGNORM.DIST` and `NORM.S.INV` give Excel's last digit next to the median.
- A function reads a position past a shorter array as `#N/A` in its turn, `SORT` orders errors by `ERROR.TYPE`, `INDEX` takes an array's row, `FILTER` takes `if_empty` by value, and number text from 1E+308 up is not a number.

## 0.48.0 (2026-10-17)

Internal release; not published to npm. Faster reading and batches of edits, measured against other libraries. The synthetic book's values and hashes are as in 0.47.0.

Changed:

- `readCells` reads a cell's value, formula, format, style, note and link from its value and input when `text` is not asked, without building the cell's whole result, and goes by the index of the sheet's cells alone where no spill or look meets the range; `getCell` reads its one cell without the work of a rectangle. On the synthetic book's data sheet of 506,594 values: the sheet by pages 358 to 157 ms, 10,000 `getCell` 19 to 8.7 ms, a window of 1,600 cells 0.41 to 0.21 ms.
- A batch of 2,000 edits in one `setCells` is 1.4 times faster: a cell's fields are copied rather than deleted from a copy, and content is compared without JSON.
- `bench/compare` compares opening, edits, structure operations, reading, xlsx and retained memory with HyperFormula 3.5.0, IronCalc 0.9.0, ExcelJS 4.4.0 and SheetJS 0.18.5 ([its README](bench/compare/README.md)).

Fixed:

- `readCells` and `findCells` with the field `text` threw on the cells of a large spill's unstored rest, which they list; they give their text now. `getCell` lists none of them, as before.

## 0.47.0 (2026-10-16)

Internal release; not published to npm. `readCells` in order and faster, and the docs' examples run. The synthetic book's values and hashes are as in 0.46.0.

Changed:

- `readCells` gives each range's cells by rows (`A1`, `B1`, `A2`), as `findCells` gives them, whatever order the book's cells were given in; before, a range's input cells came in the order of their keys, then the cells of a look alone, then spilled cells. `readCells` and `summarizeRanges` find the cells of their ranges by the index of a sheet's cells by place that Find keeps, built at the first read of a sheet and kept up to date by every write, instead of going over every cell of the sheet: on the synthetic book's data sheet of 506,594 cells a window of 1,600 cells takes 1.2 ms (64 ms before), the sheet by pages of 10,000 0.47 s (5.3 s), the first read of the sheet 97 ms (64 ms); edits as fast as before.
- A cell of a session's book is written one way alone (`putInput`), which notes each write for the index; the values a cycle holds go that way too.
- The examples of the docs print what their comments say and are run by the tests: new ones in [Spilling](docs/spill.md), [Arrays inside a formula](docs/arrays-incell.md) and [References](docs/references.md), and the examples of calls in eleven other documents.

## 0.46.0 (2026-10-16)

Internal release; not published to npm. A grid's outline commands and sorting without the engine. The synthetic book's values and hashes are as in 0.45.1.

Added:

- In `xlcalc/a1` and the root: `outlineEdit(sheet, op)`, Excel's outline commands (Group, Ungroup, Show and Hide Detail, the level buttons) on a sheet's rows and columns alone, as `editStructure` does them in a book, with its refusals in English. The commands' reading and their messages are the engine's own, in modules without it ([Outlines](docs/outline.md)).
- In `xlcalc/text` and the root: `valueComparator(opts?)`, `compareValues` with its options checked once, for sorting many values (34,000 mixed values sorted in about the time of `sortOrder` itself, half of `compareValues`'s). `compareValues` and `valueComparator` take `locale`, whose order of text they sort by; ru-RU's is en-US's, as Excel for Windows sorts in every locale (the sets `collation` and `collation-more`).

Changed:

- `xlcalc/a1` is 16.6 KB, its budget 17,000 bytes.

## 0.45.1 (2026-10-16)

Internal release; not published to npm. The functions' arguments once in each bundle. The synthetic book's values and hashes are as in 0.45.0.

Fixed:

- 0.45.0 put the table of the functions' declared arguments, which `xlcalc/entry` checks a formula's calls by, into the root as well, beside the engine's own specs: `compute` took 11 KB more (291.6 KB, 280.4 in 0.43.0). The engine now checks by its specs, the table is `xlcalc/entry`'s alone (`compute` 280.7 KB), and a check of calls with neither throws rather than passing over the counts.

## 0.45.0 (2026-10-16)

Internal release; not published to npm. What a grid needs without the engine. The synthetic book's values and hashes are as in 0.44.2.

Added:

- `xlcalc/entry`: `parseEntry`, `needsApostrophe` and `entryText` without the engine (97 KB minified, 36 KB gzip), a locale given as an object. The check of a formula's calls takes each function's declared arguments (`function-arity.ts`, made from the definitions and kept the same by a test), not its implementation ([Typed entry](docs/entry.md)).
- In `xlcalc/text` and the root: `nameAt(text, caret?, { locale })`, the name of a function being typed for Formula AutoComplete; `FUNCTION_CATEGORIES`; `columnWidthToChars` and `columnWidthFromChars`, the width Excel for Mac's Column Width box shows and takes, checked against every column the Mac's fit sets read; `DEFAULT_ROW_HEIGHT`, `MAX_ROW_HEIGHT`, `MAX_OUTLINE`; `findMatcher` and `findReplacer`, Find and Replace over a host's texts as `findCells` and `replaceInCells` match; `compareValues`, the order SORT sorts mixed values in.
- In `xlcalc/a1` and the root: `mergedArea(rect, merges)`, the rectangle a merge takes in.
- `summarizeRanges` takes `rows`, the rows a host's filter shows, as `readCells` does; an asynchronous session's `getCells` takes `signal` and `timeoutMs`.

Changed:

- `xlcalc/text` is 51 KB (19 KB gzip), its budget 55 KB; `xlcalc/a1` 11.8 KB, its budget 12,000 bytes.

## 0.44.2 (2026-10-16)

Internal release; not published to npm. A session lets go of the state it built before. The synthetic book's values and hashes are as in 0.44.1.

Fixed:

- Since 0.27.0 every operation that builds a session's book again (a table's column renamed by its header, a structure operation, new names) kept the state built before it, and so each one before: about 84 MB more at each over a table of 34,000 rows and 16 columns, never given back, so that a pool of `maxHeapMb: 384` ran `out_of_memory` at the third rename. The new state now holds nothing of the old one, and the heap stays where it was after the session opened (`test/session-rebuild-memory.test.ts`, in a fresh process).

## 0.44.1 (2026-10-16)

Internal release; not published to npm. Point mode by the rule of a formula closed on Return. The synthetic book's values and hashes are as in 0.44.0.

Fixed:

- `refAt` enters a formula left open when Return would close it without a word, and does nothing when Return would bring Excel's window: `=SUM(1)+MAX(2` and `=ABS(-(1)` do nothing, as Excel does, where `refAt` entered them ([Point mode](docs/entry.md#point-mode)). Two known differences of Point mode are left, in en-US: `=ROUND(1.5` and `=ROUND(A1*2`, too few arguments.

Changed:

- The brackets left open in typed text are read in one place for `callAt`, `refAt` and a formula closed on Return, and a function's English name is found in one place for `callAt`, `functionSignature` and the correction's check; no change of their answers.

## 0.44.0 (2026-10-16)

Internal release; not published to npm. A formula left open closed as Excel closes it on Return. The synthetic book's values and hashes are as in 0.43.0.

Added:

- A typed formula left open at its end: `parseEntry`, `setCellText`, `setCellTexts`, `fillCellText` and their previews, typed over a cell or edited in place, close it without a word when it begins with a function, a sign or a space before it or not, whose bracket is the only one left open, with no `)` typed (`=SUM(1,2` is `=SUM(1,2)`, `=-SUM(C1` is `=-SUM(C1)`). Any other formula left open is refused as before, now with `correction`, the formula Excel's window offers to correct it to: its brackets, quotes and array braces closed and the spaces after `=` dropped (`=(1+2` offers `=(1+2)`, `=SUM(MAX(1` offers `=SUM(MAX(1))`). Excel offers none, and so the engine, when the formula closed would not read, calls a name that is no function in the locale, or leaves an array constant open inside a function's brackets or after an operator. Pasted text is not closed ([A formula left open](docs/entry.md#a-formula-left-open)).
- Contract: `EntryResult`'s refusal, `CellTextRefused` (the refusal of `setCellText` and `previewCellText`) and `CellTextsRefusal` (of `setCellTexts`, `fillCellText` and their previews) carry `correction`.

Changed:

- `=SUM(A1,` and the like, refused before, are taken as Excel takes them (`SUM(A1,)`).

Measured: on the Mac's keyboard in Excel 16.113 and 16.114 for Mac, en-US and ru-RU, three sets each predicted before Excel ran it, 165 cases (69 of 79, 46 of 51 and 23 of 35 predicted); the engine as Excel on all but five known differences, none of them about brackets. Not measured: Excel for Windows.

## 0.43.0 (2026-10-15)

Internal release; not published to npm. The texts of the tips apart from the engine. The synthetic book's values and hashes are as in 0.42.0.

Added:

- `xlcalc/function-text`: `functionDescription(fn, locale?)`, what a function does in a line, by its English name (`fn` of `functionSignature` and `callAt`); the English lines, 9.4 KB gzip.
- `xlcalc/function-text/ru-RU`: `ruRUFunctionText`, the Russian names of arguments and categories and the Russian lines, 12.8 KB gzip. A locale with the tips' words is `{ ...ruRU, ...ruRUFunctionText }`; `functionSignature` and `functionDescription` take it ([Addresses and formula text](docs/addresses.md)).

Changed:

- Breaking: `FunctionSignature` gives `fn`, the function's English name, and no longer `description`; the line is `functionDescription(sig.fn, locale)`.
- `xlcalc/locales/ru-RU` computes and reads formula text without the tips' words: 15.9 KB gzip, was 28.4. `functionSignature` takes 9.1 KB of a bundle, was 15.1; `compute` and the sessions are as they were.

## 0.42.0 (2026-10-15)

Internal release; not published to npm. Live objects of a session's sheets and cells; a function's category and what it does. The synthetic book's values and hashes are as in 0.41.0.

Added:

- `liveWorkbook(session)`: the sheets and cells of a synchronous session as objects. A cell reads through `getCell`; `value` and `formula` are assigned as `setCells` (`sheet.cells.A1.value = 100`), `value = null` clears as `clearCells`, `numberFormat` and `style` are assigned as `patchCells`, `cell.type('5%')` is `setCellText`; `batch(fn)` makes the assignments inside it one `applyChanges`, one revision and one step of undo; `watch` on a cell, a sheet or the book is told of its changes, through one subscription to the session. A cell object stays on its address, a sheet object follows its sheet's name. An example with Vue's `customRef` is in the guide ([Live objects](docs/live-objects.md)).
- `FunctionSignature.category`, the category Excel's Insert Function puts the function in, and `categoryName` in the locale's words; `FUNCTION_CATEGORIES` in Excel's order, the old names under `compatibility`.
- `FunctionSignature.description`, what the function does in a line of the library's own words in en-US and ru-RU, for all 470 functions (moved to `xlcalc/function-text` in 0.43.0).

Changed:

- The documentation: a short README with a map of `docs/`; the reference `docs/api.md` made from the package's types and their comments, a test failing on a stale reference or an export without a comment; the README's former sections as guides in `docs/`, whose examples run as tests.

## 0.41.0 (2026-10-15)

Internal release; not published to npm. The call a caret stands in. The synthetic book's values and hashes are as in 0.40.0.

Added:

- `callAt(text, caret?, { locale })` in `xlcalc/text` and the root: the innermost call whose brackets hold the caret, for the tip of its arguments: `name` as typed, `fn` in English (for `functionSignature`), `index` of the argument from 0, `open` and `close` of its brackets (`close` null while it is open), or null outside any call. Strings, quoted sheets, tables' brackets, array constants and grouping brackets give no call; a separator inside them belongs to no argument. The library's own rule, not measured in Excel ([Addresses and formula text](docs/addresses.md)).

## 0.40.0 (2026-10-15)

Internal release; not published to npm. The tip of a function's arguments. The synthetic book's values and hashes are as in 0.39.0.

Added:

- `functionSignature(name, locale?)` in the root: `{ name, args: [{ name, optional }], repeat, text }` for each of the 470 functions, `text` as Excel's tip shows it (`SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`, in ru-RU `СУММ(число1;[число2];...)`), the function's name in English or in the locale in any case, null outside the whitelist. The arguments' names are the library's own in the manner of Excel's, its abbreviations spelled out (`pr` is `price`, `npery` `periods_per_year`); the Russian words are in the ru-RU locale. Which arguments are optional and which repeat is checked by a test against each function's measured arguments. Not in `xlcalc/text`, which its names would take from 45 KB past its budget of 50; a bundle that does not call it holds none of them ([Addresses and formula text](docs/addresses.md)).

## 0.39.0 (2026-10-14)

Internal release; not published to npm. The fill handle. The synthetic book's values and hashes are as in 0.38.0.

Added:

- `editStructure` `{ op: 'fillRange', sheet, range, to, type? }`, the fill handle as Excel's `Range.AutoFill`: `to` holds the source and goes on from it down, up, right or left; `type` is `default` (a drag), `series` (with Control: a single number grows too) or `copy`. The source is copied first, as `copyRange` copies (formulas, styles, formats; a part tile takes the source's nearer part), then each column (or row) goes on as a series: numbers by a step or a trend, dates by days, a step or months (one day of the month or the months' ends), times by hours or a step, text with a number at its start (with a space) or end, quarters, en-US ordinals, the names of days and months by their lists. Empty cells do not break a series; text, booleans and formulas do. Dates and text without one step are copied; a single number is copied unless the source holds numbers alone. One step of undo, in every kind of session ([The fill handle](docs/structure.md#the-fill-handle)).

Measured: in Excel for Windows, three sets each predicted before Excel ran it (`fill` 38 of 46, `fill-held` 37 of 45, `fill-third-held` 28 of 31), the engine as Excel on all 122 cases after the rules were corrected. Not measured: Excel for Mac, ru-RU's lists, zeros before a number at the end of text (`Item001`), a time beside a single number. Known difference: on a filtered sheet the fill writes the visible rows alone, as `copyRange` does; Excel fills the hidden ones too.

## 0.38.0 (2026-10-14)

Internal release; not published to npm. Point mode with the caret anywhere. The synthetic book's values and hashes are as in 0.37.0.

Changed:

- `refAt` with the caret at 0 of a formula answers `insert` with `equals`: Excel puts an `=` of its own and the reference before the text (`=A1` becomes `=D3=A1`). With the caret inside the formula the whole formula decides between `commit` and `none`. A function's bracket left open is closed as Return closes it, a grouping bracket is not ([Point mode](docs/entry.md#point-mode)).

Fixed:

- A structure operation that builds the book again answers `styles` when it grows the style table (a `copyRange` over a row with a style, a merge with centering that moves a value); before, it said nothing and a host's table fell behind.

Measured: on the Mac's keyboard in Excel 16 for Mac, en-US and ru-RU, three sets each predicted before Excel ran it (27 of 48, 32 of 38, 20 of 26). Six formulas with a function's bracket left open, where Excel does nothing, stay known differences.

## 0.37.0 (2026-10-14)

Internal release; not published to npm. Merging cells as Excel's commands do. The synthetic book's values and hashes are as in 0.36.0.

Added:

- `editStructure` `{ op: 'mergeCells', sheet, range, across?, center? }` and `{ op: 'unmergeCells', sheet, range }`: Excel's Merge Cells, Merge Across, Merge & Center and Unmerge Cells, one step of undo, in every kind of session. An empty top left cell takes the area's first value, which moves as a cut and paste moves it, the references and the look following; references to the empty top left cell itself become `#REF!`. A merge that moves no value is written as an edit of cells without building the book again (0.17 s instead of 1.5 s over 3×3 cells of the synthetic book of 34,000 rows) ([Merging cells](docs/structure.md#merging-cells)).

Changed:

- Repeated code of `src` moved into shared helpers; no change of behaviour.

Measured: in Excel for Windows through the object model (`Range.Merge`, `UnMerge`, alerts off), four sets each predicted before Excel ran it, 61 cases (20 of 28, 10 of 18, 12 of 12, 3 of 3), the engine as Excel on all of them.

## 0.36.0 (2026-10-14)

Internal release; not published to npm. What an arrow does in formula text. The synthetic book's values and hashes are as in 0.35.0.

Added:

- `refAt(text, caret?, { locale })` in `xlcalc/text` and the root: what an arrow does while a formula is typed over a cell, as Excel's Point mode decides it. `insert` (a reference goes in at the caret: after the formula's first character, an operator, `(`, the locale's separator, `:`, `!`, a point, a space), `commit` (the text is entered, when the formula reads in the locale), `none` (it does not read, or the caret is inside a string) ([Point mode](docs/entry.md#point-mode)). `xlcalc/text` takes the formula parser with it: 45 KB, 17 KB gzip.

Measured: on the Mac's keyboard in Excel 16 for Mac, en-US and ru-RU, three sets each predicted before Excel ran it (40, 27 and 20 cases; the third 19 of 20 in each locale, the one left a sheet the book does not have). Not measured: the mouse, a sheet's tab, Excel for Windows.

## 0.35.0 (2026-10-14)

Internal release; not published to npm. Each change of the book told to subscribers. The synthetic book's values and hashes are as in 0.34.0.

Added:

- `Session.subscribe(listener)` in every kind of session: after each call that changed the book (an edit, typed text, a paste, a structure operation, layout, styles, `undo`, `redo`) the listeners get a `SessionChange`, the call's `EditResult` with its `method`, before the call returns. A preview, a read, a refusal and a failed call are not told; a change a listener makes reaches every listener after the one it answers, in the order of revisions; a listener's exception does not break the call and is thrown again in a microtask. `subscribe` is synchronous in the asynchronous sessions too ([Following changes](docs/sessions.md#following-changes)).

Measured: a paste of text with quotes in ru-RU in Excel for Windows (12 cases), as en-US pastes it: `parseClipboard` has one rule for every locale.

## 0.34.0 (2026-10-13)

Internal release; not published to npm. A selection's clipboard text from the session; a width or height set by hand leaves AutoFit. The synthetic book's values and hashes are as in 0.33.0.

Added:

- `Session.copyText(sheet, range, opts?)`, also in inline and pooled sessions: a selection's plain-text copy prepared from computed values and effective formats without column clipping. Hidden rows and columns are skipped, merges and spills are handled, and the book's locale and date system are respected. Options: `trailingNewline` (false for Mac), `maxCells` (20,000 by default), `maxChars` (16 Mi characters); oversized copies are refused rather than truncated. Asynchronous calls also take `signal` and `timeoutMs`. Before, a host that fed screen text to `formatClipboard` copied narrow dates as `##` and rounded narrow numbers ([Copying plain text](docs/entry.md#copying-plain-text)).

Fixed:

- A height assigned through `setLayout` without `bestFit` now clears the row's old fitted mark. A manually set height stays on the next text entry, including when assigned the same height; undo and redo restore the mark with the height. `height: null` returns to the default, and an explicitly given `bestFit` is honored. Before, a height set after AutoFit kept `bestFit: true` and wrapped text changed it again ([Fitting rows](docs/formats.md#fitting-rows)).
- A column width set through `setLayout` without `bestFit` clears the column's fitted mark, as a row height does: a number typed later no longer widens a column whose width was set by hand after AutoFit. An explicit `bestFit: true` keeps it fitted ([Fitting columns](docs/formats.md#fitting-columns)).

## 0.33.0 (2026-10-13)

Internal release; not published to npm. Row heights as Excel fits them, wrapped text on entry, the clipboard's plain text, the saved active sheet and view. The synthetic book's values and hashes are as in 0.32.1.

Added:

- `setLayout` takes `{ sheet, rows: '2:5', fit: true }`: the heights Excel for Mac gives rows by AutoFit (a double click below a row's number), one revision and one step of undo, in every kind of session; `fitRows(book, result, sheet, rows)` gives them for a computed book without a session. A row is as high as the tallest cell's text above the common baseline plus the deepest below it, by the font of the cell, else of its row's or column's style, wrapped text by its column's width ([Fitting rows](docs/formats.md#fitting-rows)). `RowInput.bestFit` marks a height Excel fitted rather than one set by hand.
- Text with a line break typed through `setCellText`, `setCellTexts` or `fillCellText` gets a style with `alignment.wrap` (an equal style found in the table, else added) and its row fitted in the same step of undo, unless the row's height was set by hand or the row is hidden; `pasteText` sets the wrap and leaves the height. The result carries `styles` when the table grew and `layout` when a height changed ([A line break](docs/entry.md#a-line-break)).
- `formatClipboard(rows, opts?)` in `xlcalc/text` and the root: a rectangle of shown texts as Excel puts plain text on the clipboard, fields at tabs, rows at CRLF, a field with a tab or line break in quotes with its quotes doubled, a CRLF at the end unless `trailingNewline: false` (Excel for Mac leaves it out) ([Copying plain text](docs/entry.md#copying-plain-text)).
- `columnWidthToPoints(width?)` and `columnWidthFromPoints(points)` in `xlcalc/text` and the root: a column's xlsx width in points of the library's geometry (Excel for Mac, a column without a width 53 points) and back, for a grid that draws columns as the library lays out `####` and overflow; CSS pixels are points × 4/3 ([Column geometry for a grid](docs/formats.md#column-geometry-for-a-grid)).
- `BookInput.activeSheet` and `SheetInput.view` (`activeCell`, `selection`, `topLeft`, `zoom`): `readXlsx` reads the active sheet and each sheet's first view from the file, `writeXlsx` writes them back, and structure operations keep `activeSheet` on a visible sheet. Neither enters computation or `inputHash`. Only the active sheet's tab is written selected: a group of selected tabs is read as the active sheet alone ([The saved active sheet and view](docs/xlsx.md#the-saved-active-sheet-and-view)).
- The asynchronous `explain` takes `signal` and `timeoutMs`, in `createInlineCalc` and the pool, and stops a waiting or running call with `CalcError` `aborted`; `ExplainNode.location` gives a node's sheet and A1 cell or range, for a panel that goes to it.

Changed:

- `parseClipboard` and `pasteText` read a field begun by a quote to its closing quote, tabs and line breaks inside it, a doubled quote as one, text after the closing quote joined to the field; a quote elsewhere is a character, as both Excels paste text. Before, quotes were characters everywhere.
- `readXlsx` reads a row height without `customHeight` as `{ height, bestFit: true }` (before, it left it out); 409.6, which Excel for Mac writes as its highest, reads as 409.5. A book read from a file now has `activeSheet` and sheets' `view` where the file has them.

Measured: row heights in Excel 16.113 for Mac and Excel for Windows (the step `rows`, object model and keyboard, en-US and ru-RU), the engine as Excel for Mac on 709 of 723 rows (the rest in `test/rows-known.ts`: fonts other than Calibri, a formula typed over lines); the clipboard's text copied in both Excels in both locales; column widths through both object models. Known differences: Windows fits rows in pixels (heights differ from Mac, as widths do); heights are not refitted when a book opens or a font changes.

## 0.32.1 (2026-10-13)

Internal release; not published to npm. The synthetic book's values and hashes are as in 0.32.0.

Fixed:

- The refusal of `limits.writes` is in the book's language, as a person's edit brings it: in a ru-RU book `Правка записывает слишком много ячеек: 1048576 при пределе 100000`. The other limits' messages, which a host's settings bring, stay in English.

## 0.32.0 (2026-10-13)

Internal release; not published to npm. A bound on the cells one edit writes. The synthetic book's values and hashes are as in 0.31.0.

Added:

- `limits.writes` (`CalcLimits`, for `createSession`, the calculators and a call's options, field by field as the other limits): the cells with content (a value or a formula) one edit may write, checked before anything is written or built; past it the call throws `CalcError` with code `limit`, `limit: 'writes'`, `max` and `size`, and the book, revision and history stay. A copy counts its source's cells with content times the tiles of its target, a move its source's, `fillCellText` its range, `pasteText` the clipboard's rectangle, `replaceInCells` the cells found, `setCells`, `replaceCells`, `setCellTexts` and `applyChanges` their edits; previews alike. Inserting and deleting lines and other structure operations, clearing, undo and redo are not bounded. None by default, as Excel writes a copy onto a whole column: with `writes: 100000` such a copy is refused in under a millisecond, where it took 2.5 s and 205 MB for a value and 6.7 s and 1.4 GB for a formula ([docs/async.md](docs/async.md)).

Changed:

- A copy from or onto a filtered sheet finds its runs of visible rows by the spans of hidden rows, not row by row: 1,000 hidden spans over a column take 1 ms instead of 1 s.

Measured: every bounded edit refused one under its count and taken at it, in every kind of session and both execution modes, with the book, revision, history and hash after a refusal as before; the runs of visible rows against the former row by row reading on 3,000 random sheets.

## 0.31.0 (2026-10-13)

Internal release; not published to npm. Where Ctrl+arrow lands. The synthetic book's values and hashes are as in 0.30.0.

Added:

- `session.getEnd(sheet, cell, direction)` (`'up' | 'down' | 'left' | 'right'`, type `EndDirection`): the cell Excel's Ctrl+arrow (`Range.End`) lands on, in every kind of session. A cell with a value or a formula is not empty, whatever it shows (`=""`, a zero, a space, an error, spilled cells); a cell with a look, a note or a link alone is empty, in the `looks` layer too. From a cell whose next one is not empty it goes to the end of the block, else to the next cell that is not empty or the sheet's edge; hidden rows and columns are passed over, never landed on; a merged area is its cells. It reads the cells of one line, under 1 ms on the synthetic book of 34,000 rows ([docs/live-books.md](docs/live-books.md)).

Measured: Excel for Windows through COM by a new step of the harness (`end`), three sets of 190 cases (en-US, and ru-RU for 4), each predicted before Excel ran it; all came out as predicted.

## 0.30.0 (2026-10-13)

Internal release; not published to npm. Searches that cost only their matching, and Excel's Replace. The synthetic book's values and hashes are as in 0.29.0.

Added:

- `rows` in the options of `findCells` and `readCells`, as `getViewport` takes it: only those rows are read and counted (a host's own filter of rows shown).
- `after` in the options of `findCells` (a `CellRef`): the cells come from the place after it in Find's order, round to the start and to it last, as Excel's Find Next goes on from the active cell.
- `editText` as a field of `readCells` and `findCells`: the text a cell opens with for editing, as a window gives it.
- `replaceInCells(ranges, query, replacement, opts)`: Excel's Replace (Look in Formulas, the options of `findCells` but `in` and `byColumns`), in one revision and one step of the undo, the answer with `replaced` and the cells whose new text Excel does not take (`refused`); the rules as measured: `*` in the query takes as little as it can but the last one, a cell found is entered again even when its text is unchanged, a percent has no percent editor, a formula that no longer reads leaves its cell as it was while the others are written, an old-kind array is replaced whole ([docs/live-books.md](docs/live-books.md)).

Changed:

- A session keeps the texts it searches by, by sheet and kind, built at a sheet's first search and kept by edits cell by cell (a change of a sheet's layout or of the styles starts them again): on 510,000 cells a search again or narrowed takes 20 to 50 ms instead of 0.4 s, and about 28 MB is held. A page after a cursor goes on from its place without matching again what came before, and several ranges read their sheet once: the component's filtered search of 1,726 rows went from 5.3 s to 0.1 s, and 50 ranges from 4 s to 0.04 s.

Measured: Excel for Windows through COM: Find after a cell by two sets (53 cases in each locale, predicted before each run), Replace by a new step of the harness (`Range.Replace`) and three sets (92 cases in each locale, predicted before each run); the known differences are a filtered sheet's second round by columns (Find Next skips a column there), text that reads as a number in a file without an apostrophe, and a range of one cell. That the kept texts change no answer: after each edit of a series of 23 every kind of search against a fresh session, in every kind of session.

## 0.29.0 (2026-10-13)

Internal release; not published to npm. What Excel writes for an item picked from a dropdown list. The synthetic book's values and hashes are as in 0.28.0.

Added:

- Each item of a list in `getValidation(sheet, cell).list` has `input`: what Excel writes into the cell when the item is picked from its dropdown, for `setCells`, which keeps the cell's format. A list of cells gives the cell's value as it is (`001` stays text, a number shown `50%` is 0.5); a literal list's item is read as typed in en-US, the locale a file holds the list in, in every locale (`"001,1/2,x"` gives 1 and the 2nd of January, `d-mmm`); a blank cell null. Excel checks no rule on a picked item ([docs/validation.md](docs/validation.md#an-item-picked-from-the-dropdown)).

Measured: Excel for Mac 16.113 on the keyboard, the new action `pick` of the harness's step `ui` (the cell selected, its list opened with Option+Down, the item reached with Down and taken with Return), the set `ui/pick-held`, 14 picks in en-US and ru-RU; its prediction, each item typed as its text with the rule checked, was committed before the run and came out on 5 in each locale. The engine gives Excel's answer to every pick.

## 0.28.0 (2026-10-13)

Internal release; not published to npm. Edits over a large look layer as fast as over a few cells, and Excel's own text for a rule's alert. The synthetic book's values and hashes are as in 0.27.0.

Changed:

- Rows and columns inserted and deleted move a sheet's `looks` by their blocks of rows and patterns of columns, the new lines with the look of the line before them, without turning the layer into entries; edits of looks (`patchCells`, `setStyles`, `clearCells` with `formats` and `all`, text and values typed into a cell of the layer) change the layer alone, its list made when it is read. On a sheet of the wide shape (5.6 million cells of the layer) inserting row 10 takes 6 ms and 1 MB of heap, where 0.26.0 took 38 s and 730 MB; deleting rows and undoing either take milliseconds. Copies and moves of ranges turn only their source and target into entries (0.1 s for 344 by 100 cells of that layer).
- `findCells` and `summarizeRanges` leave out the cells of a layer, which hold no value; `getLayout` does not make the list of looks.
- A rule's alert without a message of its own (`error`) says what Excel's alert says: `This value doesn't match the data validation restrictions defined for this cell.`, in ru-RU `Это значение не соответствует ограничениям по проверке данных, установленным для этой ячейки.` (the dialogs Excel for Mac showed in the set `ui/fill-held`).

Measured: each edit of a book with looks, its undo and redo against the same book with an entry per cell, and random series of 40 steps against full computations, in every kind of session; 120 random series of 60 steps against 0.24.0's entries, without a difference. Books without looks: structure operations, their undo and redo and cell edits within 4% of 0.27.0 in one process.

## 0.27.0 (2026-10-13)

Internal release; not published to npm. A circular reference without iteration holds its values as Excel does. The synthetic book's values and hashes are as in 0.26.0; a book with a circular reference computes otherwise, and its `inputHash` changes once its cells hold values.

Added:

- `CellInput.held`: the value a formula holds while a circular reference stops it, with iteration off. A formula in a cycle, or reading one, keeps the value it has, as Excel does: an edit of its other inputs and a full recomputation leave it, a formula entered there holds 0, and a value typed into a cell of the cycle breaks it, after which every cell computes. The values are part of the book (`getBook()` gives them), so a session is a full computation of its book after every edit and undo ([docs/cycles.md](docs/cycles.md)).
- `cached` in the options of `compute` and `createSession` (and the calculators): a file's stored values (`readXlsx().cached`, or a result of the same shape), which the cells of a circular reference and their readers hold when the book opens, as Excel opens a file; a cell the file stores no value for holds empty text.
- `CachedValues`, the type of those values.

Changed:

- A formula in a circular reference, or reading one, is no longer `#CYCLE!` without iteration: it holds its value, and its issue (`getCellIssue`, `diagnostics`, a window's `diagnostics`, `explain`) says the circular reference and where it runs through. A name that refers to itself is still `#CYCLE!`.
- `writeXlsx` writes a held value as the formula's value.

Known differences: an array that an edit brings into a cycle keeps its whole spill in Excel and its first value here; `INDIRECT` and `OFFSET` cycles and iteration from a file's stored values stay as before.

Measured: Excel 16 for Windows through COM, the workbook sets `cycles` (21 books: a file's values, `CalculateFull`, edits of predecessors, a value typed into the cycle, cycles through names and arrays, a file asking for iteration) and the held-out `cycles-held` (6 books, 129 reads, predicted before the run, 125 confirmed); the engine gives Excel's answer to all of `cycles` but the three named above.

## 0.26.0 (2026-10-12)

Internal release; not published to npm. Cells formatted across a sheet without an object per cell, and Excel's grouping commands. The synthetic book's values and hashes are as in 0.25.0; a book read from xlsx whose cells hold a look alone has a new `inputHash`.

Added:

- `SheetInput.looks`, the look layer of a sheet: rectangles (`LookRange`, `{ range, style?, numberFormat? }`), each over those before it, give their style and number format to the cells without an entry in `cells`; a cell with an entry shows its own fields, as before. A book is taken with every entry that has only a style or a number format moved into the layer, and `getBook`, a snapshot and `readXlsx` give the layer back in one form. `readXlsx` reads such cells into the layer without an object per cell and `writeXlsx` writes them back cell by cell, as Excel stores them: a book of the wide shape (5.6 million formatted cells, 344 rows to XFD) reads in 1.5 s with 2 MB of heap and opens a session in 7 ms, where it took 6 s and 815 MB and 10.5 s and 545 MB ([docs/xlsx.md](docs/xlsx.md)).
- `outlineGroups(sheet)` (`xlcalc`, `xlcalc/a1`): a sheet's groups of rows and columns as Excel draws them, each `{ first, last, level, summary, collapsed }`.
- `editStructure` operations `group`, `ungroup` (`rows` or `columns`), `hideDetail`, `showDetail` (`row` or `column`) and `showLevels` (`rows`, `columns`), as Excel's Group, Ungroup, Hide and Show Detail and the level buttons; refusals `outline_deep`, `not_grouped`, `detail` ([docs/outline.md](docs/outline.md)).

Changed:

- `getLayout()` gives each sheet without `looks`, as without `cells`. A result lists no cell of a look alone.
- `writeXlsx` writes a group's collapsed mark as Excel keeps and saves it: on the summary row of a group whose lines are all hidden, through `setLayout` too, and not on one whose lines show.

Known limits: inserting or deleting rows and columns over a large look layer turns the part of the layer it moves into entries for the operation (correct, slow on a sheet of the wide shape); the next release moves the layer itself.

Measured: the layer by equivalence, a book with looks and the same book with an entry per cell giving the same windows, cells, used ranges, hashes, files and books after each edit, in every kind of session, and by reading all 934 files of `test/excel` alike before and after. Grouping by a new step of the Windows harness (`Group`, `Ungroup`, `ShowDetail`, `Outline.ShowLevels` through COM, the collapsed marks read from a copy of the book after each command), 150 cases in four sets, each predicted before Excel ran it; the engine gives Excel's answer to all but 12, known differences (column groups of a span to XFD, lines hidden and shown by hand).

## 0.25.0 (2026-10-12)

Internal release; not published to npm. Excel's Find in a session, and two numeric rules of Excel that models depend on. The synthetic book's values and hashes are as in 0.24.0.

Added:

- `findCells` searches as Excel's Find ([docs/live-books.md](docs/live-books.md)): `wildcards` (`*`, `?`, `~`; `?` is one UTF-16 unit, so a character outside the BMP takes two), `wholeCell` (Match entire cell contents), `byColumns` (Search By Columns), `skipHidden` (cells of hidden rows and columns skipped as Find skips them: always in Look in Values, in every search on a filtered sheet, where the cells come by the blocks the hidden lines leave), and `in: 'edit'`, the text a cell opens with (Look in Formulas: a formula in the book's locale, a number by its format, a text without its apostrophe). Without Match case `ß` is `ẞ`, `σ`, `ς` and `Σ` are one letter, and `i` is not `İ`. On 34,000 rows a search of the whole data sheet takes 0.4 to 0.5 s, where reading every cell's edit text through windows took 0.7 to 1.2 s.

Changed:

- `findCells` gives the cells in Find's order, range by range, by rows or by columns, and an empty query finds nothing; a page without `count` reads every cell of the ranges.

Fixed:

- A negative number to a fraction is its real root where 1/y, to 15 significant digits, is an odd whole number below 2^31 - 1, computed as Excel computes it: `(-32)^(1/5)` and `POWER(-32,0.2)` are -2, `(-8)^(1/3)` is -1.9999999999999998; `(-8)^(2/3)` and `(-8)^(1/2)` stay `#NUM!`.
- A formula's last operation, outside parentheses, when it is `+` or `-`, is 0 where its answer is below eight units in the last place of its left operand, as Excel makes it: `=0.5-0.4-0.1` is 0, `=(0.5-0.4-0.1)` stays -2.78E-17. `SUM`, `SUBTOTAL` and `AGGREGATE` take the rule at their last addition, and `AVERAGE` is 0 where that addition is made 0.

Measured: Find by a new step of the Windows harness (`Range.Find` and `FindNext` through COM), three sets of 251 cases in en-US and ru-RU, each predicted before Excel ran it (136 of 155, 45 of 64 and 32 of 32 confirmed); the engine gives Excel's answer to all but seven cases in each locale, known differences (an empty query, `~` alone, a range of one cell). The power and the near-zero rules by five sets on Excel for Mac and Windows (334 cases), predicted before each run; Mac and Windows answer alike but the last digit of three roots. Excel's handling of circular references without iteration was measured on Windows and is not changed here.

## 0.24.0 (2026-10-12)

Internal release; not published to npm. A cell edited in place, as Excel reads it. The synthetic book's values and hashes are as in 0.23.0.

Added:

- `inPlace` in the options of `setCellText`, `setCellTexts`, `fillCellText` and their previews (`CellTextOptions`, `FillTextOptions`): the text was edited in the cell's own text (F2, a double click, the formula bar), not typed over the cell. Excel's percent editor then adds `%` only to a number begun by a sign: in a cell of `0%`, `6%` is 0.06 and `6` is 6 (shown 600%), where typed over the cell `6%` is the text `6%%` and `6` is 6%; `-5` is -5% either way ([docs/entry.md](docs/entry.md#a-percent-cell)).

Fixed:

- Typed over a percent cell, `-5%` and `+5%` are the formulas `-5%%` and `5%%`, as Excel makes them with its editor's second `%`; the engine wrote `-5%` and `5%`.

Measured: Excel for Mac 16.113 in en-US and ru-RU, on the keyboard. The new step `edit` of the harness's `ui` (the cell selected, F2, its text erased or a digit changed, the text typed, Return) ran the set `ui/edit-held`, 20 cases in each locale; its prediction, a paste's reading with no percent editor, was committed before the run and came out on 14 and 13 cases, the misses being numbers begun by a sign and a column Excel widened. The set `entry/typed-keys-held-3` (6 signed percents typed over a percent cell in each locale) found the fix above; its prediction came out on 3 and 4. The engine now gives Excel's answer to every case of both sets. `bench/coverage/input-source.swift` switches the keyboard's layout for such runs.

## 0.23.0 (2026-10-12)

Internal release; not published to npm. Calls for a grid or an agent working on a live session, cheaper incremental edits, and formula text over many linked sheets. The synthetic book's values and hashes are as in 0.22.2.

Added, in every kind of session (synchronous, `createInlineCalc`, the pool) and both execution modes ([docs/live-books.md](docs/live-books.md)):

- `applyChanges(operations, opts?)`: `setCells`, `replaceCells`, `clearCells`, `patchCells`, `setNames`, `setStyles` and `setLayout` operations applied in order, atomically, with one recalculation, one revision and one undo step (`UndoCall` `applyChanges`); `opts.clock` fixes the clock.
- `previewChanges(operations, opts?)`: their complete effects (`changes`, `impacts` with each cell's value before and after and its error, `totalChanges`, `truncated`, the clock used), the session, its revision and history left as they were. Applying it with `preview.revision` and `preview.clock` gives the same changes.
- `readCells(ranges, opts?)`, `findCells(ranges, query, opts?)` (a literal substring of the shown text, the formula or the value), `getDependents(ranges, opts?)` (immediate formula readers, their references as rectangles) and `summarizeRanges(ranges, opts?)` (cells, numbers, errors, a compensated sum, min and max): bounded pages with `maxCells`, a cursor of the same query and revision, and `count`.
- Types `SessionOperation`, `ChangesOptions`, `PreviewChangesOptions`, `ChangesResult`, `ChangesPreview`, `CellImpact`, `CellReadField`, `ReadCell`, `PageOptions`, `ReadCellsOptions`, `FindCellsOptions`, `QueryPage`, `ReadCellsResult`, `DependentCell`, `DependentsOptions`, `DependentsResult`, `RangeSummary`.

Changed:

- A session recomputes less after an edit: independent formulas the edit reaches compute on only where a value changed; a formula replaced with one of the same references keeps the dependency graph; `VLOOKUP` and `HLOOKUP` over a reference read their key column alone and keep its index across edits of the result columns ([docs/cache.md](docs/cache.md#shared-formula-bodies-and-lookup-key-indexes)).
- `eval` shares one compiled body among formulas of one shape: on 50,000 formulas `$A$1*r` the heap is 62 MiB instead of 91 and RSS 252 MiB instead of 391. `closures` keeps a tree for each formula, which opens such a book faster than a shared body did.

Fixed:

- A session's window with formula text (`formulaLocal`, `editText`), its explanations and its typed text checked the book's links again for each formula: with 300 links of 80 sheets a window of 600 formulas took 0.26 s, now 0.006 s (the rest of 0.22.2's fix).

Measured: the new calls by their tests in every kind of session, and a random series of edits, mixed batches, previews and undos (60 steps in each kind and mode) with the session's hash equal to a full computation after each step and every preview leaving the book, hash and revision as they were. Speed on the synthetic book against 0.22.2, both in one process: cold computation and a cell's edit within 2.6 % either way at 3,400 and 34,000 rows, a fill of 2,000 rows within 6 %.

## 0.22.2 (2026-10-12)

Internal release; not published to npm. A window's formula text over links to other workbooks as fast as without them. The synthetic book's values and hashes are as in 0.22.1.

Fixed:

- `getViewport` with `formulaLocal` or `editText`, `toLocal`, `toInvariant`, `validateFormula`, `validateLocalFormula`, `validateNames`, `parseEntry` and the text-entry calls of a session checked every stored value of every linked book in their options, once for each formula: a window of 600 formulas over 100 links of 1,000 stored cells took 8.7 s, over 300 links of 5,000 cells 76 s (0.22.0). Their options now check the links' targets, sheets and names alone, which is all they read; a book's stored values are checked once, as it opens. The same window takes 0.03 s.

## 0.22.1 (2026-10-12)

Internal release; not published to npm. The ribbon's Number Format box in ru-RU as Excel for Windows shows it. The synthetic book's values and hashes are as in 0.22.0.

Changed:

- `numberFormatCategory(format, 'ru-RU')` follows Excel 16.0 for Windows, not Excel for Mac, which names nearly every code `(все форматы)` in ru-RU as it compares the code with its formats written in Russian syntax. `0.00`, `#,##0` and `0.0%` are Числовой and Процентный as in en-US; the ruble quoted after the amount (`#,##0.00\ "₽"`, the Currency item's format) is Денежный and ru-RU's accounting form Финансовый, the dollar's forms `(все форматы)`; a fraction needs its space quoted (`#" "?/4`); the short date is `dd/mm/yyyy`; built-in formats by number take the category of ru-RU's own code (5 to 8 Денежный, 41 to 44 Финансовый, 37 to 40 `(все форматы)`). en-US is unchanged.

Measured: a new Windows step, `ribbon`, reads the box by UI Automation in a session logged on through Remote Desktop, over the books of the Mac's ribbon set (342 cells, all named as Excel for Windows) and a held-out set of 50 cells (34; the 16 others are dates and times with `;@` that Windows names by the lists of its Format Cells window, pinned with the reason). Excel for Mac's `(все форматы)` in ru-RU is a known difference by rule ([docs/formats.md](docs/formats.md#the-ribbons-number-format-box)).

## 0.22.0 (2026-10-12)

Internal release; not published to npm. Links to other workbooks: a book that reads other workbooks computes from the values Excel stores of them in the file, as Excel computes a file it opens without updating its links, and the ribbon's Number Format box without the engine. The synthetic book's values and hashes are as in 0.21.2.

Added:

- `BookInput.externalBooks`: the other workbooks a book's formulas read, each its link's `target`, its sheets with the values stored of them (`null` for a cell stored empty, `refreshError` for the mark of a failed update) and its names; element `i` answers `[i+1]`. It goes into `inputHash` when it is not empty. `readXlsx` reads it from a file's links (`xl/externalLinks`) and `writeXlsx` writes it back as Excel writes links; a DDE or OLE link keeps its number and is reported.
- References to other workbooks in formulas and names: `[1]Sheet1!A1`, `'[1]My Sheet'!A1:B9`, the book's own name `[1]!Rate`, a sheet's `[1]Sheet1!Local`, a lost one `[1]Sheet1!#REF!`, stored as `'[1]Sheet1'!A1`. A cell the link stores gives its value; one it does not is blank, or `#REF!` on a sheet marked `refreshError` or of which nothing is stored. Ranges, whole columns and rows, `INDEX`, `MATCH`, `VLOOKUP`, `XLOOKUP`, `SUMPRODUCT` and the rest read the stored values; the criteria functions' ranges, `COUNTBLANK` and the database functions give `#VALUE!` over a link, as Excel gives them over a closed book. Entering or pasting a formula with a link marks that book as Excel marks it, and a session's book takes the mark. A copy moves a link's relative references; inserting rows and renaming sheets leave them.
- `ValidateContext.externalBooks`, `NamesContext.externalBooks`, `EntryOptions.externalBooks` and the option `externalBooks` of `toLocal` and `toInvariant`: a reference to a link is checked and numbered, shown as Excel shows it (`[Prices.xlsx]Sheet1!A1`, `'../data/[Prices.xlsx]Sheet1'!A1`, `Prices.xlsx!Rate`) and read back by the file it names; one to a workbook the links do not hold is `external_ref`, a new `Translation` code too. A window's `formulaLocal` and `editText`, `explain` and `setCellText` use the book's links.
- `numberFormatCategory(format, locale?)` and `numberFormatChoices(locale?)` in `xlcalc/text` and the root: the category the ribbon's Number Format box shows for a number format (General, Number, Currency, Accounting, Date, Time, Percentage, Fraction, Scientific, Text, Special, Custom) with its name in the locale, and the box's items with their codes, by Excel for Mac in en-US and ru-RU. Types `NumberFormatCategory`, `FormatCategory`, `FormatChoice`, `FormatChoiceKey`; ru-RU's names in `regional.formats`.

Changed:

- `validateFormula` without `externalBooks` refuses a reference to another workbook as before; a quoted sheet with a book in brackets (`'[1]My Sheet'!A1`) is now such a reference, not a sheet named `[1]My Sheet`.

Bundle sizes: `compute` 262 KB, `xlcalc/text` 36 KB (14.6 KB gzip), `xlcalc/xlsx` 184 KB, `xlcalc/a1` as before.

Measured: the step `links` of the harnesses (Excel for Mac with `update links do not update links`, Excel for Windows with `UpdateLinks` 0), six sets, both Excels alike; the ribbon by the step `ui` on Excel for Mac, 342 formats in each locale. Known differences: a closed workbook's reference is shown with the path its link holds, where Excel shows the absolute path; in ru-RU Excel for Mac's box shows "(все форматы)" for most formats, which the function follows; Windows's box was not measured ([docs/links.md](docs/links.md), [docs/formats.md](docs/formats.md)).

## 0.21.2 (2026-10-12)

Internal release; not published to npm. Faster incremental edits and less retained engine storage, with the same public API, values and synthetic hashes. The reproducible HyperFormula 3.5.0 comparison and memory controls are in `bench/hyperformula/README.md`.

Changed:

- Contiguous edits in a column query dependent ranges once for each contiguous span. Address parsing and change tracking allocate less temporary data; sessions without arrays skip spill lookups.
- Shared aggregate prefixes keep checkpoints every 64 cells instead of 256. Repeated short aggregates can reuse a cached fold; the full-fold threshold and cache budget are unchanged.
- Small acyclic edits stop an unchanged scalar path until another result changes or an array appears. Structural edits, existing arrays and iterative calculation keep the full pass. An unchanged INT at the start of a 5,000-cell chain evaluates one formula instead of 4,999 in both execution modes.
- Removed formulas release their storage slots for reuse, trailing holes are trimmed, and spill timestamps are released after the round settles. Range values share their metadata's map. Eval releases generated text after linking its executable functions; diagnostic text is regenerated on request.

Measured in Node 26.7.0, medians of five fresh-process pairs: eval retains 90.99 MiB instead of 105.76 MiB on the fanout book, and 56.22 MiB instead of 58.05 MiB on the 34,000-row synthetic book. A cache of 50,000 small entries retains 10.81 MiB instead of 12.62 MiB. Peak process RSS on the large books changes little because compilation still needs temporary storage. The ordinary edit comparison includes both execution modes; bulk edits and indexed lookups remain slower than HyperFormula.

## 0.21.1 (2026-10-12)

Internal release; not published to npm.

Fixed:

- `setCellTexts` and `previewCellTexts` with `refuseInvalid` refuse a batch by the strictest rule its values break (stop, then warning, then information; the first cell of that style), not by its first cell. A batch with a warning before a stop was refused by the warning, and a host that showed it and kept the value called again without the option, writing the value the stop does not take.

## 0.21.0 (2026-10-12)

Internal release; not published to npm. What the sheet component asked for as it moves to the final API: a typed value a rule of data validation does not take refused before it is written, a formula shown with a sheet in quotes only where Excel puts them, an empty cell of a formatted row or column in a window, the issues of xlsx in the book's language, and no string `xlcalc` in the subpaths without the engine. Values and hashes are as in 0.20.0.

Added:

- `refuseInvalid: true` for `setCellText`, `setCellTexts`, `fillCellText` and their previews: a value a cell's rule of data validation does not take, of any style, is refused (`code: 'validation'`, the cell, the rule, its style and message; the types `ValidationRefusal` and `CellValidationRefusal`) and nothing changes, no revision and no step of history. A host shows Excel's alert first and calls again without the option where the person keeps the value (a warning's Yes, an information's OK). A batch is refused whole, with the first cell whose value breaks its rule. Without the option a value is written with its alert, as before.
- `readXlsx(data, { locale })` and `writeXlsx(snapshot, { locale })`: each issue's message in the locale, a loaded one's tag or a locale object, its code and other fields as in en-US; ru-RU has all 95 messages. Without the option the issues are as before. A file does not say its language: a host gives the locale the person chose.

Changed:

- A formula shown to a person (`toLocal`, a window's `formulaLocal` and `editText`, `explain`'s text) writes a sheet in quotes only where Excel puts them, by `quoteSheetIfNeeded`: `=Data!B2*2`, `='Q 4'!A1`, and `'ИСТИНА'!A1` in ru-RU. The stored canonical form quotes every sheet as before, and `toInvariant` reads both.
- A window (`getViewport`, `viewportOf`) lists an empty cell with a number format of its row or column alone, with `numberFormat`, as it lists one with their style, so a grid shows the format Excel shows for it.
- Without the root, a locale not loaded is refused asking for the locale as an object; the root names its module and `registerLocale` as before. So `xlcalc/text` and `xlcalc/a1` bundle no string `xlcalc`, which a host may look for as the engine's sign.

Bundle sizes (README): the root 1.5 KB more for the keys of messages a locale may give, `xlcalc/xlsx` 174 KB, the ru-RU module 92 KB (15 KB gzip), the pool's thread 499 KB; `compute` and `xlcalc/text` as before.

Known differences: a reference across sheets where one end needs quotes (`'S1:S3'!A1`) is shown as Excel's grammar writes it, not measured; the reason a table of a file is left out stays in English in ru-RU.

## 0.20.0 (2026-10-11)

Internal release; not published to npm. A snapshot is a copy in every kind of session, and an edit's changes say when a note or a hyperlink changed. Where two copies of one piece of logic answered differently, Excel for Mac and for Windows were asked and the copies made one; undoing the deletion of a large sheet is faster. The synthetic book's values and hashes are as in 0.19.0; values change only where a copy disagreed with Excel.

Added:

- `CellChange.note` and `CellChange.link`: an edit lists a cell whose note or hyperlink changed, with the new one, or null when it was removed, in `setCells`, `replaceCells`, `patchCells`, `clearCells`, `undo` and `redo`, on an otherwise empty cell too; equal notes and links are no change. `meta` stays the host's data and is not listed.

Changed:

- `getSnapshot()` gives an independent copy of the book and its result in every kind of session: later edits do not change it, and changing it does not change the session. A synchronous and an inline session gave their own book, which later edits changed, beside the result of the moment, so a snapshot written later held a book and a result that did not match; a pool gave a copy. Copying takes about 0.27 s on the 34,000-row book (0.94 s to 1.20 s) and the memory of the copy while it is held; it goes in portions, so an inline call gives way and can be stopped. A pool's snapshot costs nothing more: the thread's transport copies it.
- `EntryOptions.cell` says what to pass for a row's or a column's number format (the cell of `getViewport`, or the format resolved): a cell of `getCell` or `getCells` carries its own format alone. `parseEntry` reads as before.
- Undoing the deletion of a large sheet takes 35 to 41 % less on the 34,000-row book (1.77 s to 1.05 s without the shown text), inserting and deleting rows and their undo and redo 13 to 22 % less. A state built again copies only the cells `getBook` gave a host and takes the others as they are. Every answer, book and step of history is as before.

Fixed:

- A moment to the second: `DAY`, `MONTH`, `YEAR`, `WEEKDAY`, `HOUR`, `MINUTE`, `SECOND` and a number format's whole seconds split a number as Excel does, half a second added and the sum divided into days and seconds, as both Excels answer 420 cases alike. The date functions and the time functions split it two ways before and disagreed on 318 of 2 million serial numbers tried: 0.4 s before midnight is now the next day in each, and `SECOND(45688+1/256)`, 37.5 s past a minute, is 37.
- Words of formula text (function names, `TRUE` and `FALSE`, LET's names, a locale's words) compare as Excel for Windows compares them, each letter by its own upper case, the dotless `ı` and the long `ſ` apart: `falſe` is a name, not FALSE, and `gauß` is no `GAUSS`. `falſe` read as FALSE before, and `LET(falſe,1,falſe)` lost its name. Defined names and tables (`nameKey`, `sameName`) compare in any case with `ß` as `SS` and `ﬁ` as `FI`, but `ı` and `ſ` apart from `I` and `S`.
- A file's formula reads `?` inside a name as typed text does: `a?_xlfn.b` keeps its prefix and `T?[x]` is the table `T?` (they read `a?b` and an external reference).
- xlsx: a rule's cells (`sqref`) of conditional formats and data validation are read in any case and with `$`, as both Excels read them; a conditional format on `a1:b2` and a rule of data validation on `$A$1:$B$2` were lost.
- xlsx: a note on `A01`, `$A$1`, `a1` or `A1:A1` is on A1, as both Excels show it (its key `A01` made `createSession` throw on the book read); a note on a range, a file no Excel opens, is left out and reported.
- xlsx: fills are read as both Excels show them. A cell's fill without a pattern type is none, and its solid fill shows its `fgColor`, white without one. A conditional format's (a dxf's) fill is solid without a pattern type and shows its `bgColor` alone; a colour that names none (`auto`, indexed 64 and 65) is white, and an empty pattern fill changes nothing. What the model cannot hold (a gradient, a dxf's fill that takes the cell's away or keeps the cell's colour) is reported lost.
- A conditional style's border side given as `undefined` is no side, as in a cell's style and every field of input; it was an `InputError`.

Known differences (README): Excel for Mac reads `ſ` as `s`; Excel compares sheet names letter by letter (a sheet `ß` beside `SS`) and reads a typed function's dotted `İ` as `I`; a cell's shown text of a moment a few millionths of a second from half a second may show another second, as the text is written from the value to 15 digits.

Excel harnesses: the steps `cond`, `valid` and `parts` send a case's file changed by hand (`raw`), and a repair window on another Space is declined after Excel is brought forward. New sets, each predicted before Excel ran it and answered by Excel for Mac and Windows: `seconds-held`, `name-case-held`, `name-case-sheets-held`, `name-case-values-held`, `name-case-ru-held`, `name-question-held`, `cond/sqref-held`, `valid/sqref-held`, `parts/refs-held`, `cond/fills-held`, `cond/fills-held-2`.

## 0.19.0 (2026-10-11)

Internal release; not published to npm. The old names of the API are gone, so that what our applications forget to move fails at once; a paste and an entry into a selection as Excel takes them; a saved version's conditional formats; structure operations faster. Values and hashes are as in 0.18.0.

Breaking changes (each a line of `docs/migration.md`, which now says what an old name gives):

- The names 0.17 renamed are removed, no longer synonyms: `openSession`, `restructure` (the function and a session's method), a session's `book()`, `layout()`, `extent()`, `cells()`, `validation()`, `result()`, `hash()` (calling one is a `TypeError`; the names are `createSession`, `editStructure`, `getBook()` and the rest), a calculator's `openSession`, the object `engine`, `fromLegacy` from the root (it is in `xlcalc/xlsx`), `writeXlsx(book, { result })` (it takes `{ book, result }`), the types `RestructureResult` and its kin, `ConditionalLook`, `FormulaEngine`.
- Input takes the number format as `numberFormat` alone: `format` in a cell, a row or a column, a patch, a layout edit, `parseEntry`'s cell, `displayValue` or `entryWidth` is an `InputError` that names `numberFormat`.
- `viewportOf` takes a snapshot of the current shape (`BookSnapshot`); results saved under the old field names are not read.

Added:

- `Session.pasteText(sheet, cell, text)` and `previewPasteText`, with `parseClipboard(text)` (also in `xlcalc/text`): clipboard text pasted as Excel pastes it, measured in Excel for Mac and Windows. Each value reads as typed, but the paste widens no column, checks no rule of data validation, keeps an apostrophe in the text (`'007`), pastes a formula Excel does not read as text, has no percent editor in a percent cell (`5` is 500 %), and fills the clipboard's rectangle, a short row clearing its missing cells.
- `Session.fillCellText(sheet, range, text)` and `previewFillCellText`, an entry into a selection (Control+Return in Excel): the text read once in the active cell, its formula moved into each cell as a copy moves it, the active cell's number format and rule of data validation, a column widened as typing widens it.
- `viewportOf(snapshot, sheet, range, { appearance: true })`: what the version's conditional formats give its cells, computed over the saved values at the version's clock without computing the book's formulas, cached per snapshot (the first window of the 34,000-row book about 0.16 s more).

Changed:

- After a structure operation one pass over the cells serves the comparison with each cell's old place and the history: inserting or deleting a row on the 34,000-row book takes 1.05 s (1.29 s before), 0.24 to 0.33 s less with the values shown, undoing the deletion of a large sheet about 0.27 s less. Every answer is as before.

Fixed:

- The text an old-kind array formula's cell opens with for editing (`editText`) is the array's formula in every cell of its area, as Excel opens it (it was empty but in the top-left cell).
- Many cells written into one column in one call were quadratic (each cell sorted the column again): 40,000 cells took 8 s, now 0.2 s.

Known differences (README): a pasted or entered formula over a date takes the date's format in Excel; Excel for Mac's short date in en-US is `m/d/yy`. A saved version's appearance differs from a session's where a rule reads digits past the 15 a result keeps, a `SUBTOTAL` over a cell with its own `SUBTOTAL`, or a function the host adds.

Excel harnesses: a step `ui` reads the formula bar and pastes and enters text on the keyboard; the Windows answers of `struct/edge-wide` are recorded again (the engine now matches them) and `probe.ps1` is gone.

## 0.18.0 (2026-10-11)

Internal release; not published to npm. What our applications still wrote themselves over 0.17, added: a window over a saved version without a session, typing into many cells at once, helpers for styles, lines and formula text without a session, and `xlcalc/text` for a browser without the engine. Repeated logic of the engine and the Excel harnesses is in shared helpers. Values, hashes and the files written are as in 0.17.0; nothing that worked is called otherwise.

Added:

- `viewportOf({ book, result }, sheet, range, opts?)`: the window `Session.getViewport` gives, over a book and its result saved earlier, without a session and without computing anything again: values, text and errors' causes from the result, styles, formats and the layout from the book. A result saved by 0.11 to 0.16 (`format`, `issues`, `nameIssues`) is read too. A result carries no conditional formats' appearance, so this window has none.
- `getViewport` and `viewportOf` take `rows` and `columns`, the numbers of the lines to give within the range (rows a filter shows, columns apart): each row and column of a window has its number (`row`, `column`); and `editText`, the text each cell opens with for editing, as Excel's formula bar shows it (a formula with `=` in the book's locale, a value by its format, an apostrophe where it would read otherwise).
- `Session.setCellTexts(entries)` and `previewCellTexts`: texts typed into many cells as `setCellText` reads each in its cell, in one step of undo and one recalculation, a column widened once to the widest; a formula Excel would not take in one cell changes nothing and names its cell; broken rules of data validation listed by cell. `Session.getCell(sheet, cell)`.
- `numberFormatCode` beside `numberFormat` in a change and a cell's view: the code the format shows by in the book's locale (14 is `dd.mm.yyyy` in ru-RU).
- `internStyle(styles, style)`: a style's number in a table of styles, added when it is not there, as `patchCells` finds one; fast along a chain of calls. `patchLines(map, span, patch)`: a map of columns or rows (`SheetInput.columns`, `rows`) with a span patched as `setLayout` patches it; also in `xlcalc/a1`.
- `toInvariant(text, locale, { table })` and `ValidateContext.table`: a formula typed in a table's cell (`[@Qty]`) translated and checked without the book, as `toLocal` takes `table`. `escapeColumn`, a column's name as a structured reference writes it.
- `xlcalc/text` (29 KB minified): `scanReferences`, `quoteSheetIfNeeded` and `escapeColumn` without the engine, for a browser that colours references as a formula is typed.

Changed:

- `scanReferences` reads on past a token that does not read: `=A1+)B2` gives `A1` and `B2` (it stopped at `)`); a reference inside an unclosed string is none.

Fixed:

- `readXlsx` looks a file's words up by its tables' own keys: a stored error `toString` was read as a function, and a rule of data validation of type `constructor`, an icon set or a relationship of such a name made the reader throw. They are words the model does not know, read as other unknown ones, with an issue.

Internal: the duplicates a scan found are in shared helpers: numbers with the first error the answer, the 30/360 count, serial numbers, statistics; the scanners of formula text, a formula printed for a file, the error texts from one table, the words of table specifiers; the lists of the xlsx reader, writer and input checks, the cells of a rectangle; one frame for the Excel harnesses of the Mac and Windows. Repeated code by `jscpd` (8 lines, 70 tokens): `src` 2 to 0, `bench` 47 to 29. The dead argument kind `D` is gone. Bundle sizes in the README measured again one way: `createSession` 387 KB, the root 421 KB.

## 0.17.0 (2026-10-11)

Internal release; not published to npm. The API made simpler before publication, in one go so that its users change once: a session for a grid (typing text into a cell, a window with its styles resolved, editing that keeps a cell's look, a style by its fields), one shape of the answer of every edit, xlsx with a snapshot, helpers without the engine, locales as modules, the new names with the old ones as deprecated synonyms. `docs/migration.md` lists every rename and every call that answers otherwise, a line each. Values and hashes are as in 0.16.3; `inputHash` too, whichever name a book gives the number format.

Breaking changes (each a line of `docs/migration.md`):

- `setCells` sets the fields an edit gives and keeps the cell's others (its style, number format, note, link, `meta`), as Excel keeps a cell's look when a value is typed over it; `null` removes a field, and `input: null` clears the content and the hyperlink as Excel's Delete key does. `replaceCells` replaces a cell whole, as `setCells` did.
- Every call that changes a session's book answers an `EditResult`, `{ revision, changes, layout?, styles?, moved?, repaint? }`, instead of an array of changes (`setStyles` answered nothing): the array is `changes`.
- A book in ru-RU needs its locale loaded: `import { ruRU } from 'xlcalc/locales/ru-RU'` and `registerLocale(ruRU)` once (a pool's threads get it themselves); a book in a locale not loaded is an `InputError` that says what to import. The root no longer exports `ruRU`.
- `readXlsx`'s `cached` has the shape of a result (`cached.sheets[0].cells.B1.value`), so it can be written again as a result.
- Answers name the number format `numberFormat` (it was `format`), a cell's conditional look `appearance` (`look`), and a result's issues `diagnostics`, one list of the cells with an error (`kind: 'cell'`) and the names with a problem (`kind: 'name'`) (`issues`, `nameIssues`). Input takes `format` as before.
- A session's `editStructure` answers without the book (`getBook()` gives it), as an asynchronous session did; `mapCell` answers null for a sheet it does not know, as for a cell gone; a book in an answer is a `Book`, the input type without the deprecated `format`, so reading `format` from it does not compile.
- `validateEntry` refuses with `{ ok: false, code: 'validation', message, rule, severity }`; `explain`'s text writes an error as the cells show it in the book's locale (`#DIV/0!`), with the cell it began in, and an empty cell as `empty`; after a structure operation `changes` also list the cells whose formula's text, own style or own number format differ from the cell that was at their old place; `undo` and `redo` answer `call: 'editStructure'` for a structure operation.

Renamed, the old names deprecated synonyms until our applications move (to be removed before publication): `openSession` → `createSession` (and a calculator's), `restructure` → `editStructure`, a session's `book()`, `layout()`, `extent()`, `cells()`, `validation()`, `result()`, `hash()` → `getBook()`, `getLayout()`, `getUsedRange()`, `getCells()`, `getValidation()`, `getResult()`, `getHash()`; the input field `format` → `numberFormat`; the types `RestructureResult` and its kin → `StructureResult`, `ConditionalLook` → `ConditionalAppearance`; `fromLegacy` from `xlcalc/xlsx`. The object `engine` is deprecated whole.

Added:

- `Session.setCellText(sheet, cell, text)` and `previewCellText`: text typed into a cell as Excel reads it in the book's locale, date system and clock, by the number format the cell shows (its row's or column's too, not written into the cell when inherited), keeping its style, `meta`, note and link; the column widened as Excel widens it, in the same step of undo; a broken rule of data validation in the answer (`validation: { valid, rule, severity, message }`); a formula Excel would not take a `Refusal` with its position, in the book's locale.
- `Session.getViewport(sheet, range, opts?)`: what a grid draws of a window, its cells with their own and shown style by number and the objects of those styles in `styles`, the number format as a code and where it comes from, what conditional formats give them, notes, links, empty cells a row's or column's style shows, the widths and heights of its lines, the merged areas and tables that meet it, the sheet's frozen panes, the used range and the session's revision; formulas in the book's locale (`formulaLocal`, `[@Qty]` inside its table) and errors' causes (`diagnostics`) when asked; `maxCells` (20,000 by default).
- `Session.getCellIssue(sheet, cell)`: an error's reason and the cell it began in, without the book's result; `explain` follows an error to where it began, through ranges and past its depth, never round a cycle.
- `Session.clearCells(ranges, 'contents' | 'formats' | 'all')`, as Excel's Clear Contents, Clear Formats and Clear All clear cells, notes, hyperlinks, conditional formats and data validation (measured in Excel for Mac and Windows).
- `patchCells` over ranges and whole rows and columns, with a style by its fields merged into the look a cell shows (`null` removes a field), no duplicates in the style table; whole lines formatted as Excel formats them, the cells where they cross merged field by field.
- `Session.revision` and `baseRevision` on every edit: a call on a revision the session has left changes nothing and throws `CalcError` `stale_revision`. `Session.getSnapshot()`: the book and its result in one call.
- `writeXlsx({ book, result? }, { values: 'require' | 'omit' | 'preserve' })`: a result of `compute`, of a session or a read file's `cached`; without values of formulas the file asks Excel to compute them and `issues` says so (`noCachedValues`). `readXlsx` takes `Uint8Array`, `ArrayBuffer` and `Buffer`; a session, a path or a result of another book is an `InputError` with what to pass.
- `xlcalc/a1` and the root: `parseCell`, `cellAddress`, `colLetters`, `colNumber`, `parseRange`, `rangeAddress`, `parseLines`, `linesKey`, `MAX_ROWS`, `MAX_COLS`, `rectsMeet`, `rectContains`, `quoteSheetIfNeeded` (Excel's rule for quoting a sheet name, measured), `nameKey`, `sameName`, without the engine (6 KB). From the root: `validateSheetName`, `errorText`, `scanReferences` (the references of a formula being typed), `entryText` and `needsApostrophe` (the text a cell opens with for editing, measured against Excel's formula bar), `sameValue` and `canonValue` (as the values hash compares), `builtinFormatCode`, `dateToSerial` and `serialToDate`, `registerLocale`.
- A check answers `{ ok: true, … }` or a `Refusal` (`{ ok: false, code, message, at? }`); `validateName` has codes.
- A locale by its tag or as an object in every function that takes one.
- A call refuses an option or an argument it does not take, with the call that takes it (`getCells` given `{ formulaLocal: true }` says it is an option of `getViewport`); `writeXlsx` given a result alone says to pass `{ book, result }`.
- The package holds `docs/`; the README starts with ten recipes, says which call does which edit, and each part of a session is in one place.

Changed:

- A formula's leading `=` is dropped where a book or an edit gives a formula (`==1` stays an error). `displayValue` without a width does not narrow a value to `#`.
- `compute` alone bundles to 254 KB minified (304 KB before), held under 300 KB by a test: the data of ru-RU is a module of its own (52 KB).
- Structure operations describe where cells go once, by numbers, and derive `mapCell`, `moved`, the geometry of undo and redo and the comparison with a cell's old place from it; each sheet's cells are read once per operation. On the 34,000-row synthetic book inserting a row takes 1.25 s (1.94 s before), deleting it 1.15 s (1.69 s), copying and moving about 40 % less, their undo and redo about 25 % less. The undo geometry is exact where it was guessed: a delete on a filtered sheet lists one move per run of visible rows, and a redo of `deleteTable` no longer lists the table's cells as changed; `mapCell` answers null for a cell pushed past the last row.

Fixed (measured with the engine's answers committed before each run, in Excel for Mac and Windows):

- `BINOM.DIST.RANGE` past n = 170 (`#NUM!` before), as Excel up to 2^31 − 2.
- Functions over two arrays (`SUMX2MY2` and its kin, `CORREL`, `PEARSON`, `RSQ`, `SLOPE`, `INTERCEPT`, `STEYX`, `COVARIANCE.P` and `.S`, `FORECAST`, `CHISQ.TEST`, `T.TEST`) and `SUMPRODUCT` read a reference to one cell first, its error the answer and a blank `#VALUE!`, then the sizes, then the errors inside.
- A rule of data validation of one cell on the sheet's last row or column stretches when lines above it are deleted, as conditional formats do.
- A copy of tables onto a target that is no multiple of the source pastes one copy at its top-left.
- Names shaped like cells: `XFE1` and `A0` are names, as LET's names are any cell but those read as R1C1 rows or columns; LET's names with `?` and `\` read from files.
- A calculated column's formula reports what the engine cannot take of it in `issues`; a conditional format's built-in number format shows by its en-US code in every locale; `{ date: '1900-02-29' }` is 60; a format's fill (`*x`) shown without a width no longer throws.

Known differences (README): a formula pasted into part of a calculated column, which Excel spreads to the whole column; `LET(IW1,5,IW1)+IW1`, whose last `IW1` Excel reads as a defined name.

## 0.16.3 (2026-10-11)

Internal release; not published to npm. A fill no longer walks the whole sheet for each copy it pastes.

Fixed:

- `copyRange` (and so a fill down or right, and a copy over a target several times the source) kept the look of each pasted position by walking every cell of the source sheet once per copy pasted: filling one row of the 3,400-row synthetic book down 2,000 rows took 11.4 s (192 s on the 34,000-row book). The source's cells and their looks are now found once for all the copies: 0.2 s and 2.6 s. The book, the values and the changes are the same.

## 0.16.2 (2026-10-11)

Internal release; not published to npm. `LINEST` and its relatives far from 0, `SORTBY` with a by-array of one value.

Fixed:

- `LINEST`, `LOGEST`, `TREND` and `GROWTH` with a constant take y and each variable about their means before the Householder QR and bring the intercept back after, so data far from 0 against their spread keep their digits: of 10,000 values 1E12 + 0, 0.01, 0.02 on their steps 0, 1, 2 the slope is 0.010009765625, as Excel gives it (0.0101547 before). The known difference `linestOffset` goes. The variables Excel leaves out and the refusals are as before. Measured with the engine's answers committed before the run: the set `fit-offset-held` (43 cases: `LINEST` with two variables and without a constant, `TREND`, `LOGEST`, `GROWTH` on shifted data) in Excel for Mac and Windows, which answered alike to the bit; 34 cases as predicted, two more once `TREND` stayed b + m·x as Excel evaluates it.
- `SORTBY` takes a by-array of one value as one row and one column, and the first by-array decides whether rows or columns are sorted: `SORTBY(D!B1:B8,D!C1:C8,1,5,1)` is `#VALUE!`, as in Excel (the engine sorted). The set `sortby-one-held` (19 cases) as predicted in both Excels.

Known differences (README): `linestMean`, Excel's intercept and regression sum of squares over y far from 0 follow the mean it sums plainly, so they and `TREND` there differ from the exact ones here; `growthRange`, `GROWTH` far from 0 in x is `#NUM!` in Excel where the fitted value fits a double.

## 0.16.1 (2026-10-10)

Internal release; not published to npm. The cells of formatted rows and columns in xlsx as Excel shows them.

Fixed:

- `writeXlsx` wrote a cell without a format or style of its own in a formatted row or column with none, and Excel showed it in General (46096 for a date in a column of dates): Excel shows a stored cell by its own cell format alone. Every cell is now written with the look the engine shows for it (its own over its row's, else its column's), and where a row with a look crosses a column with another one, the empty cells of the crossing are written too, as Excel writes them. A sheet without formatted lines is written as before. `readXlsx` keeps of a cell what its row or column does not give it, so a written book reads back as it was, and spells out a default a formatted row would not give (General, a plain style). Measured with the engine's answers committed before each run: the set `lines` (`files`, `held`) in Excel for Mac and Windows, 500 of 501 cells held out as Excel shows them; the one left is below.
- The look rows and columns give a cell is worked out in one place for showing cells, AutoFit, copy and paste, the lists of data validation and xlsx.

Known difference: when Excel formats a line across a line with a style, it gives the cells of the crossing both styles merged field by field; `setLayout` sets a line's style whole, so the engine shows the crossing by the row's style alone. A host formatting a line across a formatted one gives the crossing cells their merged style, as Excel does.

## 0.16.0 (2026-10-10)

Internal release; not published to npm. Undo and redo in sessions.

Added:

- Contract: `Session.undo()` and `Session.redo()` undo and redo the last change of the book as a whole, as Excel does: a call of `setCells`, `patchCells`, `setNames`, `setStyles`, `setLayout` (AutoFit too) or `restructure` (hiding and showing a sheet too) is one step, and a new change after an undo drops what could be redone. They give an `UndoResult`, `null` when there is nothing to undo or redo: the `call` undone, its `changes` as `setCells` gives them, the `cells` whose inputs came back (for a step of `setCells` or `patchCells`) and a structure operation's `mapCell`. `history` tells how many steps each way, for a toolbar. The same in same-thread and pool sessions, with `timeoutMs` and `signal`.
- Contract: `SessionOptions.undo` (`depth`, 100 by default, 0 turns the history off; `maxMb`, 64 by default): older steps are dropped past either. In a pool thread with `maxHeapMb` a session without its own `maxMb` takes `maxHeapMb / (8 × maxSessionsPerThread)` (8 MB for 384 MB and six sessions), so histories cannot fill a thread's heap.
- A step keeps what its call replaced: the inputs of the cells it changed, the layout fields and lines of a sheet, the style table, and for a structure operation or new names the part of the book its geometry does not carry back (deleted cells, rewritten formulas). A deleted sheet's cells are kept as they are; when they are larger than the budget, the history is cleared, as Excel cannot undo deleting a sheet. A pool journals undos and redos and carries the history in its snapshots, so a session restored after its thread was lost or after it left a full thread keeps it.
- On the synthetic book of 34,000 rows a cell edit's step holds 0.08 MB, `insertRows` 0.43 MB (about 10 % slower with the history), deleting 1,000 rows 1.86 MB (17 to 22 % slower); undoing a structure operation builds the state again, as a cold computation does.

Changed:

- `Session.cells()` reads a sheet's rows and columns once for a session's state and finds a row by halving: a window over 34,000 row heights takes 0.9 ms instead of 9 ms, 15 ms instead of 42 ms from a thread.

Known: after an undo of a structure operation the session's book holds the same content, but the keys of its objects may come in another order.

## 0.15.0 (2026-10-10)

Internal release; not published to npm. Columns as wide as Excel makes them.

Added:

- Contract: AutoFit. The layout edit `{ sheet, columns, fit: true }` gives the columns the widths Excel for Mac's AutoFit gives them (a double click on a column's border) by what their cells show: text, numbers by their formats, dates, booleans, errors, spills, hidden rows, cells merged down; cells merged across columns are left out, a column with nothing to show keeps its width and a hidden one is shown. `fitColumns(book, result, sheet, columns)` gives the same widths for a computed book without a session (`null`: the default width).
- Contract: `entryWidth(value, opts)`, the width Excel gives a column whose width nobody set when a number, a date or a formula's number is typed into it and would not show whole (`15.03.2026` in ru-RU): AutoFit's width when it is more, never narrower; General that only drops decimals keeps the column. Excel widens a column only for what is typed on the keyboard, not for values set through its object model.
- Contract: `ColumnInput.bestFit`, a width Excel fitted, which typing widens again; read and written in xlsx (`bestFit`).
- Text outside ASCII is measured by the advance widths of Carlito, a font metrically compatible with Calibri (SIL OFL 1.1, `THIRD_PARTY.md`), as Excel for Mac measures it for AutoFit.
- Measured with the engine's answers committed before each run: on the Mac 994 columns in four sets (`fit/autofit`, `typed`, `glyphs`, `held` in en-US and ru-RU), the widths and shown text as Excel's but for 67 known differences (`test/fit-known.ts`: Excel for Mac's own short date and month names where the engine follows Windows, Cyrillic within a point, a cell's own font, scripts outside the tables); on Windows 912 columns, the same columns fitted and kept, at Windows's own widths (below).

Known differences: Excel for Windows lays text out in pixels, so its widths differ from the Mac's the engine follows (a long date 21.71 there, 19.66 here); a cell's own font, size and boldness are not taken into account. Found on the way and not yet fixed: `writeXlsx` writes a cell without a format of its own in a formatted row or column without that format, and Excel shows it as General (46096 for a date).

## 0.14.0 (2026-10-10)

Internal release; not published to npm. Hidden sheets, and percent text read to the last bit.

Added:

- Contract: `SheetInput.hidden`, `true` for a hidden sheet (xlsx `state="hidden"`) and `'very'` for a very hidden one (`state="veryHidden"`), which Excel's window does not offer to show. Formulas read a hidden sheet as any other: `SUM('Jan:Mar'!B2)` over a hidden `Feb`, `SUBTOTAL` and `AGGREGATE` give what they gave. A book keeps one visible sheet at least (`InputError` otherwise); `inputHash` takes the state of a hidden sheet, and a book without one hashes as before.
- Contract: the structure operations `hideSheet` (`very` for very hidden) and `unhideSheet`. A session applies them without building its state again (`changes` is empty). Refused as Excel refuses: the last visible sheet is neither hidden nor deleted (`last_sheet`), a very hidden sheet is neither deleted nor moved (`very_hidden`, new), with messages in the book's locale.
- `readXlsx` and `writeXlsx` read and write the sheets' state; the writer makes the first visible sheet the active one, since Excel shows a hidden sheet that a file makes active. A file whose worksheets are all hidden is read with its first one visible, with an issue.
- Case sets measured in Excel for Mac and Windows with the engine's answers committed before each run: `hidden/sheets-held` (29 cases, 25 held), `sheets-held-2` (11, 6 held) and `sheets-held-3` (10, all held), both Excels alike; `entry/typed-pct-held` (47 per locale) and `pct-text` (76 per locale).

Changed:

- `moveSheet` before a hidden sheet puts the sheet before the next visible one, and after the last visible one when none follows, as both Excels move it. An end of a reference across sheets that moves stays its end until it passes the other end (`'Jan:Apr'` keeps Jan moved before Apr); it gave way to the next sheet inward.

Fixed:

- Percent text is read by its decimal text, as Excel reads it: `1,1%` typed in ru-RU is 0.011, where it was 0.011000000000000001, and so are `VALUE`, operators and numeric arguments of percent text (`1.255%` is 0.01255). A fraction with a percent (`1 1/2%`) is still divided. Results rounded to 15 digits did not show it; the stored input and the xlsx file did.
- `$1.1%` is text in en-US, not a number (`VALUE("$1.1%")` is `#VALUE!`), as Excel reads it, and `NUMBERVALUE` divides by 100 for each `%` in turn (`"1.1%%"` is 1.1/100/100).

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
