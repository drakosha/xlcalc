# Structure operations

Inserting and deleting rows and columns, moving and copying ranges, renaming and deleting sheets, with every reference in formulas and names adjusted as Excel adjusts it, and with the cells' styles, the columns and rows and the merged areas changed as Excel changes them. This is the library layer for a grid interface: a pure function over `BookInput`, and the same operation as a method of a session.

## API

```ts
restructure(book: BookInput, op: StructureOp): RestructureResult
session.restructure(op: StructureOp): SessionRestructureResult   // the same, plus `changes`

type StructureOp =
  | { op: 'insertRows' | 'deleteRows'; sheet: string; row: number; count?: number }          // 1-based row
  | { op: 'insertColumns' | 'deleteColumns'; sheet: string; column: string; count?: number } // letters
  | { op: 'moveRange'; from: RangeRef; to: CellRef }                       // cut and paste, `to` is the top-left cell
  | { op: 'copyRange'; from: RangeRef; to: RangeRef; withMeta?: boolean }  // copy and paste, fill
  | { op: 'renameSheet'; sheet: string; name: string }
  | { op: 'deleteSheet'; sheet: string }

type RangeRef = { sheet: string; range: string }       // B2, B2:D9, B:D, 2:9
type RestructureResult =
  | { ok: true; book: BookInput; mapCell(ref: CellRef): CellRef | null }
  | { ok: false; code: 'off_sheet' | 'sheet_name' | 'last_sheet' | 'merged'; message: string }
```

The types are in the contract (`xlcalc/contract`), and `FormulaEngine` has `restructure`.

- One function with the operation as data, rather than eight functions: an operation can be stored, sent from the interface to the server and replayed, and the session's method takes the same object. Rows are numbers and columns letters, as in A1 addresses; sheets are matched case-insensitively.
- The input book is not changed. Sheets, cells and names that the operation does not touch are shared with the input.
- `mapCell` says where a cell of the old book is now: a new address, the new sheet name, or null when the cell was deleted or overwritten. It is what a host needs for anything it keeps by address (comments, selection, links to source rows), and what the session uses to report changes.
- Input of the wrong shape (unknown sheet, row 0, a malformed range, unknown fields) throws `InputError`, as everywhere in the library. An operation Excel refuses returns `ok: false`: an interface shows the message, it is not a programming error. `off_sheet`: inserting would push a cell or a merged area past row 1,048,576 or column XFD (Excel refuses for cells that are not empty; here any cell counts, a cell with only `format`, `style` or `meta` too, so nothing is dropped silently), or a move or copy does not fit on the sheet. `sheet_name`: a name that breaks Excel's rules (empty, over 31 characters, one of `: \ / ? * [ ]`, an apostrophe at either end, `History` in any case) or one another sheet has. `last_sheet`: the only sheet. `merged`: a move whose source cuts a merged area, a copy that lands on part of one.

## Formula text

Only the references that change are rewritten, in place: the rest of the formula keeps its spelling (spaces, lower case), and a formula without a changed reference keeps its text byte for byte. A rewritten reference is written as Excel writes it: a range over every row as columns (`A1:A1048576` becomes `A:A`), over every column as rows, a new or changed sheet name quoted.

A formula that does not parse moves with its cell and is not rewritten. Excel has no such formulas (it rejects them on entry, or keeps the text as a value), so there is nothing to copy; rewriting would need a parse, and keeping the text lets the user see and fix what they wrote.

A lost reference is written as Excel writes it: `#REF!` in place of a reference without a sheet, `'S'!#REF!` when the reference named its sheet, `#REF!A5` (the reference after `#REF!`) when the sheet was deleted, `#REF!#REF!` for `'S'!#REF!` when `S` is deleted. A spill reference (`E2#`, `docs/spill.md`) moves as a reference to its cell and keeps its `#`, lost too: `#REF!#`, `'S'!#REF!#`, `#REF!E2#`. The parser reads these forms (they also come in xlsx files) as `#REF!` errors and prints them back unchanged. A name whose reference is lost stays in the book, as in Excel, and gives `#REF!` where it is used; it is no name issue.

## Semantics

Each rule below is checked against Excel 16 for Mac (en-US) unless marked otherwise; see [Checked against Excel](#checked-against-excel).

**Insert rows (columns).** Cells from the row on move down, with their `format`, `style` and `meta`. The inserted rows take the format of the row above, as Excel does by default: its height, style and format, and the style and format of each of its cells, without values; rows inserted above the first are plain (columns: the column on the left). Every reference to the sheet, from any sheet and from names, global and local, moves: absolute ones too. A range grows when rows are inserted between its first and last row, moves when they are inserted at or above its first row, and stays when they are inserted below its last row. A whole column (`B:B`) does not change; whole rows (`2:3`) behave as ranges. At the edge: a single cell pushed past the last row is `#REF!`; a range keeps ending on the last row (`A5:A1048576` becomes `A7:A1048576`), and is `#REF!` when its first row is pushed past it; a range over every row does not change.

**Delete rows (columns).** The cells in the deleted rows are removed, the cells below move up. A reference inside the deleted rows is `#REF!`; a range loses its deleted rows (`A3:A8` with rows 5–6 deleted is `A3:A6`; `A5:A6` with row 6 deleted is `A5:A5`, still a range) and is `#REF!` when all its rows are deleted. A range ending on the last row keeps ending there (deleting row 1,048,576 leaves `A5:A1048576` as it is), unless all its rows are deleted.

**Move (cut and paste).** The cells move with their `format`, `style` and `meta`; the cells they land on are replaced (cleared where the source cell is empty). A reference, single cell or range, entirely inside the moved cells follows them, to another sheet too, absolute or not. A reference entirely inside the overwritten cells is `#REF!`. A range with one end among the moved or the overwritten cells changes by these rules, along the axis of that end, when the cells cover the range's other axis and move along it:

- an end among the moved cells that moves away from the other end follows the cells (`SUM(A1:A10)`, cut `A10` to `A15`: `SUM(A1:A15)`);
- an end moving towards the other end, to cells outside the range (past the other end): no change;
- if the range held all the moved cells: it drops them and takes in where they land (`A4:A9`, cut `A4:A5` to `A7`: `A6:A9`);
- if the moved cells straddle the end: the end follows them, unless that leaves fewer than two rows or, for a move that does not overlap itself, a range inside the overwritten cells; then the range drops the moved cells; either way it takes in the overwritten cells beyond it;
- an end among the overwritten cells: the range drops them, unless it holds them all (`A1:A7`, cut `A4:A5` to `A7`: `A1:A6`; `A1:A8` stays).

Cells moved across the range's axis (sideways for a column range) do not change it, and a range over a whole column or row keeps it. Moved to another sheet, the cells leave a range with one end among them shrunk past them. A moved formula's references to cells that did not move keep pointing at them: a reference without a sheet gets the name of the sheet it pointed to. A moved formula that uses a name local to its old sheet gets that name defined on the new sheet too, when the new sheet sees no name of that name (neither its own nor a global one), as Excel does.

**Copy and paste, fill.** The copied cells replace the target's, with `format` and `style`, and without `meta` unless the operation has `withMeta: true` (a copy is a new cell: what a host keeps in `meta`, such as the ids of source rows, belongs to the original); empty source cells clear the cells they land on. Relative parts of the references move with the copy, absolute ones (`$`) stay, a whole column keeps its rows and a whole row its columns. A reference pushed past the sheet's edge is `#REF!` (`'S'!#REF!` when it names its sheet); a range is not clamped. A target that is a whole multiple of the source's size is tiled with copies (fill down is a copy of the first row over the rows below); any other target gets one copy at its top-left cell, as Excel's `Range.Copy`. Other formulas and names do not change. A formula copied to another sheet with a name local to the source sheet gets the name defined on the target sheet, its references to the source sheet pointed at the target sheet, under the same condition as for a move.

**Rename a sheet.** Every reference to the sheet, `'S'!#REF!` included, takes the new name, and so do the scope of its local names. A new name follows Excel's rules (1 to 31 characters, none of `: \ / ? * [ ]`, no apostrophe at either end, not `History` in any case; a book's sheets follow them too, else `InputError`) and differs from the other sheets' names, compared case-insensitively; renaming to another case is allowed.

**Delete a sheet.** The sheet and its cells are removed. References to it become `#REF!` followed by the reference (`#REF!A5`, `SUM(#REF!$A$3:$A$8)`), in formulas and in global names and names local to other sheets; names local to the deleted sheet are removed.

## Styles and layout

A book keeps styles as an xlsx file does (contract: `BookInput.styles`, `CellInput.style`, `SheetInput.columns`, `rows`, `merges`): a table of cell styles (font, solid fill, borders, alignment), a style number in a cell, and per sheet the columns and rows by line or span (`C`, `B:D`; `5`, `5:9`) with a width or height, a style and a number format, and the merged areas. A cell's number format stays its `format`. A cell shows its own style, else its row's, else its column's, and the same for the format, separately; widths are in characters, as xlsx stores them, heights in points. Computation does not read styles, and they are not in the values hash; `inputHash` has them when a book has any. Excel resolves an empty cell's style and format together, from its row if the row has either, else from its column: the two part only where a row with a style and no format (or a format and no style) crosses a column with the other, which an xlsx import spells out on the row.

Under the operations (checked against Excel, see below):

- **Insert and delete.** Columns and rows keep their entries as they move; a span holding the line before the inserted ones takes them in (they get its entry), a span reaching past the sheet's end is cut there, deleted lines leave their spans. The inserted lines' cells get the style and format of the cells of the line before them. A merged area changes as a range: it grows with lines inserted inside it, moves with lines inserted at or above its first line, stays when they come right after it (the new line is not merged), loses deleted lines, and is gone when one cell or none is left.
- **Move and copy.** A cell takes its own style and format along. Every pasted position shows the style and format its source position showed: where the rows and columns of the destination would show another (a cell moved out of a bold column, into a blue row), the cell gets the source's spelled out, `General` for no format and a plain style, which is added to the book's table when it has none. Whole rows or columns take their entries along (a move leaves the source lines plain, a copy repeats them on every tile). Merged areas inside the source go along; a copy also takes the part of a merged area its source cuts, when two cells or more; merged areas the pasted cells land on are gone (all of them for a move, those inside the target for a copy).
- **Refusals.** A move whose source cuts a merged area, and a copy that lands on part of one, are refused with `merged`, as Excel refuses them; a move that lands on part of one removes it, and a copy whose source cuts one copies the part. One difference is known: whole columns copied onto part of a merged area, which Excel pastes keeping the area (two observations), the engine refuses.
- **Sheets.** A renamed sheet keeps its layout; a deleted one takes it along.

## Session

`session.restructure(op)` applies the operation to the session's book and recomputes. The result is `restructure`'s (`ok`, `book`, `mapCell` or the refusal), where `book` is the session's book after the operation, plus `changes`: every cell whose value differs from the one its old cell had, by new address, where `mapCell` carries each old value to its new place and a cell nothing came to was empty. An interface that moves the values it shows by `mapCell` and applies `changes` has the new values; an insert into a block of data with no positional formulas reports none. A refused operation changes nothing; a live session takes a new clock, as for any recomputation.

The state is built again from the new book, the simple way: on the synthetic book (3,400 data rows, 3,907 formulas) a row inserted in the middle of the data sheet takes 190 ms (median of 8, a cold computation of the same book 197 ms; Apple M5 Pro, Node 26), within the 1.5 s budget of a cold computation, so the compiled formulas are not kept.

`book()` returns the session's book with every change so far, what a host saves as a version; its objects are the session's own and later calls change them. Styles and the layout change through the session, so its book does not part from the host's: a cell's style through `setCells` (an edit that keeps the cell's value and formula recomputes nothing and returns no changes), the style table through `setStyles(styles)` (every style in use must stay in it), columns, rows and merged areas through `setLayout(edits)`: `{ sheet, columns: 'B:D', set: { width: 20 } }` patches each of the lines (null removes a field; spans are split where an edit cuts them), `{ sheet, merges: [...] }` replaces the sheet's merged areas. Neither recomputes.

Other shapes were weighed and left:

- The host keeps its own book, calls the pure `restructure` and passes the new book to the session, which only recomputes. Two copies of the book that must be kept equal by every caller, and the session could not tell an operation from a new book, so it would rebuild and compare every cell.
- The session rewrites only what changed (keeping compiled formulas whose text did not change and moving cell keys by `mapCell`). Nearer an edit's cost, but the cost of the simple way is a cold computation, within budget, and the remapping would be a second implementation of every operation's cell mapping.
- Style edits as a whole new book (`setBook`). The host would rebuild the book for a width; a patch by lines and a style table set alone keep edits small and leave the values alone.
- Column widths and row heights as a map with one key per line. Simple to edit, but a style or width over the whole sheet is one span in xlsx and would be 16,384 keys here.

## Checked against Excel

`npx tsx bench/coverage/run.ts struct <set>` applies each case's operation to its book in Excel (AppleScript, alerts off for the operation, every window that is not a book stops the run) and stores what every sheet shows in `A1:Z60` and the defined names in `test/excel/struct/<set>.json`; `test/structure-excel.test.ts` compares `restructure` with them without Excel, formulas in Excel's notation with sheet names quoted. A book with styles, columns, rows, merged areas or number formats is written with them (ExcelJS), and the look of `A1:J24` is read back too: for each cell bold, the fill colour and the number format, the widths of the columns and heights of the rows that are not the standard ones, and the merged areas; the test compares the look the engine's book gives, with sizes as Excel 16 for Mac reports them (a width of w characters as (6w − 5)/6, a height in points as pixels, to a quarter). Known differences are listed in `STRUCT_KNOWN` there.

| set | cases | what |
|---|---|---|
| `rows` | 26 | inserting and deleting rows and columns of a book with every kind of reference: relative, absolute, mixed, ranges, whole columns and rows, other sheets, a quoted sheet name, global and local names |
| `edge` | 10 | references at the last row and column, the whole sheet; the two insertions Excel refuses |
| `move` | 27 | cut and paste and copy and paste on the same book, onto referenced cells, whole columns and rows, to other sheets, fill down and right, tiles |
| `moves` | 20 | every range `Ai:Aj` and `Ai:Bj` of a 12-row window under 20 moves; the move rules were fitted on it |
| `sheets` | 9 | renaming (quotes, case, an apostrophe) and deleting sheets, references already lost |
| `history` | 6 | renaming a sheet to `History`, `history`, `HISTORY` (Excel refuses, with an error and no window), `History1`, `My History`, `Histor` |
| `checks` | 26 | held out: moves the fitting did not see, edge cases, copies with mixed `$` and local names, empty cells moved over values, a copy that does not fit; predicted before Excel ran |
| `styles` | 46 | a styled book (cell styles and formats, a row at a height with a style and format, columns with widths and styles, merged areas): inserting and deleting rows and columns inside, at the edges and next to them, under styled rows; moving and copying ranges, whole rows and columns, merged areas; moves and copies that cut a merged area at either end; the style rules were fitted on it |
| `styles-checks` | 18 | held out: spans of rows and columns under inserts and deletes, merged areas tiled, cut and moved to another sheet, whole rows and columns moved and tiled, cells moved into a styled row and out of a styled column, whole columns copied onto part of a merged area; predicted before Excel ran |
| `spill`, `spill-more` | 8, 4 | spill references (`E2#`, `$E$2#`, from another sheet, in a name) under rows and columns inserted and deleted, the anchor cut and pasted, copies, a sheet renamed and deleted; dynamic array formulas, so the anchors spill (cutting an anchor to another sheet crashed Excel 16.113 twice and is not a case) |

The 118 cases without styles read back 8,057 formulas and 311 names; about 1,000 formulas and 83 names were changed by the operations, 265 formulas to a lost reference. All match. The held-out set matched in 21 of 26 cases as predicted; the other five showed two rules that were wrong and are fixed (a range lying in the deleted rows is `#REF!` even when it ends on the last row; Excel does not carry a local name to a sheet where a global name of that name exists). One move rule (a straddled end does not follow into the overwritten cells of a non-overlapping move) rests on a single observation.

The 64 cases with styles read back 294 formulas, 4,406 cells that are not plain, 380 sizes and 212 merged areas. The insert and delete rules were written after the first five answers of `styles` and matched its other 20 such cases; the move and copy rules were fitted on the rest. The held-out set matched in 17 of 18 cases as predicted; the comparison's own conversion of row heights was wrong (20 points read 26.75, not 26.67) and is fixed. The case off, whole columns copied onto part of a merged area, led to three more cases in `styles`: a range landing on the area's top-left cell or on another of its cells is refused, whole columns are pasted and the area stays. That is the one known difference (above).

Not checked against Excel: a move that does not fit (refused like the copy), moves onto another sheet beyond the cases of `checks` and `styles-checks`, deleting lines at the sheet's end under a span that reaches it, fonts other than bold, borders and alignment (they move with the style number), what a whole-column copy onto a merged area does to the area's cells.

## Limits

- No hidden rows or columns, outline levels, conditional formats, validation, comments or charts: the book has none.
- References the parser does not read (3D, structured, other workbooks) cannot occur in a valid formula; a formula that does not parse is not rewritten.
- Whole rows or columns copied onto part of a merged area are refused (Excel pastes them).

## Questions settled

Proposals of the first step, implemented so that dropping any one is a local change:

- Copies have no `meta` unless `withMeta: true`; a move keeps it (`structure/copy.ts`, one function).
- Sheet names follow Excel's rules in the input and in `renameSheet` (`sheetNameProblem` in `input.ts`), the reserved name `History` too, in any case (Excel refuses to rename a sheet so, the set `history`; `History1` and `My History` are taken).
- An insert is refused when any cell would leave the sheet, even one with only a format, style or meta (stricter than Excel), and a merged area as well.
- Inserted rows take the format of the row above, and columns of the column on the left, as Excel does by default (`formatBefore` in `structure/layout.ts`, one call in `shift.ts`, and the span growth in `shiftLines`).
