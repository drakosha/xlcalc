# Structure operations

Inserting and deleting rows and columns, moving and copying ranges, renaming and deleting sheets, with every reference in formulas and names adjusted as Excel adjusts it. This is the library layer for a grid interface: a pure function over `BookInput`. Session methods come in a second step (see the end).

## API

```ts
restructure(book: BookInput, op: StructureOp): RestructureResult

type StructureOp =
  | { op: 'insertRows' | 'deleteRows'; sheet: string; row: number; count?: number }          // 1-based row
  | { op: 'insertColumns' | 'deleteColumns'; sheet: string; column: string; count?: number } // letters
  | { op: 'moveRange'; from: RangeRef; to: CellRef }   // cut and paste, `to` is the top-left cell
  | { op: 'copyRange'; from: RangeRef; to: RangeRef }  // copy and paste, fill
  | { op: 'renameSheet'; sheet: string; name: string }
  | { op: 'deleteSheet'; sheet: string }

type RangeRef = { sheet: string; range: string }       // B2, B2:D9, B:D, 2:9
type RestructureResult =
  | { ok: true; book: BookInput; mapCell(ref: CellRef): CellRef | null }
  | { ok: false; code: 'off_sheet' | 'sheet_name' | 'last_sheet'; message: string }
```

- One function with the operation as data, rather than eight functions: an operation can be stored, sent from the interface to the server and replayed, and the session's method will take the same object. Rows are numbers and columns letters, as in A1 addresses; sheets are matched case-insensitively.
- The input book is not changed. Sheets, cells and names that the operation does not touch are shared with the input.
- `mapCell` says where a cell of the old book is now: a new address, the new sheet name, or null when the cell was deleted or overwritten. It is what a host needs for anything it keeps by address (comments, selection, links to source rows), and what the session will use to report changes.
- Input of the wrong shape (unknown sheet, row 0, a malformed range, unknown fields) throws `InputError`, as everywhere in the library. An operation Excel refuses returns `ok: false`: an interface shows the message, it is not a programming error. `off_sheet`: inserting would push a cell past row 1,048,576 or column XFD (Excel refuses for cells that are not empty; here any cell counts, a cell with only `format` or `meta` too, so nothing is dropped silently), or a move or copy does not fit on the sheet. `sheet_name`: an empty name or one another sheet has. `last_sheet`: the only sheet.

## Formula text

Only the references that change are rewritten, in place: the rest of the formula keeps its spelling (spaces, lower case), and a formula without a changed reference keeps its text byte for byte. A rewritten reference is written as Excel writes it: a range over every row as columns (`A1:A1048576` becomes `A:A`), over every column as rows, a new or changed sheet name quoted.

A formula that does not parse moves with its cell and is not rewritten. Excel has no such formulas (it rejects them on entry, or keeps the text as a value), so there is nothing to copy; rewriting would need a parse, and keeping the text lets the user see and fix what they wrote.

A lost reference is written as Excel writes it: `#REF!` in place of a reference without a sheet, `'S'!#REF!` when the reference named its sheet, `#REF!A5` (the reference after `#REF!`) when the sheet was deleted, `#REF!#REF!` for `'S'!#REF!` when `S` is deleted. The parser reads these forms (they also come in xlsx files) as `#REF!` errors and prints them back unchanged.

## Semantics

Each rule below is checked against Excel 16 for Mac (en-US) unless marked otherwise; see [Checked against Excel](#checked-against-excel).

**Insert rows (columns).** Cells from the row on move down, with their `format` and `meta`; the inserted rows are empty (Excel gives them the format of the row above; a book has no row formats, and a cell's `format` is not copied). Every reference to the sheet, from any sheet and from names, global and local, moves: absolute ones too. A range grows when rows are inserted between its first and last row, moves when they are inserted at or above its first row, and stays when they are inserted below its last row. A whole column (`B:B`) does not change; whole rows (`2:3`) behave as ranges. At the edge: a single cell pushed past the last row is `#REF!`; a range keeps ending on the last row (`A5:A1048576` becomes `A7:A1048576`), and is `#REF!` when its first row is pushed past it; a range over every row does not change.

**Delete rows (columns).** The cells in the deleted rows are removed, the cells below move up. A reference inside the deleted rows is `#REF!`; a range loses its deleted rows (`A3:A8` with rows 5–6 deleted is `A3:A6`; `A5:A6` with row 6 deleted is `A5:A5`, still a range) and is `#REF!` when all its rows are deleted. A range ending on the last row keeps ending there (deleting row 1,048,576 leaves `A5:A1048576` as it is), unless all its rows are deleted.

**Move (cut and paste).** The cells move with their `format` and `meta`; the cells they land on are replaced (cleared where the source cell is empty). A reference, single cell or range, entirely inside the moved cells follows them, to another sheet too, absolute or not. A reference entirely inside the overwritten cells is `#REF!`. A range with one end among the moved or the overwritten cells changes by these rules, along the axis of that end, when the cells cover the range's other axis and move along it:

- an end among the moved cells that moves away from the other end follows the cells (`SUM(A1:A10)`, cut `A10` to `A15`: `SUM(A1:A15)`);
- an end moving towards the other end, to cells outside the range (past the other end): no change;
- if the range held all the moved cells: it drops them and takes in where they land (`A4:A9`, cut `A4:A5` to `A7`: `A6:A9`);
- if the moved cells straddle the end: the end follows them, unless that leaves fewer than two rows or, for a move that does not overlap itself, a range inside the overwritten cells; then the range drops the moved cells; either way it takes in the overwritten cells beyond it;
- an end among the overwritten cells: the range drops them, unless it holds them all (`A1:A7`, cut `A4:A5` to `A7`: `A1:A6`; `A1:A8` stays).

Cells moved across the range's axis (sideways for a column range) do not change it, and a range over a whole column or row keeps it. Moved to another sheet, the cells leave a range with one end among them shrunk past them. A moved formula's references to cells that did not move keep pointing at them: a reference without a sheet gets the name of the sheet it pointed to. A moved formula that uses a name local to its old sheet gets that name defined on the new sheet too, when the new sheet sees no name of that name (neither its own nor a global one), as Excel does.

**Copy and paste, fill.** The copied cells replace the target's, with `format` and `meta`; empty source cells clear the cells they land on. Relative parts of the references move with the copy, absolute ones (`$`) stay, a whole column keeps its rows and a whole row its columns. A reference pushed past the sheet's edge is `#REF!` (`'S'!#REF!` when it names its sheet); a range is not clamped. A target that is a whole multiple of the source's size is tiled with copies (fill down is a copy of the first row over the rows below); any other target gets one copy at its top-left cell, as Excel's `Range.Copy`. Other formulas and names do not change. A formula copied to another sheet with a name local to the source sheet gets the name defined on the target sheet, its references to the source sheet pointed at the target sheet, under the same condition as for a move.

**Rename a sheet.** Every reference to the sheet, `'S'!#REF!` included, takes the new name, and so do the scope of its local names. A new name must be non-empty and differ from the other sheets' names, compared case-insensitively; renaming to another case is allowed. Excel also limits names to 31 characters and forbids `: \ / ? * [ ]` and an apostrophe at either end; the contract does not, so neither does `restructure` (a question below).

**Delete a sheet.** The sheet and its cells are removed. References to it become `#REF!` followed by the reference (`#REF!A5`, `SUM(#REF!$A$3:$A$8)`), in formulas and in global names and names local to other sheets; names local to the deleted sheet are removed.

## Checked against Excel

`npx tsx bench/coverage/run.ts struct <set>` applies each case's operation to its book in Excel (AppleScript, alerts off for the operation, every window that is not a book stops the run) and stores what every sheet shows in `A1:Z60` and the defined names in `docs/excel-struct/<set>.json`; `test/structure-excel.test.ts` compares `restructure` with them without Excel, formulas in Excel's notation with sheet names quoted. Known differences would be listed in `STRUCT_KNOWN` there; there are none.

| set | cases | what |
|---|---|---|
| `rows` | 26 | inserting and deleting rows and columns of a book with every kind of reference: relative, absolute, mixed, ranges, whole columns and rows, other sheets, a quoted sheet name, global and local names |
| `edge` | 10 | references at the last row and column, the whole sheet; the two insertions Excel refuses |
| `move` | 27 | cut and paste and copy and paste on the same book, onto referenced cells, whole columns and rows, to other sheets, fill down and right, tiles |
| `moves` | 20 | every range `Ai:Aj` and `Ai:Bj` of a 12-row window under 20 moves; the move rules were fitted on it |
| `sheets` | 9 | renaming (quotes, case, an apostrophe) and deleting sheets, references already lost |
| `checks` | 26 | held out: moves the fitting did not see, edge cases, copies with mixed `$` and local names, empty cells moved over values, a copy that does not fit; predicted before Excel ran |

The 118 cases read back 8,057 formulas and 311 names; about 1,000 formulas and 83 names were changed by the operations, 265 formulas to a lost reference. All match. The held-out set matched in 21 of 26 cases as predicted; the other five showed two rules that were wrong and are fixed (a range lying in the deleted rows is `#REF!` even when it ends on the last row; Excel does not carry a local name to a sheet where a global name of that name exists). One move rule (a straddled end does not follow into the overwritten cells of a non-overlapping move) rests on a single observation.

Not checked against Excel: a move that does not fit (refused like the copy), and moves onto another sheet beyond the two cases of `checks`.

## Limits

- A name whose reference is lost (`'S'!#REF!`) stays in the book, as in Excel, but the engine reports it in `nameIssues` and formulas using it give `#NAME?` where Excel gives `#REF!` (a change to the engine's name preparation, second step).
- No row or column formats, merged cells, conditional formats, validation or charts: the book has none.
- References the parser does not read (3D, structured, other workbooks) cannot occur in a valid formula; a formula that does not parse is not rewritten.

## Second step: session

- `session.restructure(op)`: applies the operation to the session's book and recomputes. Rebuilding the state from the new book is the simple way (a cold computation of the 3,400-row synthetic book takes 0.2–0.3 s); keeping compiled formulas whose text did not change and remapping cell keys by `mapCell` would bring it near an edit's cost. Its result: the changed values at their new addresses, plus `mapCell`, so the interface can move selection and comments.
- Contract: `StructureOp`, `RangeRef`, `RestructureResult` move to `contract.ts`, and `FormulaEngine` gains `restructure`.
- The engine: a name whose reference is lost evaluates to `#REF!` instead of being dropped.
