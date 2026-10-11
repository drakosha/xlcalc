# Structure operations

Inserting and deleting rows and columns, moving and copying ranges, renaming, deleting, inserting, moving, hiding and showing sheets. References in formulas and names change as Excel changes them, and so do cell styles, column and row settings, merged areas, notes, hyperlinks, frozen panes, conditional formats, rules of data validation and outline levels. It is the layer a grid interface needs: a pure function over `BookInput`, and the same operation as a session method.

The rules below were measured in Excel 16 for Mac (en-US).

## API

```ts
editStructure(book: BookInput, op: StructureOp): StructureResult
session.editStructure(op: StructureOp): SessionStructureResult   // the same, plus `changes`

type StructureOp =
  | { op: 'insertRows' | 'deleteRows'; sheet: string; row: number; count?: number }          // 1-based row
  | { op: 'insertColumns' | 'deleteColumns'; sheet: string; column: string; count?: number } // letters
  | { op: 'moveRange'; from: RangeRef; to: CellRef }                       // cut and paste, `to` is the top-left cell
  | { op: 'copyRange'; from: RangeRef; to: RangeRef; withMeta?: boolean }  // copy and paste, fill
  | { op: 'renameSheet'; sheet: string; name: string }
  | { op: 'deleteSheet'; sheet: string }
  | { op: 'insertSheet'; name: string; before?: string }  // an empty sheet, at the end without `before`
  | { op: 'moveSheet'; sheet: string; before?: string }   // to the end without `before`
  | { op: 'hideSheet'; sheet: string; very?: boolean }   // very hidden with `very`
  | { op: 'unhideSheet'; sheet: string }

type RangeRef = { sheet: string; range: string }       // B2, B2:D9, B:D, 2:9
type StructureResult =
  | { ok: true; book: BookInput; mapCell(ref: CellRef): CellRef | null }
  | { ok: false; code: 'off_sheet' | 'sheet_name' | 'last_sheet' | 'very_hidden' | 'merged' | 'filtered' | 'array'; message: string }
```

The types are in `xlcalc/contract`.

An operation is plain data, so an interface can store it, send it to a server and replay it. Rows are numbers and columns are letters, as in A1 addresses. Sheet names are matched without regard to case.

The input book is not changed. Sheets, cells and names the operation does not touch are shared with it.

`mapCell` tells where a cell of the old book is now: a new address and sheet name, or `null` if the cell was deleted or overwritten. A host uses it for anything it keeps by address, such as comments, the selection or links to source rows.

Input of the wrong shape (an unknown sheet, row 0, a malformed range, unknown fields) throws `InputError`. An operation that Excel would refuse returns `ok: false` with a message for the user:

| code | when |
|---|---|
| `off_sheet` | an insert would push a cell or a merged area past row 1,048,576 or column XFD, or a move or copy does not fit on the sheet |
| `sheet_name` | the new name breaks Excel's rules or another sheet has it |
| `last_sheet` | deleting or hiding the last visible sheet (the only sheet too) |
| `very_hidden` | deleting or moving a very hidden sheet |
| `merged` | a move whose source cuts a merged area, or a copy that lands on part of one |
| `filtered` | a copy that left out rows a filter hid, pasted over hidden rows of a filtered sheet |
| `array` | lines inserted inside the area of an array formula of the old kind, part of one deleted or moved, a paste over part of one |

Excel refuses an insert only when a non-empty cell would leave the sheet. Here any cell counts, even one with only a format, style or `meta`, so nothing is dropped silently.

## Formula text

Only the references that change are rewritten, in place. The rest of the formula keeps its spelling, spaces and lower case included, and a formula without a changed reference keeps its text byte for byte. A rewritten reference is written as Excel writes it: a range over every row as columns (`A1:A1048576` becomes `A:A`), a range over every column as rows, a new or changed sheet name in quotes.

A formula that does not parse moves with its cell as written. Excel never has such formulas, and keeping the text lets the user see and fix it.

A lost reference is written the way Excel writes it:

| reference | after the loss |
|---|---|
| `A5`, its cells deleted | `#REF!` |
| `'S'!A5`, its cells deleted | `'S'!#REF!` |
| `'S'!A5`, sheet S deleted | `#REF!A5` |
| `'S'!#REF!`, sheet S deleted | `#REF!#REF!` |
| `E2#` (a spill reference), E2 deleted | `#REF!#`, or `'S'!#REF!#` with a sheet |
| `'S'!E2#`, sheet S deleted | `#REF!E2#` |

The parser reads all these forms, which also occur in xlsx files, as `#REF!` errors and prints them back unchanged. A name whose reference is lost stays in the book, as in Excel, and gives `#REF!` where it is used.

## Inserting rows and columns

Cells from the inserted row on move down, with their `numberFormat`, `style` and `meta`. The inserted rows copy the row above, as Excel does by default: its height, style and format, and the style and format of each of its cells, without values. Rows inserted above row 1 are plain. Columns work the same way with the column on the left.

Every reference to the sheet moves, absolute references included, from any sheet and from global and local names.

- A range grows when rows are inserted between its first and last row, moves when they are inserted at or above its first row, and stays when they are inserted below it.
- A whole column (`B:B`) does not change. Whole rows (`2:3`) behave like ranges.
- At the edge of the sheet, a single cell pushed past the last row becomes `#REF!`. A range that ends on the last row keeps ending there (`A5:A1048576` becomes `A7:A1048576`) and becomes `#REF!` when its first row is pushed past it. A range over every row does not change.

## Deleting rows and columns

The cells in the deleted rows are removed and the cells below move up. A reference inside the deleted rows becomes `#REF!`. A range loses its deleted rows: `A3:A8` with rows 5 and 6 deleted is `A3:A6`, and `A5:A6` with row 6 deleted is `A5:A5`, still a range. A range whose rows are all deleted becomes `#REF!`. A range ending on the last row keeps ending there unless all its rows are deleted.

## Moving a range (cut and paste)

The cells move with their `numberFormat`, `style` and `meta`, and replace the cells they land on; an empty source cell clears its target. A reference that lies entirely inside the moved cells follows them, to another sheet too, absolute or not. A reference entirely inside the overwritten cells becomes `#REF!`.

A range with one end among the moved or overwritten cells changes along that end's axis, when the cells cover the range's other axis and move along it:

- An end among the moved cells that moves away from the other end follows them: cutting `A10` to `A15` turns `SUM(A1:A10)` into `SUM(A1:A15)`.
- An end that moves towards the other end, to cells outside the range, does not change it.
- A range that held all the moved cells drops them and takes in where they land: cutting `A4:A5` to `A7` turns `A4:A9` into `A6:A9`.
- When the moved cells straddle an end, the end follows them, unless that leaves fewer than two rows or, for a move that does not overlap itself, a range inside the overwritten cells. Then the range drops the moved cells. Either way it takes in the overwritten cells beyond it.
- An end among the overwritten cells drops them, unless the range holds them all: cutting `A4:A5` to `A7` turns `A1:A7` into `A1:A6` and leaves `A1:A8` as it is.

Cells moved across a range's axis, such as sideways for a column range, do not change it, and a range over a whole column or row stays whole. When the cells move to another sheet, a range with one end among them shrinks past them.

A moved formula's references to cells that did not move keep pointing at them; a reference without a sheet gets the name of the sheet it pointed to. A moved formula that uses a name local to its old sheet gets that name defined on the new sheet too, if the new sheet sees no name of that name, its own or a global one.

## Copying a range (copy and paste, fill)

The copied cells replace the target's, with `numberFormat` and `style`. They do not take `meta` unless the operation has `withMeta: true`: a copy is a new cell, and what a host keeps in `meta`, such as the ids of source rows, belongs to the original. Empty source cells clear the cells they land on.

Relative parts of references move with the copy and absolute ones (`$`) stay. A whole column keeps its rows and a whole row its columns. A reference pushed past the sheet's edge becomes `#REF!`, or `'S'!#REF!` when it names its sheet; a range is not clamped.

A target that is a whole multiple of the source's size is tiled with copies, so fill down is a copy of the first row over the rows below. Any other target gets one copy at its top-left cell, as Excel's `Range.Copy` does. Other formulas and names do not change. A formula copied to another sheet that uses a name local to the source sheet gets the name defined on the target sheet, with its references to the source sheet pointed at the target, on the same condition as for a move.

`copyFormula(text, from, to, book?)` gives the text of one formula copied from cell `from` to cell `to` (each `{ sheet, cell }`) by the same rules, for an interface that shows a paste before it makes it. Given the book (a `BookInput`, or its sheets' names and tables alone), a formula copied out of a table names the table in its references to it, as a copy does (`[@Price]` becomes `Sales[[#This Row],[Price]]`); without it the formula's references keep the sheet names they are written with.

```js
import { copyFormula } from 'xlcalc'

console.log(copyFormula('A1+$B$2+SUM(C1:C3)', { sheet: 'S', cell: 'D1' }, { sheet: 'S', cell: 'E3' })) // → B3+$B$2+SUM(D3:D5)
console.log(copyFormula('A1*2', { sheet: 'S', cell: 'B2' }, { sheet: 'S', cell: 'A1' })) // → #REF!*2
```

What a copy does beyond one formula's text is `copyRange`'s alone: it leaves out the rows a filter hid, copies part of an old-kind array formula's area as arrays of their own, points the formulas of a table copied whole at the copy, and turns a formula pasted on a table's header into the header's text.

## Renaming and deleting sheets

Renaming a sheet renames every reference to it, `'S'!#REF!` included, and the scope of its local names. The new name must follow Excel's rules (1 to 31 characters, none of `: \ / ? * [ ]`, no apostrophe at either end, not `History` in any case) and differ from the other sheets' names, compared without regard to case. Renaming to another case is allowed. A book whose sheet names break these rules is an `InputError`.

Deleting a sheet removes it with its cells. References to it become `#REF!` followed by the reference (`#REF!A5`, `SUM(#REF!$A$3:$A$8)`), in formulas, in global names and in names local to other sheets. Names local to the deleted sheet are removed.

## Inserting and moving sheets

`insertSheet` adds an empty sheet before another one or at the end. `moveSheet` moves a sheet. Neither changes a formula, except a reference across sheets whose end moves. Such a reference covers the sheets between its ends in the book's order, so a sheet inserted or moved between them is read and one moved out is not. An end that moves stays the end while it does not pass the other end: with Jan, Feb, Mar and Apr, Jan moved before Apr leaves `'Jan:Apr'` as it is, now over Jan and Apr, while `'Jan:Mar'` becomes `'Feb:Mar'`. [References](references.md) has the rules for its ends.

## Hidden sheets

A sheet is hidden by `hidden: true`, or very hidden by `hidden: 'very'`, as Excel's Hide and its object model's `xlSheetVeryHidden` hide it (an xlsx file's `state`). `hideSheet` hides a sheet (`very: true` for very hidden) and `unhideSheet` shows it. Formulas read hidden sheets as any other: `SUM('Jan:Mar'!B2)` with Feb hidden takes in Feb, and `SUBTOTAL` leaves out hidden rows alone, never a hidden sheet.

A book keeps one visible sheet at least, as Excel does: hiding or deleting the last visible sheet is refused with `last_sheet`, and a book whose every sheet is hidden is an `InputError`. A hidden sheet is renamed, deleted and moved as any other and keeps its state; a very hidden one is renamed alone, and deleting or moving it is refused with `very_hidden`, as Excel's object model refuses both.

Excel moves a sheet by the visible sheets. Moved before a hidden sheet, it goes before the next visible sheet; with none after it, and moved to the end past hidden sheets, it goes after the last visible sheet. With Jan, Feb (hidden) and Mar, Mar moved before Feb stays where it is, and with Mar hidden at the end, Jan moved to the end goes between Feb and Mar. An inserted sheet goes where it is asked, before a hidden sheet too.

```js
import { editStructure } from 'xlcalc'

const book = { sheets: [{ name: 'Data', cells: {} }, { name: 'Keys', cells: {}, hidden: 'very' }, { name: 'Report', cells: {} }] }
const r = editStructure(book, { op: 'hideSheet', sheet: 'Data' })
console.log(r.ok && r.book.sheets.map((s) => s.hidden)) // → [ true, 'very', undefined ]
console.log(r.ok && editStructure(r.book, { op: 'hideSheet', sheet: 'Report' }).code) // → last_sheet
console.log(editStructure(book, { op: 'deleteSheet', sheet: 'Keys' }).code) // → very_hidden
```

## Tables

Tables move with their cells and grow and shrink with the rows and columns inserted and deleted inside them; structured references follow them by name. Inserted rows get the formulas of calculated columns and inserted columns a name (`Column1`); references to a deleted column or table become `#REF!`. Rows or columns inserted into two tables at once, deleted from two tables unless both go whole, and a header row deleted alone are refused with `table`. `addTable`, `renameTable`, `deleteTable` and `convertTable` change tables themselves. [Tables](tables.md) has the rules.

## Names, references across sheets and old-kind arrays

Names' formulas change like cells' formulas: each part of a union on its own, a reference without `$` not at all under inserted and deleted lines ([Names](names.md)). A reference across sheets stays under rows and columns and changes with its end sheets ([References](references.md)). The area of an array formula of the old kind moves with its cell; an operation that would change part of it is refused with `array` ([Formulas of the old kind](legacy.md)).

## Styles and layout

A book keeps styles the way an xlsx file does (`BookInput.styles`, `CellInput.style`, `SheetInput.columns`, `rows`, `merges`): a table of cell styles (font, solid fill, borders, alignment), a style number in each cell, and per sheet the columns and rows by line or span (`C`, `B:D`; `5`, `5:9`) with a width or height, a style and a number format, plus the merged areas. A cell's number format is its `format`. A cell shows its own style, else its row's, else its column's, and the same for the format. Widths are in characters as xlsx stores them, heights in points. Computation does not read styles, and they are not in the values hash; `inputHash` includes them when a book has any.

Excel resolves an empty cell's style and format together: from its row if the row has either, else from its column. The two can differ only where a row with a style but no format (or the other way round) crosses a column with the other; an xlsx reader spells such a case out on the row.

**Inserting and deleting.** Columns and rows keep their settings as they move. A span that holds the line before the inserted ones takes them in. A span that would reach past the sheet's end is cut there, and deleted lines leave their spans. The cells of inserted lines get the style and format of the cells of the line before them.

A merged area changes like a range: it grows with lines inserted inside it, moves with lines inserted at or above its first line, stays when lines are inserted right after it (the new line is not merged), loses deleted lines, and disappears when one cell or none is left.

At the end of the sheet, a span of columns that reaches XFD keeps reaching it when columns are deleted, as xlsx stores such a span (`<col max=16384>`). If the deleted columns include XFD, the columns that come in at the end take the span's style and format but not its width (seen once). Rows are stored one by one and do not reach on. A merged area that ends on the last row or column keeps ending there when a deletion reaches the end of the sheet, unless all of it is deleted.

**Moving and copying.** A cell takes its own style and format along. Every pasted cell shows the style and format its source showed: where the target's rows and columns would show something else (a cell moved out of a bold column into a blue row), the cell gets the source's style spelled out, with `General` for no format and a plain style that is added to the book's table if missing. Whole rows or columns take their settings along; a move leaves the source lines plain, and a copy repeats them on every tile.

Merged areas inside the source go along. A copy also takes the part of a merged area that its source cuts, when that part has two cells or more. Merged areas that pasted cells land on are removed: all of them for a move, those inside the target for a copy.

**Refusals.** A move whose source cuts a merged area and a copy that lands on part of one are refused with `merged`, as Excel refuses them. A move that lands on part of a merged area removes it, and a copy whose source cuts one copies the part.

A renamed sheet keeps its layout, and a deleted sheet takes it along.

## Notes, hyperlinks and frozen panes

A cell's `note` and `link` go with the cell, as its style does:

- Inserted rows and columns get none, even where the line before them has some.
- A deleted cell takes them with it.
- A move takes them along. The cells it lands on lose theirs, even where the moved cell has none.
- A copy copies them to every cell it pastes, and a pasted cell without them clears the target's.
- A hyperlink's `location` stays as written when its sheet is renamed or deleted, or when the rows it points to move. Excel does not follow it either.

Frozen panes (`frozen`) keep the lines before their split. Rows inserted before the split push it down: with two rows frozen, a row inserted at row 1 or 2 makes three, one inserted at row 3 leaves two. Deleted rows before the split draw it up, but one row stays frozen at least: deleting rows 1 and 2 of two frozen ones leaves row 1 frozen. Columns work the same way. A move or a copy does not change them, and a renamed sheet keeps them.

The parts sets `struct-held` and `frozen-held` measured these rules in Excel for Mac and for Windows, with the engine's answers committed before Excel ran them ([docs/xlsx.md](xlsx.md#notes-hyperlinks-and-frozen-panes)).

## Conditional formats

A rule's ranges and formulas change as Excel changes them ([docs/conditional-formats.md](conditional-formats.md#structure-operations)):

- Inserted lines that follow a line of a rule join it, as they take the format of the line before them; a rule after them moves.
- Deleted lines leave its ranges, and a rule with nothing left is gone. A colour scale, a data bar, an icon set, a top, average or duplicate rule that loses lines goes after the other rules.
- A move takes its cells' rules along and the cells it lands on keep theirs; to another sheet the moved part is a rule there, first. A rule whose formula the move rewrote goes after the others.
- A copy pastes its cells' rules as new rules, first, and the pasted cells lose theirs; a fill down a rule's own columns grows the rule.
- Formulas change as the formula of the rule's first cell (the top-left of its cells) would; a reference to a deleted sheet is `#REF!`.

The cond sets `struct-held`, `held` and `held-2` measured these rules in Excel for Mac and for Windows, with the engine's answers committed before Excel ran them.

## Data validation and outlines

A rule of data validation (`SheetInput.validations`, [docs/validation.md](validation.md)) is a range of cells and formulas written for its first cell. Its ranges change as references do, with one difference at their end:

- Lines inserted after a range's first line grow it, lines inserted at or before its first line move it, and lines inserted right after its last line join it: the inserted cells take the rule of the cells before them, as they take their format. With a rule over `B2:B6`, a row inserted at row 4 or at row 7 makes it `B2:B7`, one inserted at row 2 makes it `B3:B7`. Columns work the same way with the column on the left.
- Deleted lines leave the ranges; a range that loses every line is gone, and a rule that loses every range.
- A move takes the rules of the moved cells along. The cells it lands on keep their own rules where the moved cells have none, and take the moved ones where they have some. A move from `A1`, which has no rule, onto a cell of a whole-number rule leaves that cell's rule.
- A copy pastes the rules of the copied cells on every tile, and a pasted cell without a rule clears the target's rule, as Excel's Paste All does.
- References in a rule's formulas change as a cell's formulas do: a list `'L'!$A$1:$A$4` follows rows inserted on L, becomes `'Lists'!$A$1:$A$4` when L is renamed and `#REF!$A$1:$A$4` when it is deleted. A rule whose first cell moves or goes has its formulas written again for its new first cell: deleting row 2 of a rule `D2:D6` with `D2<A2` leaves `D2:D5` with `D2<A2`, the old D3's formula. A rule moved to another sheet moves its formulas as a copy moves them: `B2>A2` moved from B2 of S to D2 of T is `D2>C2` there, while a move on its own sheet keeps the references to the cells that did not move (`D2<A2` moved to D20 is `D20<A2`).

An outline is the outline levels of rows and columns (`outlineLevel`, 1 to 7) with their collapsed marks (`collapsed`), and the sheet's summary settings (`SheetInput.outline`, below and right by default). A collapsed group's lines are hidden lines, which `SUBTOTAL` 101 to 111 leaves out.

- Inserted lines take the outline level of the line before them, as they take its height, and are shown. Lines inserted into a collapsed group after one of its hidden lines open it: its summary line, past the group (before it with summaries above or left), loses its collapsed mark. Lines inserted at the summary line, or above the group, leave it collapsed.
- Deleted lines take their levels and marks along; moved and copied whole rows and columns carry them, as they carry their other settings.

The valid sets `struct-held` and `struct-held-2` measured these rules in Excel for Mac and for Windows, with the engine's answers committed before Excel ran them ([docs/validation.md](validation.md#how-it-was-checked)).

## Hidden rows and filtered sheets

A row or a column is hidden by `hidden: true` in `rows` or `columns`, or by a height or width of 0, as Excel takes it. A sheet is marked `filtered: true` when a filter of it has criteria, its autofilter or a table's. Excel then treats every hidden row of the sheet as one the filter hid, whatever hid it. In an xlsx file the criteria decide, not the attribute `filterMode`: Excel takes a sheet as filtered when its autofilter has criteria and the attribute is missing, does not when the attribute stands beside an autofilter without criteria, and writes no attribute for a table's filter ([docs/xlsx.md](xlsx.md)).

On any sheet:

- Inserted rows and columns are shown, even inside a hidden block. They copy every other setting of the line before them, its size too: a row inserted under a row of height 0 has height 0 and so is hidden, under a row hidden at height 30 it is shown at 30 (the set `zero-held`, both Excels).
- Deleted lines take their `hidden` along, and moved or copied whole rows and columns carry it.
- Cells move and copy without regard to `hidden` on a sheet that is not filtered. A copy of `A2:C16` with rows 4 and 7 hidden by hand copies all 15 rows.

On a filtered sheet Excel acts on the visible rows:

- Deleting rows deletes the visible rows among them. Deleting rows 6 to 9 when rows 7 and 8 are hidden deletes rows 6 and 9.
- Inserting rows inserts as many rows as are visible among them, before the first. Inserting two rows at row 8, when row 8 is hidden and row 9 is shown, inserts one row.
- A copy whose source covers hidden rows takes the visible rows only, one after another, and pastes their values instead of formulas. A copy of `A2:C16` with rows 4, 7, 8 and 11 hidden pastes 11 rows of values.
- A paste of several cells over hidden rows goes into the visible rows. Each run of visible rows takes the copy as a target of its own. A copy of `D2:E2` onto `D3:E16` fills the visible rows and leaves the hidden ones as they were.
- A copy that left rows out cannot be pasted over hidden rows. The operation is refused with `filtered`, as Excel refuses it.
- A selection of hidden rows only is taken whole.

The values a copy pastes come from the computed book. `session.editStructure` takes the session's values. `editStructure` computes the book first, at the current moment, which matters only for volatile formulas.

`copyRange` over a filtered sheet fills its visible rows, as Excel's Fill Down does. Excel's fill handle fills the hidden rows too.

## Sessions

`session.editStructure(op)` applies the operation to the session's book and recomputes. The result is that of `editStructure`, with `book` being the session's book after the operation, plus what every edit of a session answers ([Editing a grid](edits.md#what-an-edit-answers)): `moved`, how the addresses moved, and `changes`, every cell whose value or formula text differs from those of its old cell. An interface that moves its window by `moved` (or each cell by `mapCell`) and applies `changes` has the new values and formulas. An insert into a block of data lists the formulas over it, which read one row more, and no value. A refused operation changes nothing.

The session rebuilds its state from the new book. On the synthetic book (3,400 rows, 3,907 formulas) inserting a row in the middle of the data takes 190 ms, about as long as a cold computation of the same book, well within the 1.5 s budget.

`getBook()` returns the session's book with all changes so far, which is what a host saves as a version. Its objects belong to the session and change with later edits, so serialize it right away.

Styles and layout change through the session so that its book stays the host's book:

- a cell's style, note and hyperlink through `setCells` or `patchCells`; an edit that keeps a cell's value and formula recomputes nothing and lists the cell with its new style or number format;
- the style table through `setStyles(styles)`; every style in use must stay in it;
- columns, rows and merged areas through `setLayout(edits)`: `{ sheet, columns: 'B:D', set: { width: 20 } }` patches each of the lines (`null` removes a field, spans are split where an edit cuts them), and `{ sheet, merges: [...] }` replaces the sheet's merged areas;
- the filter mode through `setLayout`: `{ sheet, filtered: true }`;
- hidden sheets through `editStructure`: `hideSheet` and `unhideSheet` recompute nothing and do not build the state again;
- frozen panes through `setLayout`: `{ sheet, frozen: { rows: 1 } }`, and `{ sheet, frozen: null }` unfreezes them;
- conditional formats through `setLayout`: `{ sheet, conditionalFormats: [...] }` replaces the sheet's rules;
- rules of data validation through `setLayout`: `{ sheet, validations: [...] }` replaces the sheet's rules;
- outlines through `setLayout`: a group is the lines' `outlineLevel`, collapsing it hides its lines and marks its summary line `collapsed` (`{ sheet, rows: '3:5', set: { hidden: true } }`, `{ sheet, rows: '6', set: { collapsed: true } }`), and `{ sheet, outline: { summaryBelow: false } }` sets where the summaries are.

Only two of these recompute values. A merged area added or removed over an array's spill range recomputes that array. Rows hidden or shown, and a new filter mode, recompute the `SUBTOTAL` and `AGGREGATE` formulas over the rows concerned and what depends on them.

## How it was checked

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

Not checked against Excel: a move that does not fit (it is refused like a copy), most moves to another sheet, fonts other than bold, borders and alignment (they move with the style number).

Copies were applied through AppleScript's `copy range` with a destination, Excel's `Range.Copy`. On a filtered sheet it pasted values. Copying and pasting in Excel's window was not measured.

Hidden sheets were measured by the step `hidden` in Excel for Mac and for Windows (`npx tsx bench/coverage/run.ts hidden <set>`, `bench/windows/run.ts hidden <set>`, `test/hidden-sheets-excel.test.ts`): sheets hidden and shown through the object model (`Visible`), renamed, deleted, inserted and moved beside hidden ones, with references across sheets over them. Each set's answers were predicted before Excel ran it: `sheets-held` held in 25 of 29 cases, the misses being the very hidden sheet that neither Excel deletes or moves and a sheet moved before a hidden one, which went after it; `sheets-held-2` in 6 of 11, the misses being moves past hidden sheets with none visible after them and ends of references across sheets that moved without passing the other end; `sheets-held-3` in 10 of 10. Both Excels answered alike ([docs/xlsx.md](xlsx.md#hidden-sheets-2)).

## Known differences and limits

- Whole columns copied onto part of a merged area: Excel pastes them and keeps the area, the engine refuses with `merged` (two observations).
- Excel's fill handle fills the hidden rows of a filtered sheet. `copyRange` fills the visible rows only, as Fill Down does.
- One move rule rests on a single observation: a straddled end does not follow into the overwritten cells of a move that does not overlap itself.
- Charts are not part of a book. Threads of comments are notes in a book ([docs/xlsx.md](xlsx.md#threads-of-comments)).
- References to other workbooks are not read, so they never occur in a valid formula.
