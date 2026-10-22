# Structure operations

Inserting and deleting rows and columns, moving and copying ranges, merging and unmerging cells, and renaming, deleting, inserting, moving, hiding and unhiding sheets. References in formulas and names change as Excel changes them, and so do cell styles, column and row settings, merged areas, notes, hyperlinks, frozen panes, conditional formats, data validation rules and outline levels. This is the layer a grid interface needs: a pure function over `BookInput`, and the same operation as a session method.

The rules below were verified in Excel 16 for Mac (en-US), and those for merging cells in Excel for Windows.

## In short

`editStructure(book, op)` inserts or deletes rows and columns, moves (cut and paste) or copies (copy and paste, fill) a range, renames or deletes a sheet, and adjusts every reference in formulas and names on every sheet as Excel adjusts it: ranges grow and shrink, lost references become `#REF!`, and absolute references move with inserted rows. It returns the new workbook and `mapCell`, which tells where a cell of the old workbook is now (null for a cell that was deleted or overwritten, or that belongs to a sheet the old workbook does not have); the input workbook is not changed. Treat the returned workbook as read-only: the calls you give it to next (`editStructure`, `compute`, a session) skip checking the cells `editStructure` made, so change a copy of it (`structuredClone`) instead of changing it in place.

```js
import { editStructure } from 'xlcalc'

const book = {
  sheets: [{
    name: 'Sales',
    cells: {
      A1: { value: 120 }, A2: { value: 80 }, A3: { value: 45 },
      C1: { formula: 'SUM(A1:A3)' }, C2: { formula: 'A3*2' },
    },
  }],
}

const ins = editStructure(book, { op: 'insertRows', sheet: 'Sales', row: 2 })
console.log(ins.book.sheets[0].cells.C1.formula, ins.book.sheets[0].cells.C3.formula) // → SUM(A1:A4) A4*2
console.log(ins.mapCell({ sheet: 'Sales', cell: 'A3' })) // → { sheet: 'Sales', cell: 'A4' }

const del = editStructure(book, { op: 'deleteRows', sheet: 'Sales', row: 3 })
console.log(del.book.sheets[0].cells.C1.formula, del.book.sheets[0].cells.C2.formula) // → SUM(A1:A2) #REF!*2
const last = editStructure(book, { op: 'deleteSheet', sheet: 'Sales' })
console.log(last.ok, last.code) // → false last_sheet
```

The operations are:
- `insertRows` and `deleteRows` (`row`, a row number, and `count`), `insertColumns` and `deleteColumns` (`column`, letters, and `count`);
- `moveRange` (`from: { sheet, range }`, `to: { sheet, cell }`) and `copyRange` (`from` and `to` as `{ sheet, range }`; a destination that is a whole multiple of the source is tiled, as in fill down; copies have no `meta` unless `withMeta: true`);
- `renameSheet` (`name`, by Excel's rules for sheet names), `deleteSheet`, `insertSheet` (`name`, `before`), `moveSheet` (`sheet`, `before`; Excel positions a sheet relative to the visible ones), `hideSheet` (`sheet`, `very` for a very hidden sheet) and `unhideSheet` (`sheet`);
- for tables, `addTable` (`sheet`, `table`), `renameTable` (`table`, `name`), `deleteTable` and `convertTable` (`table`; [docs/tables.md](tables.md));
- Excel's outline commands `group` and `ungroup` (`sheet`, `rows` or `columns`), `hideDetail` and `showDetail` (`sheet`, `row` or `column`) and `showLevels` (`sheet`, `rows`, `columns`; [docs/outline.md](outline.md));
- Excel's Merge Cells and Unmerge Cells, `mergeCells` (`sheet`, `range`, `across` for Merge Across, `center` for Merge & Center) and `unmergeCells` (`sheet`, `range`): a merge keeps the top-left cell, or moves the area's first value into it when it is empty, and clears the rest ([docs/structure.md](structure.md#merging-cells));
- Excel's fill handle, `fillRange` (`sheet`, `range`, and `to`, which contains `range` and extends beyond it; `type` is `default`, `series` or `copy`): copies of the source, with numbers, dates, times, text with a number, and names of days and months continuing as Excel continues them ([docs/structure.md](structure.md#the-fill-handle)).

Only the references that change are rewritten, in place; a formula that does not parse moves with its cell as written. An operation Excel rejects returns `{ ok: false, code, message }`, with the message in the workbook's locale (`Нельзя изменить часть массива S!C1 (C1:C3)` in a ru-RU workbook) and one of these codes: `off_sheet` (cells or a table would be pushed off the sheet), `sheet_name` (a name Excel would not accept, including `History`, or one another sheet already has), `last_sheet` (deleting or hiding the last visible sheet), `very_hidden` (deleting or moving a very hidden sheet), `merged` (a move that cuts a merged area, or a copy that lands on part of one), `filtered` (a copy that left out rows hidden by a filter, pasted over hidden rows of a filtered sheet), `array` (part of a legacy array formula's range would change, or a merge would cover one), `table` (inserting rows or columns into two tables at once, deleting from two tables unless both are deleted entirely, deleting a header row alone, moving or copying an array formula of several cells onto a table, or merging cells of a table), `table_name`, `outline_deep` (Group on rows or columns at the deepest level), `not_grouped` (Ungroup on rows or columns in no group, or a group button where there is none) or `detail` (Show Detail on an expanded group, or Hide Detail on a collapsed one). A name whose reference is lost returns `#REF!` where it is used, as in Excel; the parser reads the lost references Excel writes (`'S'!#REF!`, `#REF!A1`) as `#REF!`.

A workbook keeps styles as xlsx does: `styles`, a table of cell styles (font, solid fill, borders, alignment); `style` in a cell, an index into that table; and per sheet, `columns` and `rows` (keyed by a single row or column, `C` or `5`, or a span, `B:D` or `5:9`, with a width or height, a style and a number format) and `merges` (`B2:D3`). A cell without its own style or number format displays its row's, otherwise its column's. Calculation ignores them, except that a merged area stops an array from spilling over it, as in Excel. The operations change them as Excel does: inserted rows take the format of the row above (its height, style and number format, and its cells' styles and number formats) and inserted columns that of the column to the left, merged areas grow, shrink and move like ranges, moved and copied cells keep their formatting, and entire rows and columns take their sizes and styles along. A cell's note and hyperlink go with the cell, as its style does; frozen panes keep the rows and columns before their split; data validation rules move as references do, and rows or columns inserted immediately after a rule get it; and inserted rows and columns get the outline level of the one before them. The rules, and how they were checked against Excel on 317 workbooks and the validation sets, merging on 61 more, and the fill handle on 122 in Excel for Windows, are in [docs/structure.md](structure.md).

A session applies the same operations to its workbook with `session.editStructure(op)` and recalculates, keeping the workbook's styles, so a host can hold the workbook in the session alone and save `getBook()` as a version:

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{ name: 'Plan', cells: { A1: { value: 100 }, A2: { value: 250 }, A3: { formula: 'SUM(A1:A2)', style: 0 } } }],
  styles: [{ font: { bold: true } }],
})
const r = session.editStructure({ op: 'deleteRows', sheet: 'Plan', row: 1 })
console.log(r.mapCell({ sheet: 'Plan', cell: 'A3' }), r.moved)
// → { sheet: 'Plan', cell: 'A2' } [ { kind: 'lines', sheet: 'Plan', axis: 'rows', at: 1, count: -1 } ]
console.log(r.changes.map((c) => [c.cell, c.value, c.formula])) // → [ [ 'A2', 250, 'SUM(A1:A1)' ] ]
console.log(session.getBook().sheets[0].cells) // → { A1: { value: 250 }, A2: { formula: 'SUM(A1:A1)', style: 0 } }
```

It returns `mapCell`, as `editStructure` does, but not the workbook, which `getBook()` provides; `moved`, how the addresses moved (`{ kind: 'lines', sheet, axis, at, count }` for inserted or deleted rows and columns, a moved range, and a renamed, deleted, inserted or moved sheet); and `changes`, the cells whose value or formula text differs from that of the cell previously at their position (`SUM(A1:A2)` above became `SUM(A1:A1)`), so an interface that shifts its viewport by `moved` (or maps each cell with `mapCell`) and applies `changes` is up to date. It rebuilds the state, at the cost of a cold calculation and a comparison of the cells with those previously at their positions (170 ms for the 3,400-row synthetic workbook), except for hiding or unhiding a sheet (`hideSheet`, `unhideSheet`), which changes no value; the outline commands, which recalculate only the `SUBTOTAL` and `AGGREGATE` formulas over the rows they hide and show; and a merge that moves no cell, which writes the cells of its range as an edit does. Each operation is one undo step ([Undo and redo](sessions.md#undo-and-redo)).

## API

```ts
editStructure(book: BookInput, op: StructureOp): StructureResult
session.editStructure(op: StructureOp): SessionStructureResult   // `mapCell` and what an edit answers, without the book

type StructureOp =
  | { op: 'insertRows' | 'deleteRows'; sheet: string; row: number; count?: number }          // 1-based row
  | { op: 'insertColumns' | 'deleteColumns'; sheet: string; column: string; count?: number } // letters
  | { op: 'moveRange'; from: RangeRef; to: CellRef }                       // cut and paste, `to` is the top-left cell
  | { op: 'copyRange'; from: RangeRef; to: RangeRef; withMeta?: boolean }  // copy and paste, fill
  | { op: 'fillRange'; sheet: string; range: string; to: string; type?: FillType } // the fill handle, `to` holds `range`
  | { op: 'renameSheet'; sheet: string; name: string }
  | { op: 'deleteSheet'; sheet: string }
  | { op: 'insertSheet'; name: string; before?: string }  // an empty sheet, at the end without `before`
  | { op: 'moveSheet'; sheet: string; before?: string }   // to the end without `before`
  | { op: 'hideSheet'; sheet: string; very?: boolean }   // very hidden with `very`
  | { op: 'unhideSheet'; sheet: string }
  | { op: 'mergeCells'; sheet: string; range: string; across?: boolean; center?: boolean }
  | { op: 'unmergeCells'; sheet: string; range: string }
  | { op: 'addTable'; sheet: string; table: TableInput }    // Format as Table
  | { op: 'renameTable'; table: string; name: string }
  | { op: 'deleteTable' | 'convertTable'; table: string }   // Delete Table, Convert to Range
  | { op: 'group' | 'ungroup'; sheet: string; rows: string }        // 5:7, or one row 5
  | { op: 'group' | 'ungroup'; sheet: string; columns: string }     // B:D, or one column C
  | { op: 'showDetail' | 'hideDetail'; sheet: string; row: number } // a group's + and − buttons
  | { op: 'showDetail' | 'hideDetail'; sheet: string; column: string }
  | { op: 'showLevels'; sheet: string; rows?: number; columns?: number } // the level buttons

type RangeRef = { sheet: string; range: string }       // B2, B2:D9, B:D, 2:9
type FillType = 'default' | 'series' | 'copy'
type StructureResult =
  | { ok: true; book: Book; mapCell(ref: CellRef): CellRef | null }
  | { ok: false; code: StructureRefusal; message: string }
type StructureRefusal =
  | 'off_sheet' | 'sheet_name' | 'last_sheet' | 'very_hidden' | 'merged' | 'filtered' | 'array'
  | 'table' | 'table_name' | 'outline_deep' | 'not_grouped' | 'detail'
```

The types are in `xlcalc/contract`.

An operation is plain data, so an interface can store it, send it to a server and replay it. Rows are numbers and columns are letters, as in A1 addresses. Sheet names are matched case-insensitively.

The input workbook is not changed. Sheets, cells and names the operation does not touch are shared with it.

`mapCell` tells where a cell of the old workbook is now: a new address and sheet name, or `null` if the cell was deleted or overwritten, or if its sheet is not one of the old workbook's (a viewport can show a sheet that no longer exists). A host uses it for anything it keeps by address, such as comments, the selection or links to source rows.

Input of the wrong shape (an unknown sheet, row 0, a malformed range, unknown fields) throws `InputError`. An operation that Excel would reject returns `ok: false` with a message for the user:

| code | when |
|---|---|
| `off_sheet` | an insert would push a cell, a merged area or a table past row 1,048,576 or column XFD, or a move or copy does not fit on the sheet |
| `sheet_name` | the new name breaks Excel's rules or another sheet already has it |
| `last_sheet` | deleting or hiding the last visible sheet (including the only sheet) |
| `very_hidden` | deleting or moving a very hidden sheet |
| `merged` | a move whose source cuts a merged area, or a copy that lands on part of one |
| `filtered` | a copy that left out rows hidden by a filter, pasted over hidden rows of a filtered sheet |
| `array` | rows or columns inserted inside the range of a legacy array formula, part of one deleted or moved, a paste over part of one, or a merge over one |
| `table` | rows or columns inserted into two tables at once, or deleted from two tables unless both are deleted entirely, a table's header row deleted without the rest, a move that cuts a table, a move or a copy that lands an array formula of several cells on a table (as `setCells` refuses one), or a merge over cells of a table ([Merging cells](#merging-cells)); for operations on tables, see [docs/tables.md](tables.md) |
| `table_name` | `renameTable` to a name Excel would not accept, or one a table or a defined name already has ([docs/tables.md](tables.md#tables-themselves)); a copy of an entire table whose numbered name Excel would not accept (`Tab3`, [docs/tables.md](tables.md#copies-and-moves)) |
| `outline_deep` | Group on rows or columns that are all at the deepest outline level, 7 ([docs/outline.md](outline.md)) |
| `not_grouped` | Ungroup on rows or columns none of which is grouped; Show or Hide Detail at a row or column in no group, or in a group without a summary row or column |
| `detail` | Show Detail on a group that is already expanded, or Hide Detail on one that is already collapsed |

Excel rejects an insert only when a non-empty cell would be pushed off the sheet, and drops the formatting pushed off. Here a cell counts also when it holds only `meta`, a note or a link, so none of these is dropped silently, and each cell of a legacy array formula's area counts, as it holds a value in Excel. Formatting alone is dropped as in Excel: the part of the sheet's `looks` pushed off, and the formatting of entire rows and columns past the edge. A look over an entire column stays over the entire column when rows are inserted, and one over an entire row when columns are, so the wide shape of [docs/grid.md](grid.md#formatting-of-many-cells) (`A2:XFD4`) takes inserted columns.

## Formula text

Only the references that change are rewritten, in place. The rest of the formula keeps its spelling, including spaces and lowercase, and a formula without a changed reference keeps its text byte for byte. A rewritten reference is written as Excel writes it: a range covering every row as columns (`A1:A1048576` becomes `A:A`), a range covering every column as rows, and a new or changed sheet name in quotes.

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

The parser reads all these forms, which also occur in xlsx files, as `#REF!` errors and prints them back unchanged. A name whose reference is lost stays in the workbook, as in Excel, and returns `#REF!` where it is used.

## Inserting rows and columns

Cells from the inserted row onward move down, with their `numberFormat`, `style` and `meta`. The inserted rows copy the row above, as Excel does by default: its height, style and number format, and the style and number format of each of its cells, without values. Rows inserted above row 1 are unformatted. Columns work the same way, with the column to the left.

Every reference to the sheet moves, absolute references included, from any sheet and from global and local names.

- A range grows when rows are inserted between its first and last row, moves when they are inserted at or above its first row, and stays when they are inserted below it.
- An entire column (`B:B`) does not change. Entire rows (`2:3`) behave like ranges.
- At the edge of the sheet, a single cell pushed past the last row becomes `#REF!`. A range that ends on the last row keeps ending there (`A5:A1048576` becomes `A7:A1048576`) and becomes `#REF!` when its first row is pushed past it. A range covering every row does not change.

## Deleting rows and columns

The cells in the deleted rows are removed and the cells below move up. A reference inside the deleted rows becomes `#REF!`. A range loses its deleted rows: `A3:A8` with rows 5 and 6 deleted becomes `A3:A6`, and `A5:A6` with row 6 deleted becomes `A5:A5`, still a range. A range whose rows are all deleted becomes `#REF!`. A range that ends on the last row keeps ending there unless all its rows are deleted.

## Moving a range (cut and paste)

The cells move with their `numberFormat`, `style` and `meta`, and replace the cells they land on; an empty source cell clears its destination. A reference that lies entirely inside the moved cells follows them, to another sheet too, absolute or not. A reference entirely inside the overwritten cells becomes `#REF!`.

A range with one end among the moved or overwritten cells changes along that end's axis, when the cells span the range's other axis and move along this one:

- An end among the moved cells that moves away from the other end follows them: cutting `A10` to `A15` turns `SUM(A1:A10)` into `SUM(A1:A15)`.
- An end that moves toward the other end, to cells outside the range, does not change it.
- A range that contained all the moved cells drops them and takes in where they land: cutting `A4:A5` to `A7` turns `A4:A9` into `A6:A9`.
- When the moved cells straddle an end, the end follows them, unless that leaves fewer than two rows or, for a move that does not overlap itself, a range inside the overwritten cells. In that case the range drops the moved cells. Either way it takes in the overwritten cells beyond it.
- An end among the overwritten cells drops them, unless the range contains them all: cutting `A4:A5` to `A7` turns `A1:A7` into `A1:A6` and leaves `A1:A8` unchanged.

Cells moved across a range's axis, such as sideways for a column range, do not change it, and a range covering an entire column or row stays entire. When the cells move to another sheet, a range with one end among them shrinks past them.

A moved formula's references to cells that did not move keep pointing at them; a reference without a sheet gets the name of the sheet it pointed to. A moved formula that uses a name local to its old sheet gets that name defined on the new sheet too, if the new sheet sees no name with that name, its own or a global one.

## Copying a range (copy and paste, fill)

The copied cells replace the destination's, with `numberFormat` and `style`. They do not take `meta` unless the operation has `withMeta: true`: a copy is a new cell, and what a host keeps in `meta`, such as the IDs of source rows, belongs to the original. Empty source cells clear the cells they land on.

Relative parts of references move with the copy and absolute ones (`$`) stay. An entire column keeps its rows and an entire row its columns. A reference pushed past the edge of the sheet becomes `#REF!`, or `'S'!#REF!` when it names its sheet; a range is not clamped.

A destination that is a whole multiple of the source's size is tiled with copies, so fill down is a copy of the first row over the rows below. Any other destination gets one copy at its top-left cell, as Excel's `Range.Copy` does. Other formulas and names do not change. A formula copied to another sheet that uses a name local to the source sheet gets the name defined on the destination sheet, with its references to the source sheet pointed at the destination, under the same condition as for a move.

`copyFormula(text, from, to, book?)` returns the text of one formula copied from cell `from` to cell `to` (each `{ sheet, cell }`) by the same rules, for an interface that previews a paste before making it. Given the workbook (a `BookInput`, or just its sheet names and tables), a formula copied out of a table names the table in its references to it, as a copy does (`[@Price]` becomes `Sales[[#This Row],[Price]]`); without it, the formula's references keep the sheet names they are written with.

```js
import { copyFormula } from 'xlcalc'

console.log(copyFormula('A1+$B$2+SUM(C1:C3)', { sheet: 'S', cell: 'D1' }, { sheet: 'S', cell: 'E3' })) // → B3+$B$2+SUM(D3:D5)
console.log(copyFormula('A1*2', { sheet: 'S', cell: 'B2' }, { sheet: 'S', cell: 'A1' })) // → #REF!*2
```

What a copy does beyond one formula's text is specific to `copyRange`: it leaves out the rows hidden by a filter, copies part of a legacy array formula's range as separate arrays, points the formulas of an entirely copied table at the copy, and turns a formula pasted on a table's header into the header's text.

## Renaming and deleting sheets

Renaming a sheet renames every reference to it, including `'S'!#REF!`, and the scope of its local names. The new name must follow Excel's rules (1 to 31 characters, none of `: \ / ? * [ ]`, no apostrophe at either end, not `History` in any case) and differ from the other sheets' names, compared case-insensitively. Renaming to a different case is allowed. A workbook whose sheet names break these rules causes an `InputError`.

Deleting a sheet removes it with its cells. References to it become `#REF!` followed by the reference (`#REF!A5`, `SUM(#REF!$A$3:$A$8)`), in formulas, in global names and in names local to other sheets. Names local to the deleted sheet are removed.

## Inserting and moving sheets

`insertSheet` adds an empty sheet before another one or at the end. `moveSheet` moves a sheet. Neither changes a formula, except a reference across sheets whose end moves. Such a reference covers the sheets between its ends in the workbook's order, so a sheet inserted or moved between them is included and one moved out is not. An end that moves remains the end as long as it does not pass the other end: with Jan, Feb, Mar and Apr, moving Jan before Apr leaves `'Jan:Apr'` unchanged, now covering Jan and Apr, while `'Jan:Mar'` becomes `'Feb:Mar'`. [References](references.md) has the rules for its ends.

## Hidden sheets

A sheet is hidden by `hidden: true`, or very hidden by `hidden: 'very'`, as Excel's Hide and its object model's `xlSheetVeryHidden` hide it (an xlsx file's `state`). `hideSheet` hides a sheet (`very: true` for very hidden) and `unhideSheet` unhides it. Formulas read hidden sheets like any other: `SUM('Jan:Mar'!B2)` with Feb hidden includes Feb, and `SUBTOTAL` skips only hidden rows, never a hidden sheet.

A workbook keeps at least one visible sheet, as Excel does: hiding or deleting the last visible sheet is rejected with `last_sheet`, and a workbook whose sheets are all hidden causes an `InputError`. A hidden sheet is renamed, deleted and moved like any other and keeps its state; a very hidden one can only be renamed, and deleting or moving it is rejected with `very_hidden`, as Excel's object model rejects both.

Excel positions a moved sheet relative to the visible sheets. Moved before a hidden sheet, it goes before the next visible sheet; if there is none after it, or when moved to the end past hidden sheets, it goes after the last visible sheet. With Jan, Feb (hidden), Mar and Sum, moving Mar before Feb leaves it where it is, and with Jan, then Feb and Mar hidden at the end, moving Jan to the end puts it between Feb and Mar. An inserted sheet goes where requested, even before a hidden sheet.

```js
import { editStructure } from 'xlcalc'

const book = { sheets: [{ name: 'Data', cells: {} }, { name: 'Keys', cells: {}, hidden: 'very' }, { name: 'Report', cells: {} }] }
const r = editStructure(book, { op: 'hideSheet', sheet: 'Data' })
console.log(r.ok && r.book.sheets.map((s) => s.hidden)) // → [ true, 'very', undefined ]
console.log(r.ok && editStructure(r.book, { op: 'hideSheet', sheet: 'Report' }).code) // → last_sheet
console.log(editStructure(book, { op: 'deleteSheet', sheet: 'Keys' }).code) // → very_hidden
```

## Tables

Tables move with their cells and grow and shrink with the rows and columns inserted and deleted inside them; structured references follow them by name. Inserted rows get the formulas of calculated columns and inserted columns get a name (`Column1`); references to a deleted column or table become `#REF!`. Inserting rows or columns into two tables at once, deleting from two tables unless both are deleted entirely, and deleting a header row alone are rejected with `table`. `addTable`, `renameTable`, `deleteTable` and `convertTable` change tables themselves. [Tables](tables.md) has the rules.

## Names, references across sheets and legacy arrays

Names' formulas change like cells' formulas: each part of a union separately, and a reference without `$` not at all under inserted and deleted rows and columns ([Names](names.md)). A reference across sheets stays unchanged under row and column operations and changes with its end sheets ([References](references.md)). The range of a legacy array formula moves with its cell; an operation that would change part of it is rejected with `array` ([Legacy array formulas and implicit intersection](legacy.md)).

## Styles and layout

A workbook keeps styles the way an xlsx file does (`BookInput.styles`, `CellInput.style`, `SheetInput.columns`, `rows`, `merges`): a table of cell styles (font, solid fill, borders, alignment), a style index in each cell, and per sheet the columns and rows by single row or column or by span (`C`, `B:D`; `5`, `5:9`) with a width or height, a style and a number format, plus the merged areas. A cell's number format is its `numberFormat`. A cell displays its own style, otherwise its row's, otherwise its column's, and the same for the number format. Widths are in characters, as xlsx stores them, and heights in points. Calculation does not read styles, and they are not part of the values hash; `inputHash` includes them when a workbook has any.

Excel resolves an empty cell's style and number format together: from its row if the row has either, otherwise from its column. The two can differ only where a row with a style but no number format (or the other way round) crosses a column with the other; an xlsx reader writes such a case out explicitly on the row.

**Inserting and deleting.** Columns and rows keep their settings as they move. A span that contains the row or column before the inserted ones takes them in. A span that would extend past the end of the sheet is truncated there, and deleted rows and columns are removed from their spans. The cells of inserted rows and columns get the style and number format of the cells of the one before them, including the sheet's `looks` (`SheetInput.looks`): these move with their rows and columns, and the new ones get the formatting of the one before them, without an entry per cell.

A merged area changes like a range: it grows with rows or columns inserted inside it, moves with those inserted at or above its first row or column, stays when they are inserted immediately after it (the new row or column is not merged), loses deleted rows and columns, and disappears when only one cell or none is left.

At the end of the sheet, a span of columns that reaches XFD keeps reaching it when columns are deleted, as xlsx stores such a span (`<col max=16384>`). If the deleted columns include XFD, the columns that come in at the end get the span's style and number format but not its width (observed once). Rows are stored individually and do not extend. A merged area that ends on the last row or column keeps ending there when a deletion reaches the end of the sheet, unless all of it is deleted.

**Moving and copying.** A cell takes its own style and number format along. Every pasted cell displays the style and number format its source displayed: where the destination's rows and columns would display something else (a cell moved out of a bold column into a blue row), the cell gets the source's style written out explicitly, with `General` for no number format and an empty style, added to the workbook's table if missing. Entire rows or columns take their settings along; a move leaves the source rows or columns unformatted, and a copy repeats them on every tile.

Merged areas inside the source go along. A copy also takes the part of a merged area that its source cuts, when that part has two cells or more. Merged areas that pasted cells land on are removed: all of them for a move, and those inside the destination for a copy.

**Rejections.** A move whose source cuts a merged area, and a copy that lands on part of one, are rejected with `merged`, as Excel rejects them. A move that lands on part of a merged area removes it, and a copy whose source cuts one copies the part.

A renamed sheet keeps its layout, and a deleted sheet takes it along.

## Merging cells

`mergeCells` is Excel's Merge Cells: the range becomes one merged area (`SheetInput.merges`). With `across: true` each row of the range is merged separately, as Merge Across does; with `center: true` the area is also centered, as Merge & Center does. `unmergeCells` is Unmerge Cells.

`mergedArea(rect, merges)` from `xlcalc/a1` returns the rectangle a merge covers, extended by each merged area it overlaps, for a grid that warns before a merge discards values.

```js
import { editStructure } from 'xlcalc'

let book = { sheets: [{ name: 'Plan', cells: { B2: { value: 'Q1' }, C2: { value: 'gone' } } }] }
const merge = (op) => {
  const r = editStructure(book, { sheet: 'Plan', ...op })
  book = r.book
  return book.sheets[0].merges
}
console.log(merge({ op: 'mergeCells', range: 'B2:D2', center: true }), book.sheets[0].cells)
// → [ 'B2:D2' ] { B2: { value: 'Q1', style: 0 } }
console.log(merge({ op: 'mergeCells', range: 'B4:D6', across: true }))
// → [ 'B2:D2', 'B4:D4', 'B5:D5', 'B6:D6' ]
// unmerging at C2 splits B2:D2, which holds C2
console.log(merge({ op: 'unmergeCells', range: 'C2' }))
// → [ 'B4:D4', 'B5:D5', 'B6:D6' ]
```

### What the cells keep

The area keeps its top-left cell. Its other cells lose their values and formulas, as Excel merges them once its warning ("Merging cells only keeps the upper-left value") is accepted. Formulas elsewhere that read a cleared cell keep reading it, and now read an empty cell.

When the top-left cell is empty, the first cell of the area that contains a value or a formula, row by row, is moved into it, as cut and paste moves a cell: formulas and names that read it follow it, its own formula keeps its text, and its style and number format come with it. In `A1:C2` with A1 empty, B1 = 5 and A2 = 7, B1 moves to A1, `=B1*2` elsewhere becomes `=A1*2`, and A2 is cleared. As with cut and paste over a cell, a reference to the empty A1 itself becomes `#REF!` (`=A1` becomes `=#REF!`), in a name and on another sheet too. A cell of a spill has no content of its own and is not moved. `mapCell` tells where the moved cell went.

### Formatting

Every cell of the area gets the top-left cell's style and number format (its own, otherwise its row's, otherwise its column's), so that after Unmerge Cells the cells keep the formatting the area had. With `center`, that style gets centered horizontal alignment, added to the workbook's style table if no such style exists.

### Overlapping merged areas

A range that overlaps a merged area takes it in entirely, and extends again for each area it then overlaps: merging `D2:E2` next to the merged area `E2:F3` merges `D2:F3`, and merging cells inside a merged area merges it again. With `across`, the overlapped merged areas are split into the rows of the range. Unmerge Cells splits every merged area that overlaps its range, and the cells keep their values and formatting; a range without one changes nothing.

Entire rows and columns can be merged too: `'20:21'` becomes one area `A20:XFD21`, as in Excel.

### Rejections

A range that overlaps a table is rejected with `table`: Excel does not merge the cells of a table. A range that overlaps a legacy array formula is rejected with `array`.

## The fill handle

`fillRange` is Excel's fill handle (`Range.AutoFill`): the cells of `range` are continued over `to`, which contains `range` and extends beyond it down, up, right or left. A host that lets a user drag the handle sends the source and the full extent of the drag.

```js
import { editStructure } from 'xlcalc'

const book = { sheets: [{ name: 'Plan', cells: {
  A1: { value: 1 }, A2: { value: 2 },
  B1: { value: 'Item1' },
  C1: { value: 45688, numberFormat: 'yyyy-mm-dd' }, C2: { value: 45716, numberFormat: 'yyyy-mm-dd' }, // Jan 31, Feb 28
  D1: { value: 'Mon' },
  E1: { formula: 'A1*10' },
} }] }
const r = editStructure(book, { op: 'fillRange', sheet: 'Plan', range: 'A1:E2', to: 'A1:E5' })
const c = r.book.sheets[0].cells
console.log(c.A3.value, c.A5.value, c.B3.value, c.B5.value, c.C3.value, c.C4.value, c.D3.value, c.E3.formula)
// → 3 5 Item2 Item3 45747 45777 Tue A3*10
```

### Copies first

The added cells take the source's cells in turn, as copy and paste lays them out: formulas with their relative references adjusted, styles, number formats, notes and hyperlinks, and conditional formatting and data validation rules. When `to` is not a whole number of source lengths, the part of the source nearest to it comes at the far end: filling `A1:A3` to `A1:A8` lays `A1:A3` over `A4:A6` and `A1:A2` over `A7:A8`. Rejections are those of a copy (`merged`, `array`, `table`, `filtered`).

### What continues as a series

Each column of a fill down or up, and each row of a fill right or left, is continued independently. Its source is split into runs: adjacent cells of one kind, skipping the empty cells between them, with text without a number, booleans and formulas ending a run. A run's cells appear again in every copy of the source, so a run continues from copy to copy: `1, 2, x` continues as `3, 4, x, 5, 6, x`, and `1, (empty), 3` as `5, (empty), 7`. A fill up or left goes backward: `1, 2` filled up gives `0, -1` above.

- Numbers: two or more continue by their step, or by their least-squares trend when the steps differ (`1, 3, 4` gives `5.67, 7.17`). A single number is copied when the whole source contains only plain numbers and empty cells, and increases by 1 when the source also contains anything else (text, a boolean, a formula, a date).
- Dates (numbers displayed in a format with a day, a month or a year): a single date increases by days; dates on the same day of the month, or each on the last day of its month, a whole number of months apart, continue by months (`Jan 31, Feb 28` gives `Mar 31, Apr 30`; `Jan 15, Apr 15` gives `Jul 15`); other dates with a constant step in days continue by that step, and dates without a constant step are copied. A date given as `{ date }` stays one, except past the date system (before its day 0 or after 9999-12-31), where the cell gets the serial number, as a number in a date format would.
- Times of day: a single time increases by hours; times with a constant step continue by it, and others are copied.
- Text with a number: a number at the start followed by a space continues (`10 kg`, `11 kg`; `1 item 2`, `2 item 2`), otherwise a number at the end (`Item1`, `Item2`; `Item09`, `Item10`), with the step of the run, and a run without a constant step is copied. A number below zero is written as its absolute value (`Item1` filled up gives `Item0`, `Item1`, `Item2`). Quarters wrap around the year (`Q4`, `Q1`; `Qtr`, `Quarter 1`), and en-US ordinals take their suffixes (`1st`, `2nd`, `3rd`, `11th`, always lowercase). Text with the number inside (`a1b`) or immediately before letters (`1abc`) is copied.
- Names of days and months, short or long (`Mon`, `Monday`, `Jan`, `January`), cycle through their list in the case they are written in (`MONDAY`, `TUESDAY`): en-US's names, and in a workbook in another locale, that locale's names too.
- Everything else is copied: text, booleans, errors, formulas.

`type` changes this as Excel's `XlAutoFillType` does: `'series'` makes a single number increase by 1 as well (Excel's drag with Ctrl over a number), and `'copy'` copies every cell.

A series that continues over a table's header cells renames their columns to the text it writes there, as a paste over them does (`Col1` filled right over two headers gives the columns `Col2` and `Col3`; numbers become their text). How Excel names such columns was not recorded.

`mapCell` of a fill is the same as for a copy: the cells it overwrites are cleared, and the others stay.

## Notes, hyperlinks and frozen panes

A cell's `note` and `link` go with the cell, as its style does:

- Inserted rows and columns get none, even where the row or column before them has some.
- A deleted cell takes them with it.
- A move takes them along. The cells it lands on lose theirs, even where the moved cell has none.
- A copy copies them to every cell it pastes, and a pasted cell without them clears the destination's.
- A hyperlink's `location` stays as written when its sheet is renamed or deleted, or when the rows it points to move. Excel does not update it either.

Frozen panes (`frozen`) keep the rows and columns before their split. Rows inserted before the split push it down: with two rows frozen, a row inserted at row 1 or 2 makes three, and one inserted at row 3 leaves two. Rows deleted before the split pull it up, but at least one row stays frozen: deleting rows 1 and 2 of two frozen ones leaves row 1 frozen. Columns work the same way. A move or a copy does not change them, and a renamed sheet keeps them.

The parts sets `struct-held` and `frozen-held` verified these rules in Excel for Mac and for Windows, with the engine's results committed before Excel ran them ([docs/xlsx.md](xlsx.md#notes-hyperlinks-and-frozen-panes)).

## Conditional formats

A rule's ranges and formulas change as Excel changes them ([docs/conditional-formats.md](conditional-formats.md#structure-operations)):

- Inserted rows or columns that follow a row or column of a rule join it, as they get the format of the one before them; a rule after them moves.
- Deleted rows and columns are removed from its ranges, and a rule with nothing left is removed. A color scale, a data bar, an icon set, or a top, average or duplicate rule that loses rows or columns moves after the other rules.
- A move takes its cells' rules along, and the destination cells keep theirs; when moved to another sheet, the moved part becomes a rule there, first in order. A rule whose formula the move rewrote moves after the others.
- A copy pastes its cells' rules as new rules, first in order, and the pasted cells lose theirs; a fill down a rule's own columns extends the rule.
- Formulas change as the formula of the rule's first cell (the top-left of its cells) would; a reference to a deleted sheet becomes `#REF!`.

The conditional formatting sets `struct-held`, `held` and `held-2` verified these rules in Excel for Mac and for Windows, with the engine's results committed before Excel ran them.

## Data validation and outlines

A data validation rule (`SheetInput.validations`, [docs/validation.md](validation.md)) is a range of cells plus formulas written for its first cell. Its ranges change as references do, with one difference at their end:

- Rows or columns inserted after a range's first row or column extend it, those inserted at or before its first one move it, and those inserted immediately after its last one join it: the inserted cells get the rule of the cells before them, as they get their format. With a rule over `B2:B6`, a row inserted at row 4 or at row 7 makes it `B2:B7`, and one inserted at row 2 makes it `B3:B7`. Columns work the same way, with the column to the left.
- Deleted rows and columns are removed from the ranges; a range that loses all of them is removed, and so is a rule that loses all its ranges. A range that ends on the sheet's last row or column keeps ending there, even a single cell, unlike a reference to a cell: a rule over `B1048576` covers `B1048575:B1048576` after a row above is deleted, as a conditional format's does.
- A move takes the rules of the moved cells along. The destination cells keep their own rules where the moved cells have none, and get the moved ones where they have some. A move from `A1`, which has no rule, onto a cell with a whole-number rule leaves that cell's rule.
- A copy pastes the rules of the copied cells on every tile, and a pasted cell without a rule clears the destination's rule, as Excel's Paste All does.
- References in a rule's formulas change as a cell's formulas do: a list `'L'!$A$1:$A$4` follows rows inserted on L, becomes `'Lists'!$A$1:$A$4` when L is renamed, and `#REF!$A$1:$A$4` when it is deleted. A rule whose first cell moves or is removed has its formulas rewritten for its new first cell: deleting row 2 of a rule over `D2:D6` with `D2<A2` leaves `D2:D5` with `D2<A2`, the formula of the old D3. A rule moved to another sheet has its formulas adjusted as a copy adjusts them: `B2>A2` moved from B2 on S to D2 on T becomes `D2>C2` there, while a move on its own sheet keeps the references to the cells that did not move (`D2<A2` moved to D20 becomes `D20<A2`).

An outline consists of the outline levels of rows and columns (`outlineLevel`, 1 to 7) with their collapsed markers (`collapsed`), and the sheet's summary settings (`SheetInput.outline`, below and to the right by default). The rows and columns of a collapsed group are hidden, and `SUBTOTAL` 101 to 111 skips them.

- Inserted rows and columns get the outline level of the one before them, as they get its height, and are visible. Rows or columns inserted into a collapsed group after one of its hidden rows or columns expand it: its summary row or column, after the group (before it with summaries above or to the left), loses its collapsed marker. Rows inserted at the summary row, or above the group, leave it collapsed.
- Deleted rows and columns take their levels and markers along; moved and copied entire rows and columns carry them, as they carry their other settings.

Excel's outline commands (Group, Ungroup, Show Detail, Hide Detail and the level buttons) are also structure operations, which move no cell: [docs/outline.md](outline.md) covers them.

The validation sets `struct-held` and `struct-held-2` verified these rules in Excel for Mac and for Windows, with the engine's results committed before Excel ran them ([Development](development.md#data-validation)).

## Hidden rows and filtered sheets

A row or a column is hidden by `hidden: true` in `rows` or `columns`, or by a height or width of 0, as Excel treats it. A sheet is marked `filtered: true` when one of its filters has criteria, whether its AutoFilter or a table's. Excel then treats every hidden row of the sheet as one hidden by the filter, whatever hid it. In an xlsx file the criteria decide, not the `filterMode` attribute: Excel treats a sheet as filtered when its AutoFilter has criteria and the attribute is missing, does not when the attribute is present next to an AutoFilter without criteria, and writes no attribute for a table's filter ([docs/xlsx.md](xlsx.md)).

On any sheet:

- Inserted rows and columns are visible, even inside a hidden block. They copy every other setting of the row or column before them, including its size: a row inserted under a row of height 0 has height 0 and so is hidden, while one inserted under a row hidden at height 30 is visible at 30 (the `zero-held` set, in both versions of Excel).
- Deleted rows and columns take their `hidden` setting along, and moved or copied entire rows and columns carry it.
- Cells move and copy regardless of `hidden` on a sheet that is not filtered. A copy of `A2:C16` with rows 4 and 7 hidden manually copies all 15 rows.

On a filtered sheet, Excel acts on the visible rows:

- Deleting rows deletes the visible rows among them. Deleting rows 6 to 9 when rows 7 and 8 are hidden deletes rows 6 and 9.
- Inserting rows inserts as many rows as are visible among them, before the first. Inserting two rows at row 8, when row 8 is hidden and row 9 is visible, inserts one row.
- A copy whose source covers hidden rows takes only the visible rows, one after another, and pastes their values instead of their formulas. A copy of `A2:C16` with rows 4, 7, 8 and 11 hidden pastes 11 rows of values.
- A paste of several cells over hidden rows goes into the visible rows. Each run of visible rows is treated as a separate destination for the copy. A copy of `D2:E2` onto `D3:E16` fills the visible rows and leaves the hidden ones unchanged.
- A copy that left rows out cannot be pasted over hidden rows. The operation is rejected with `filtered`, as Excel rejects it.
- A selection consisting only of hidden rows is taken in full.

The values a copy pastes come from the calculated workbook. `session.editStructure` uses the session's values. `editStructure` calculates the workbook first, at the current time, which matters only for volatile formulas.

`copyRange` over a filtered sheet fills its visible rows, as Excel's Fill Down does. `fillRange`, the fill handle, is not copy and paste: it fills the rows hidden by a filter just as it fills rows hidden manually, for copies and series alike, and takes a source that contains hidden rows in full. `x`, `y` filled from `A1:A2` to `A6` over a hidden row 3 gives `x`, `y`, `x`, `y` in rows 3 to 6, as Excel's fill handle does.

## Sessions

`session.editStructure(op)` applies the operation to the session's workbook and recalculates. It returns `mapCell`, as `editStructure` does, without the workbook, which `getBook()` provides, and what every session edit returns ([Editing a grid](edits.md#what-an-edit-returns)): `moved`, how the addresses moved, and `changes`, every cell whose value or formula text differs from that of its previous cell. An interface that shifts its viewport by `moved` (or maps each cell with `mapCell`) and applies `changes` has the new values and formulas. An insert into a block of data lists the formulas over it, which now read one more row, and no values. A rejected operation changes nothing.

The session rebuilds its state from the new workbook. On the synthetic workbook (3,400 rows, 3,907 formulas), inserting a row in the middle of the data takes 170 ms, about as long as a cold calculation of the same workbook, well within the 1.5 s budget. On the workbook of 34,000 rows (506,594 cells on the data sheet), inserting or deleting a row takes 1.3 to 1.5 s, its undo and redo 1.4 to 1.5 s each, copying a row down 2,000 rows 1.2 s, and moving a block of 100 by 20 cells 1.1 s; rebuilding the state accounts for half of it, and the operation itself for 0.25 s.

`mapCell`, `moved`, the way undo and redo revert and reapply the operation, and the comparison of each cell with its previous cell all follow one description of where the operation moved the cells: rows or columns inserted or deleted, a range moved, cells pasted over, or a sheet renamed, inserted, deleted or moved. Rows deleted on a filtered sheet are one `lines` move for each run of visible rows, from the bottom up, and their undo restores each run; an undo's `mapCell` is the reverse of the operation's, and a redo's is the operation's own.

`getBook()` returns the session's workbook with all changes so far, which is what a host saves as a version. Its objects belong to the session and change with later edits, so serialize it immediately.

Styles and layout change through the session, so that its workbook remains the host's workbook:

- a cell's style, note and hyperlink through `setCells` or `patchCells`; an edit that keeps a cell's value and formula recalculates nothing and lists the cell with its new style or number format;
- the style table through `setStyles(styles)`; every style in use must remain in it;
- columns, rows and merged areas through `setLayout(edits)`: `{ sheet, columns: 'B:D', set: { width: 20 } }` patches each row or column (`null` removes a field, and spans are split where an edit cuts them), and `{ sheet, merges: [...] }` replaces the sheet's merged areas;
- the filter mode through `setLayout`: `{ sheet, filtered: true }`;
- hidden sheets through `editStructure`: `hideSheet` and `unhideSheet` recalculate nothing and do not rebuild the state;
- frozen panes through `setLayout`: `{ sheet, frozen: { rows: 1 } }`, and `{ sheet, frozen: null }` unfreezes them;
- conditional formats through `setLayout`: `{ sheet, conditionalFormats: [...] }` replaces the sheet's rules;
- data validation rules through `setLayout`: `{ sheet, validations: [...] }` replaces the sheet's rules;
- outlines through `setLayout`: a group is the rows' or columns' `outlineLevel`, and collapsing it hides its rows or columns and marks its summary row or column `collapsed` (`{ sheet, rows: '3:5', set: { hidden: true } }`, `{ sheet, rows: '6', set: { collapsed: true } }`), while `{ sheet, outline: { summaryBelow: false } }` sets where the summaries are.

A merge that moves no cell changes the cells of its range, the merged areas and the style table, as a formatting patch does, without rebuilding the state: on the workbook of 34,000 rows, merging 3 by 3 cells of the data that 2,800 SUMIFS read takes 0.17 s, and its undo as long, nearly all of it spent recalculating the formulas over the cleared cells. A merge that moves a cell into an empty top-left cell, or that overlaps an array formula or more than 65,536 cells, rebuilds the state, as an insert does (1.5 s on that workbook).

Only two of these recalculate values. A merged area added or removed over an array's spill range recalculates that array. Rows hidden or unhidden, and a new filter mode, recalculate the `SUBTOTAL` and `AGGREGATE` formulas over the affected rows and what depends on them.

## Checked against Excel

Tested in Excel 16 for Mac, en-US, on 317 workbooks; merging cells (61 cases) and the fill handle (138 cases) in Excel for Windows, and hidden sheets in both. All match except the known differences below. Not tested: a move that does not fit, most moves to another sheet, fonts other than bold, borders and alignment; for merging, borders, the vertical alignment, and the notes and hyperlinks of the cells it clears; and the fill handle in Excel for Mac and with a ru-RU Excel's lists.

## Known differences and limits

- Entire columns copied onto part of a merged area: Excel pastes them and keeps the area, while the engine rejects the copy with `merged` (two observations).
- The fill handle was not tested in Excel for Mac, nor in a ru-RU Excel (its lists of names), nor with types other than default, series and copy. Also not tested: whether a number at the end of text keeps its leading zeros beyond one width (`Item001`, here `Item002`), and a time next to a lone number (here it makes the number continue, as a date does).
- One move rule rests on a single observation: a straddled end does not follow into the overwritten cells of a move that does not overlap itself.
- Charts are not part of a workbook. Comment threads are stored as notes in a workbook ([docs/xlsx.md](xlsx.md#threads-of-comments)).
- References to other workbooks are not read, so they never occur in a valid formula.
