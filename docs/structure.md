# Structure operations

Inserting and deleting rows and columns, moving and copying ranges, merging and unmerging cells, renaming, deleting, inserting, moving, hiding and showing sheets. References in formulas and names change as Excel changes them, and so do cell styles, column and row settings, merged areas, notes, hyperlinks, frozen panes, conditional formats, rules of data validation and outline levels. It is the layer a grid interface needs: a pure function over `BookInput`, and the same operation as a session method.

The rules below were measured in Excel 16 for Mac (en-US), those of merging cells in Excel for Windows.

## In short

`editStructure(book, op)` inserts or deletes rows and columns, moves (cut and paste) or copies (copy and paste, fill) a range, renames or deletes a sheet, and adjusts every reference in formulas and names on every sheet as Excel adjusts it: ranges grow and shrink, lost references become `#REF!`, absolute references move with inserted rows. It returns the new book and `mapCell`, which says where a cell of the old book is now (null for a cell deleted or overwritten, or of a sheet the old book does not have); the input book is not changed.

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

The operations are `insertRows` and `deleteRows` (`row`, a row number, and `count`), `insertColumns` and `deleteColumns` (`column`, letters, and `count`), `moveRange` (`from: { sheet, range }`, `to: { sheet, cell }`), `copyRange` (`from` and `to` as `{ sheet, range }`; a target that is a whole multiple of the source is tiled, as in fill down; copies have no `meta` unless `withMeta: true`), `renameSheet` (`name`, by Excel's rules for sheet names), `deleteSheet`, `insertSheet` (`name`, `before`), `moveSheet` (`sheet`, `before`; Excel places a sheet by the visible ones), `hideSheet` (`sheet`, `very` for a very hidden sheet) and `unhideSheet` (`sheet`), over tables `addTable` (`sheet`, `table`), `renameTable` (`table`, `name`), `deleteTable` and `convertTable` (`table`; [docs/tables.md](tables.md)), and Excel's outline commands `group` and `ungroup` (`sheet`, `rows` or `columns`), `hideDetail` and `showDetail` (`sheet`, `row` or `column`) and `showLevels` (`sheet`, `rows`, `columns`; [docs/outline.md](outline.md)), and Excel's Merge Cells and Unmerge Cells, `mergeCells` (`sheet`, `range`, `across` for Merge Across, `center` for Merge & Center) and `unmergeCells` (`sheet`, `range`): a merge keeps the top-left cell, or moves the area's first value into it when it is empty, and clears the rest ([docs/structure.md](structure.md#merging-cells)), and Excel's fill handle, `fillRange` (`sheet`, `range`, `to` that holds `range` and reaches past it, `type` `default`, `series` or `copy`): copies of the source with numbers, dates, times, text with a number and names of days and months going on as Excel continues them ([docs/structure.md](structure.md#the-fill-handle)). Only the references that change are rewritten, in place; a formula that does not parse moves with its cell as written. An operation Excel refuses returns `{ ok: false, code, message }`, the message in the book's locale (`Нельзя изменить часть массива S!C1 (C1:C3)` in a ru-RU book), with `off_sheet` (cells would leave the sheet), `sheet_name` (a name Excel would not take, `History` among them, or one another sheet has), `last_sheet` (the last visible sheet deleted or hidden), `very_hidden` (a very hidden sheet deleted or moved), `merged` (a move cutting a merged area, a copy landing on part of one) `filtered` (a copy without the rows a filter hid, pasted over hidden rows of a filtered sheet) `array` (part of an old-kind array formula's area would change, or a merge would cover one), `table` (rows or columns into two tables at once, or deleted from two tables unless both go whole, a header row deleted alone, cells of a table merged), `table_name`, `outline_deep` (Group of lines at the deepest level), `not_grouped` (Ungroup of lines in no group, a group's button where there is none) or `detail` (Show Detail of a group shown, Hide Detail of one hidden). A name whose reference is lost gives `#REF!` where it is used, as in Excel; the parser reads the lost references Excel writes (`'S'!#REF!`, `#REF!A1`) as `#REF!`.

A book keeps styles as xlsx does: `styles`, a table of cell styles (font, solid fill, borders, alignment), `style` in a cell, an index into it, and per sheet `columns` and `rows` (by line `C`, `5` or span `B:D`, `5:9`: a width or height, a style, a number format) and `merges` (`B2:D3`). A cell without its own style or format shows its row's, else its column's. Computation ignores them, but a merged area stops an array spilling over it, as in Excel; the operations change them as Excel does: inserted rows take the format of the row above (its height, style and format, its cells' styles and formats) and columns of the column on the left, merged areas grow, shrink and move like ranges, moved and copied cells keep their look, whole rows and columns take their sizes and styles along. A cell's note and hyperlink go with the cell as its style does, frozen panes keep the rows and columns before their split, rules of data validation move as references do and lines inserted right after a rule take it, and inserted lines take the outline level of the line before them. The rules, and how they were checked against Excel on 317 books and the valid sets, merging on 61 more and the fill handle on 122 in Excel for Windows, are in [docs/structure.md](structure.md).

A session applies the same operations to its book with `session.editStructure(op)` and recomputes, keeping the book's styles, so a host can hold the book in the session alone and save `getBook()` as a version:

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

It answers `mapCell`, as `editStructure` does, but not the book, which `getBook()` gives; `moved`, how the addresses moved (`{ kind: 'lines', sheet, axis, at, count }` for rows and columns inserted or deleted, a range moved, a sheet renamed, deleted, inserted or moved); and `changes`, the cells whose value or formula text differs from those of the cell that was at their old place (`SUM(A1:A2)` above became `SUM(A1:A1)`), so an interface that moves its window by `moved` (or each cell by `mapCell`) and applies `changes` is up to date. It builds the state again, at the cost of a cold computation and a comparison of the cells with those of their old places (170 ms for the 3,400-row synthetic book), but for a sheet hidden or shown (`hideSheet`, `unhideSheet`), which changes no value, the outline commands, which recompute only the `SUBTOTAL` and `AGGREGATE` formulas over the rows they hide and show, and a merge that moves no cell, which writes the cells of its range as an edit does. Each operation is one step of the undo ([Undo and redo](sessions.md#undo-and-redo)).

## API

```ts
editStructure(book: BookInput, op: StructureOp): StructureResult
session.editStructure(op: StructureOp): SessionStructureResult   // `mapCell` and what an edit answers, without the book

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
  | { op: 'mergeCells'; sheet: string; range: string; across?: boolean; center?: boolean }
  | { op: 'unmergeCells'; sheet: string; range: string }

type RangeRef = { sheet: string; range: string }       // B2, B2:D9, B:D, 2:9
type StructureResult =
  | { ok: true; book: BookInput; mapCell(ref: CellRef): CellRef | null }
  | { ok: false; code: 'off_sheet' | 'sheet_name' | 'last_sheet' | 'very_hidden' | 'merged' | 'filtered' | 'array'; message: string }
```

The types are in `xlcalc/contract`.

An operation is plain data, so an interface can store it, send it to a server and replay it. Rows are numbers and columns are letters, as in A1 addresses. Sheet names are matched without regard to case.

The input book is not changed. Sheets, cells and names the operation does not touch are shared with it.

`mapCell` tells where a cell of the old book is now: a new address and sheet name, or `null` if the cell was deleted or overwritten, or if its sheet is none of the old book's (a window can hold a sheet that is gone). A host uses it for anything it keeps by address, such as comments, the selection or links to source rows.

Input of the wrong shape (an unknown sheet, row 0, a malformed range, unknown fields) throws `InputError`. An operation that Excel would refuse returns `ok: false` with a message for the user:

| code | when |
|---|---|
| `off_sheet` | an insert would push a cell or a merged area past row 1,048,576 or column XFD, or a move or copy does not fit on the sheet |
| `sheet_name` | the new name breaks Excel's rules or another sheet has it |
| `last_sheet` | deleting or hiding the last visible sheet (the only sheet too) |
| `very_hidden` | deleting or moving a very hidden sheet |
| `merged` | a move whose source cuts a merged area, or a copy that lands on part of one |
| `filtered` | a copy that left out rows a filter hid, pasted over hidden rows of a filtered sheet |
| `array` | lines inserted inside the area of an array formula of the old kind, part of one deleted or moved, a paste over part of one, a merge over one |
| `table` | a merge over cells of a table ([Merging cells](#merging-cells)); for the operations on tables see [docs/tables.md](tables.md) |

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

Names' formulas change like cells' formulas: each part of a union on its own, a reference without `$` not at all under inserted and deleted lines ([Names](names.md)). A reference across sheets stays under rows and columns and changes with its end sheets ([References](references.md)). The area of an array formula of the old kind moves with its cell; an operation that would change part of it is refused with `array` ([Legacy array formulas and implicit intersection](legacy.md)).

## Styles and layout

A book keeps styles the way an xlsx file does (`BookInput.styles`, `CellInput.style`, `SheetInput.columns`, `rows`, `merges`): a table of cell styles (font, solid fill, borders, alignment), a style number in each cell, and per sheet the columns and rows by line or span (`C`, `B:D`; `5`, `5:9`) with a width or height, a style and a number format, plus the merged areas. A cell's number format is its `numberFormat`. A cell shows its own style, else its row's, else its column's, and the same for the format. Widths are in characters as xlsx stores them, heights in points. Computation does not read styles, and they are not in the values hash; `inputHash` includes them when a book has any.

Excel resolves an empty cell's style and format together: from its row if the row has either, else from its column. The two can differ only where a row with a style but no format (or the other way round) crosses a column with the other; an xlsx reader spells such a case out on the row.

**Inserting and deleting.** Columns and rows keep their settings as they move. A span that holds the line before the inserted ones takes them in. A span that would reach past the sheet's end is cut there, and deleted lines leave their spans. The cells of inserted lines get the style and format of the cells of the line before them, a sheet's looks (`SheetInput.looks`) too: they move with their lines, and the new lines take the looks of the line before them, without an entry per cell.

A merged area changes like a range: it grows with lines inserted inside it, moves with lines inserted at or above its first line, stays when lines are inserted right after it (the new line is not merged), loses deleted lines, and disappears when one cell or none is left.

At the end of the sheet, a span of columns that reaches XFD keeps reaching it when columns are deleted, as xlsx stores such a span (`<col max=16384>`). If the deleted columns include XFD, the columns that come in at the end take the span's style and format but not its width (seen once). Rows are stored one by one and do not reach on. A merged area that ends on the last row or column keeps ending there when a deletion reaches the end of the sheet, unless all of it is deleted.

**Moving and copying.** A cell takes its own style and format along. Every pasted cell shows the style and format its source showed: where the target's rows and columns would show something else (a cell moved out of a bold column into a blue row), the cell gets the source's style spelled out, with `General` for no format and a plain style that is added to the book's table if missing. Whole rows or columns take their settings along; a move leaves the source lines plain, and a copy repeats them on every tile.

Merged areas inside the source go along. A copy also takes the part of a merged area that its source cuts, when that part has two cells or more. Merged areas that pasted cells land on are removed: all of them for a move, those inside the target for a copy.

**Refusals.** A move whose source cuts a merged area and a copy that lands on part of one are refused with `merged`, as Excel refuses them. A move that lands on part of a merged area removes it, and a copy whose source cuts one copies the part.

A renamed sheet keeps its layout, and a deleted sheet takes it along.

## Merging cells

`mergeCells` is Excel's Merge Cells: the range becomes one merged area (`SheetInput.merges`). With `across: true` each row of the range is merged on its own, as Merge Across does; with `center: true` the area is centred as well, as Merge & Center does. `unmergeCells` is Unmerge Cells.

`mergedArea(rect, merges)` of `xlcalc/a1` gives the rectangle a merge takes in, grown by each merged area it meets, for a grid that warns before a merge loses values.

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

The area keeps its top-left cell. Its other cells lose their values and formulas, as Excel merges once its warning ("Merging cells only keeps the upper-left value") is accepted. Formulas elsewhere that read a cleared cell keep reading it and now read an empty cell.

When the top-left cell is empty, the first cell of the area that holds a value or a formula, row by row, is moved into it, as cut and paste moves a cell: formulas and names that read it follow it, its own formula keeps its text, and its style and number format come with it. In `A1:C2` with A1 empty, B1 = 5 and A2 = 7, B1 moves to A1, `=B1*2` elsewhere becomes `=A1*2`, and A2 is cleared. As cut and paste over a cell, a reference to the empty A1 itself becomes `#REF!` (`=A1` is `=#REF!`), in a name and on another sheet too. A cell of a spill holds nothing of its own and is not moved. `mapCell` says where the moved cell went.

### The look

Every cell of the area takes the top-left cell's style and number format (its own, else its row's, else its column's), so that after Unmerge Cells the cells keep the look the area had. With `center`, that style gets a centred horizontal alignment, added to the book's style table if it has none such.

### Merged areas met

A range that meets a merged area takes it in whole, and grows again for each area it then meets: merging `D2:E2` beside the merged area `E2:F3` merges `D2:F3`, and merging cells inside a merged area merges it again. With `across`, the merged areas met are split into the rows of the range. Unmerge Cells splits every merged area that meets its range, and the cells keep their values and looks; a range without one changes nothing.

Whole rows and columns merge too: `'20:21'` is one area `A20:XFD21`, as in Excel.

### Refusals

A range that meets a table is refused with `table`: Excel does not merge the cells of a table. A range that meets an array formula of the old kind is refused with `array`.

## The fill handle

`fillRange` is Excel's fill handle (`Range.AutoFill`): the cells of `range` go on over `to`, which holds `range` and reaches past it down, up, right or left. A host that lets a user drag the handle sends the source and the whole frame of the drag.

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

The cells added take the source's cells in turn, as copy and paste lays them: formulas with their relative references moved, styles, number formats, notes and links, rules of conditional formats and data validation. When `to` is not a whole number of sources long, the part of the source nearest to it comes at the far end: `A1:A3` filled to `A1:A8` lays `A1:A3` over `A4:A6` and `A1:A2` over `A7:A8`. Refusals are a copy's (`merged`, `array`, `table`, `filtered`).

### What goes on as a series

Each column of a fill down or up, and each row of a fill right or left, goes on by itself. Its source is cut into runs: cells of one kind next to each other, empty cells between them skipped, and text without a number, booleans and formulas ending a run. A run's cells come again in every copy of the source, so a run goes on from copy to copy: `1, 2, x` goes on as `3, 4, x, 5, 6, x`, and `1, (empty), 3` as `5, (empty), 7`. A fill up or left goes backwards: `1, 2` filled up gives `0, -1` above.

- Numbers: two or more go on by their step, or by their least-squares trend when the steps differ (`1, 3, 4` gives `5.67, 7.17`). A single number is copied when the whole source holds plain numbers and empty cells alone, and goes on by 1 when the source holds anything else as well (text, a boolean, a formula, a date).
- Dates (numbers shown in a format with a day, a month or a year): one date goes on by days; dates on one day of the month, or each on the last day of its month, a whole number of months apart go on by months (`Jan 31, Feb 28` gives `Mar 31, Apr 30`; `Jan 15, Apr 15` gives `Jul 15`); other dates with one step in days go on by it, and dates without one step are copied.
- Times of day: one time goes on by hours; times with one step go on by it, others are copied.
- Text with a number: a number at the start followed by a space goes on (`10 kg`, `11 kg`; `1 item 2`, `2 item 2`), else a number at the end (`Item1`, `Item2`; `Item09`, `Item10`), with the step of the run, and a run without one step is copied. A number below zero is written by its size (`Item1` filled up gives `Item0`, `Item1`, `Item2`). Quarters go round the year (`Q4`, `Q1`; `Qtr`, `Quarter 1`), en-US ordinals take their endings (`1st`, `2nd`, `3rd`, `11th`, always small). Text with the number inside (`a1b`) or right before letters (`1abc`) is copied.
- Names of days and months, short or long (`Mon`, `Monday`, `Jan`, `January`), go round their list in the case they are written in (`MONDAY`, `TUESDAY`); en-US's names, and in a book of another locale its own names too.
- Everything else is copied: text, booleans, errors, formulas.

`type` changes this as Excel's `XlAutoFillType` does: `'series'` makes a single number go on by 1 as well (Excel's drag with Ctrl over a number), `'copy'` copies every cell.

`mapCell` of a fill is a copy's: the cells it writes over are cleared, the others stay.

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
- Deleted lines leave the ranges; a range that loses every line is gone, and a rule that loses every range. A range ending on the sheet's last row or column keeps ending there, one cell too, unlike a reference to a cell: a rule over `B1048576` is over `B1048575:B1048576` after a row above is deleted, as a conditional format's is.
- A move takes the rules of the moved cells along. The cells it lands on keep their own rules where the moved cells have none, and take the moved ones where they have some. A move from `A1`, which has no rule, onto a cell of a whole-number rule leaves that cell's rule.
- A copy pastes the rules of the copied cells on every tile, and a pasted cell without a rule clears the target's rule, as Excel's Paste All does.
- References in a rule's formulas change as a cell's formulas do: a list `'L'!$A$1:$A$4` follows rows inserted on L, becomes `'Lists'!$A$1:$A$4` when L is renamed and `#REF!$A$1:$A$4` when it is deleted. A rule whose first cell moves or goes has its formulas written again for its new first cell: deleting row 2 of a rule `D2:D6` with `D2<A2` leaves `D2:D5` with `D2<A2`, the old D3's formula. A rule moved to another sheet moves its formulas as a copy moves them: `B2>A2` moved from B2 of S to D2 of T is `D2>C2` there, while a move on its own sheet keeps the references to the cells that did not move (`D2<A2` moved to D20 is `D20<A2`).

An outline is the outline levels of rows and columns (`outlineLevel`, 1 to 7) with their collapsed marks (`collapsed`), and the sheet's summary settings (`SheetInput.outline`, below and right by default). A collapsed group's lines are hidden lines, which `SUBTOTAL` 101 to 111 leaves out.

- Inserted lines take the outline level of the line before them, as they take its height, and are shown. Lines inserted into a collapsed group after one of its hidden lines open it: its summary line, past the group (before it with summaries above or left), loses its collapsed mark. Lines inserted at the summary line, or above the group, leave it collapsed.
- Deleted lines take their levels and marks along; moved and copied whole rows and columns carry them, as they carry their other settings.

Excel's outline commands, Group, Ungroup, Show and Hide Detail and the level buttons, are structure operations too, which move no cell: [docs/outline.md](outline.md) has them.

The valid sets `struct-held` and `struct-held-2` measured these rules in Excel for Mac and for Windows, with the engine's answers committed before Excel ran them ([Development](development.md#data-validation)).

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

`session.editStructure(op)` applies the operation to the session's book and recomputes. It answers `mapCell`, as `editStructure` does, without the book, which `getBook()` gives, and what every edit of a session answers ([Editing a grid](edits.md#what-an-edit-answers)): `moved`, how the addresses moved, and `changes`, every cell whose value or formula text differs from those of its old cell. An interface that moves its window by `moved` (or each cell by `mapCell`) and applies `changes` has the new values and formulas. An insert into a block of data lists the formulas over it, which read one row more, and no value. A refused operation changes nothing.

The session rebuilds its state from the new book. On the synthetic book (3,400 rows, 3,907 formulas) inserting a row in the middle of the data takes 170 ms, about as long as a cold computation of the same book, well within the 1.5 s budget. On the book of 34,000 rows (506,594 cells on the data sheet) inserting or deleting a row takes 1.3 to 1.5 s, its undo and its redo 1.4 to 1.5 s, a row copied down 2,000 rows 1.2 s and a block of 100 by 20 cells moved 1.1 s; building the state is half of it, the operation itself 0.25 s.

`mapCell`, `moved`, the way an undo and a redo take the operation back and again and the comparison of each cell with its old cell all follow one description of where the operation took the cells: lines inserted or deleted, a range moved, cells pasted over, a sheet renamed, inserted, deleted or moved. Rows deleted on a filtered sheet are one `lines` move for each run of visible rows, from the bottom up, and their undo puts each run back; an undo's `mapCell` is the operation's taken back, a redo's the operation's own.

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

A merge that moves no cell changes the cells of its range, the merged areas and the style table, as a patch of looks does, without building the state again: on the book of 34,000 rows, merging 3 by 3 cells of the data that 2,800 SUMIFS read takes 0.17 s and its undo as long, nearly all of it computing the formulas over the cleared cells. A merge that moves a cell into an empty top-left cell, or that meets an array formula or more than 65,536 cells, builds the state again, as an insert does (1.5 s on that book).

Only two of these recompute values. A merged area added or removed over an array's spill range recomputes that array. Rows hidden or shown, and a new filter mode, recompute the `SUBTOTAL` and `AGGREGATE` formulas over the rows concerned and what depends on them.

## Checked against Excel

Measured in Excel 16 for Mac, en-US, on 317 books; merging cells (61 cases) and the fill handle (122 cases) in Excel for Windows, hidden sheets in both. All match but the known differences below. Not measured: a move that does not fit, most moves to another sheet, fonts other than bold, borders and alignment; for merging, borders, the vertical alignment and the notes and links of the cells it clears; the fill handle in Excel for Mac and with a ru-RU Excel's lists.

## Known differences and limits

- Whole columns copied onto part of a merged area: Excel pastes them and keeps the area, the engine refuses with `merged` (two observations).
- Excel's fill handle fills the hidden rows of a filtered sheet. `copyRange` fills the visible rows only, as Fill Down does, and so does `fillRange`, whose copies are `copyRange`'s.
- The fill handle was not measured in Excel for Mac, nor in a ru-RU Excel (its lists of names), nor with the types other than default, series and copy. Not measured either: whether a number at the end of text keeps its zeros past one width (`Item001`, here `Item002`), and a time beside a number alone (here it makes the number go on, as a date does).
- One move rule rests on a single observation: a straddled end does not follow into the overwritten cells of a move that does not overlap itself.
- Charts are not part of a book. Threads of comments are notes in a book ([docs/xlsx.md](xlsx.md#threads-of-comments)).
- References to other workbooks are not read, so they never occur in a valid formula.
