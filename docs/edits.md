# Editing a grid

A grid over a workbook edits cells the way Excel does: a value typed over a cell keeps the cell's formatting, Delete clears the contents, Clear Formats clears the formatting, Bold over a selection is applied on top of what each cell displays, and typed text is parsed in the workbook's language. A session performs each of these in one call, recalculates what depends on it, records it as one undo step and returns what the grid has to redraw, so a host needs no code of its own between what the user does and the workbook.

The formatting and clearing rules below were verified in Excel 16.113 for Mac and Excel 16.0.20430 for Windows, en-US.

An edit's `changes` include `note` and `link` when their fields change, and null when they are removed, including on an otherwise empty cell. `setCells`, `replaceCells`, `patchCells`, `clearCells`, `undo` and `redo` follow the same rule. Equal notes or hyperlinks do not produce these fields. `meta` remains host data: it is saved in the workbook and restored by the history, but its changes are not listed in `CellChange`.

## In short

| to | call |
|---|---|
| put a value or a formula into cells, keeping their formatting (typing over a cell, pasting values) | `setCells([{ sheet, cell, input: { value } }])` |
| replace cells entirely, or remove them (`input: null`) | `replaceCells([{ sheet, cell, input }])` |
| set the formatting, a note or a hyperlink of cells, a range, entire rows or columns, without touching their values (Bold, Fill, a number format) | `patchCells([{ sheet, range, set: { style, numberFormat } }])` |
| clear the contents, the formatting or both (Delete, Clear Formats, Clear All) | `clearCells([{ sheet, range }], 'contents')`, `'formats'` or `'all'` |
| accept text the user typed into a cell, parsed as Excel parses it in the workbook's locale | `setCellText(sheet, cell, text)` |
| accept text typed into several cells, each parsed as if typed into its own cell, in one step | `setCellTexts([{ sheet, cell, text }])` |
| paste text from the clipboard at a cell, as Excel pastes it | `pasteText(sheet, cell, text)` |
| enter text into every cell of a selection (Control+Return) | `fillCellText(sheet, range, text, { active })` |
| change widths, heights, hidden rows and columns, a row's or column's style, merged areas, frozen panes, conditional formats, data validation, or AutoFit columns | `setLayout([{ sheet, columns, set: { width } }])` ([Layout](grid.md#layout)) |
| replace the workbook's style table | `setStyles(styles)` |
| insert or delete rows and columns, move or copy a range, fill it as the fill handle does, merge or unmerge cells, insert, rename, move, hide or delete a sheet, add, rename, convert or delete a table | `editStructure(op)` ([Structure operations](structure.md)) |

`patchCells` accepts a style given by its fields, `{ font: { bold: true } }`: it is applied on top of the style each cell displays, and the session finds or adds the resulting style in the workbook's `styles`. `setCells` and `replaceCells` accept a cell's input as the workbook stores it, with a style given by its index in `styles`, as `getCells`, `getViewport` and the changes report it. A host that keeps a workbook's styles itself can find a style's index by the same rule: `internStyle(styles, style)` returns the index of the first equal style and the same array, or a new array with a copy of the style appended; styles are equal when their fields are, in any order.

```js
import { internStyle } from 'xlcalc'

const styles = [{ font: { bold: true } }, { fill: { color: '#FFEE99' }, font: { italic: true } }]
const found = internStyle(styles, { font: { italic: true }, fill: { color: '#FFEE99' } })
console.log(found.index, found.styles === styles) // → 1 true
const added = internStyle(styles, { font: { bold: true, italic: true } })
console.log(added.index, added.styles.length, styles.length) // → 2 3 2
```

```js
import { createSession } from 'xlcalc'

const session = createSession({
  styles: [{ font: { bold: true } }],
  sheets: [{
    name: 'Plan',
    cells: { A1: { value: 'Item', style: 0 }, A2: { value: 2, numberFormat: '0.00', note: { text: 'from the CRM' } }, A3: { formula: 'A2*10' } },
    conditionalFormats: [{ type: 'cellIs', ranges: ['A2:A3'], operator: 'greaterThan', formulas: ['20'], style: { font: { bold: true } } }],
  }],
})
const edit = session.setCells([{ sheet: 'Plan', cell: 'A2', input: { value: 5 } }])
console.log(edit.revision, edit.repaint, session.getBook().sheets[0].cells.A2)
// → 1 [ { sheet: 'Plan', range: 'A2:A3' } ] { numberFormat: '0.00', note: { text: 'from the CRM' }, value: 5 }
const fill = session.patchCells([{ sheet: 'Plan', range: 'A1:B1', set: { style: { fill: { color: '#FFEE99' } } } }])
console.log(fill.changes.map((c) => `${c.cell} ${c.style}`), fill.styles)
// → [ 'A1 1', 'B1 2' ] [
//     { font: { bold: true } },
//     { font: { bold: true }, fill: { color: '#FFEE99' } },
//     { fill: { color: '#FFEE99' } }
//   ]
session.clearCells([{ sheet: 'Plan', range: 'A2:A3' }], 'contents')
console.log(session.getBook().sheets[0].cells.A2) // → { numberFormat: '0.00', note: { text: 'from the CRM' } }
const stale = (() => {
  try {
    session.setCells([{ sheet: 'Plan', cell: 'A3', input: { value: 1 } }], { baseRevision: 1 })
  } catch (e) {
    return e.code
  }
})()
console.log(stale, session.revision) // → stale_revision 3
```

`setCells(edits)` sets the fields an edit specifies and keeps the cell's other fields: null removes a field, and the contents count as one field, so a new `value` replaces a formula and its `array`, and a new `formula` replaces a value. `input: null` clears the contents and the hyperlink and keeps the style, the number format, the note and `meta`, like Excel's Delete key. `replaceCells(edits)` replaces each cell entirely, and null removes it. `clearCells(ranges, what)` clears rectangles, including entire rows and columns, like Excel's ClearContents (`'contents'`: values, formulas and hyperlinks), ClearFormats (`'formats'`: each cell's style and number format, so that it displays the default formatting whatever its row or column has, and the conditional formats over the cells) and Clear (`'all'`: both, plus notes, `meta` and the data validation rules over the cells); entire rows and columns also lose their own style and number format.

`patchCells(patches)` sets the `style`, `numberFormat`, `meta`, `note` or `link` of a `cell` or a `range` (`B2:D9`, entire columns `B:D`, entire rows `2:5`; null removes the field) and keeps values and formulas without having to be given them. On a cell in the range of a legacy array formula or of a spill it leaves the array as it is, a cell left with nothing is removed, and it does not recalculate anything. A style given by its fields is applied on top of the style each cell displays (its own, otherwise its row's, otherwise its column's), as Excel formats a selection: `{ font: { bold: true } }` makes an italic cell bold and italic, null removes a field (`{ fill: null }`), and equal styles share one entry in the workbook's `styles`, found or added (the result's `styles` is then the new table). A style given by its index replaces the cell's style. An empty cell of a rectangle gets a cell with the formatting, up to 1,048,576 cells (one column's worth; more throws an `InputError`). Entire rows and columns are formatted as Excel formats them: the lines' own style and number format, the cells in them that have formatting of their own, and the cells where they cross a formatted line of the other direction, which get the two combined.

Besides `revision` and `changes`, a result has `layout`, the sheets whose layout changed (rows and columns, merged areas, tables, rules), so that a grid can call `getLayout()` again; `styles`; `moved`, after a structure operation and its undo; and `repaint`, the ranges of conditional formats whose cells may display differently after a value or a rule changed. An edit made against a revision other than the session's (`baseRevision`, taken from a result, `session.revision` or a viewport's `getViewport(…).revision`) changes nothing and throws `CalcError` with the code `stale_revision`, so a host does not apply an edit on top of one it has not seen. Each call is one undo step. [docs/edits.md](edits.md) has the rules and how they were verified against Excel.

## Editing a cell

`setCells(edits)` sets the fields an edit specifies and keeps the cell's other fields: a new value keeps the style, the number format, the note, the hyperlink and `meta`.

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 2, numberFormat: '0.00', note: { text: 'n' } } } }] })
session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 5 } }])
// the number format and the note stay
console.log(session.getBook().sheets[0].cells.A1)
// → { numberFormat: '0.00', note: { text: 'n' }, value: 5 }
```

- A field set to null is removed: `{ style: null }` leaves the cell with the formatting of its row or column.
- The contents count as one field: a new `value` replaces a formula and its `array`, and a new `formula` replaces a value. `{ formula: null }` removes the formula and its range.
- `input: null` clears the contents and the hyperlink and keeps the rest, like Excel's Delete key and ClearContents. A cell left with nothing is removed.
- A formula with a leading `=` is accepted without it, as everywhere a workbook is passed in.

Everything else works as before: an edit that changes part of a legacy array formula's range throws an `InputError`, a value in a spill range blocks the array, and typing over a table's header cell renames its column ([Tables](tables.md)).

`replaceCells(edits)` replaces each cell entirely: whatever the input leaves out, the cell loses, and null removes the cell.

## Clearing

`clearCells(ranges, what)` clears rectangles of a sheet (`B2:D9`), entire columns (`B:D`) and entire rows (`2:5`), like Excel's three commands:

| `what` | Excel | the cells lose | they keep |
|---|---|---|---|
| `'contents'` | ClearContents | values, formulas, hyperlinks | styles, number formats, notes, `meta` |
| `'formats'` | ClearFormats | styles, number formats; the cells of conditional formats over them | values, formulas, notes, hyperlinks, `meta` |
| `'all'` | Clear | all of these, notes, `meta`; the cells of data validation rules over them | nothing |

A cell whose formats are cleared displays the workbook's default formatting, whatever its row or column has: Excel stores it with the Normal style. Where its row or column has a style, the cell gets an empty style of its own (a style with nothing set, added to the workbook's table if missing), and where they have a number format, `General`. An empty cell of a rectangle in a formatted row or column gets such a cell too, and so does an empty cell where a cleared column crosses a formatted row. Entire rows and columns cleared of their formats lose their own style and number format.

A conditional formatting rule, or a data validation rule for `'all'`, loses the cleared cells: its ranges become what is left (`B2:D4` cleared of `B2:D2` becomes `B3:D4`), a rule left without cells is removed, and a rule whose first cell is removed has its formulas rewritten for its new first cell (`B2>3` over `B2:D4` becomes `B3>3` over `B3:D4`), so each remaining cell keeps its condition.

A host's `meta` is the cell's identity rather than its contents or its formatting, so only `'all'` removes it.

## Formatting

`patchCells(patches)` sets the formatting of a `cell` or a `range`: a rectangle, entire columns or entire rows. A patch sets `style`, `numberFormat`, `meta`, `note` and `link`; null removes a field; values and formulas stay.

### A style by its fields

A style given by its fields is applied on top of the style each cell displays, as Excel formats a selection: the cell's own style, otherwise its row's, otherwise its column's.

```js
import { createSession } from 'xlcalc'

const session = createSession({
  styles: [{ font: { italic: true } }, { fill: { color: '#FFFF00' } }],
  sheets: [{ name: 'S', columns: { C: { style: 1 } }, cells: { A1: { value: 'x', style: 0 } } }],
})
const { styles } = session.patchCells([{ sheet: 'S', range: 'A1:C3', set: { style: { font: { bold: true } } } }])
// A1 italic: bold and italic; B2 empty: bold; C1 in a filled column: bold and filled
console.log(['A1', 'B2', 'C1'].map((cell) => styles[session.getCell('S', cell).style]))
// → [
//     { font: { italic: true, bold: true } },
//     { font: { bold: true } },
//     { font: { bold: true }, fill: { color: '#FFFF00' } }
//   ]
```

- The font, the sides of the border and the alignment change field by field; the fill is replaced as a whole. `{ font: { bold: null } }` removes bold, and `{ fill: null }` the fill.
- Equal styles share one entry in the workbook's `styles`: the session finds an equal style or appends one, and the result's `styles` is then the new table. In the workbook, styles remain indexes.
- A style given by its index (`style: 3`) replaces the cell's style entirely, as before.
- A style left with nothing is the empty style, which keeps a cell unformatted in a formatted row or column.

### Rectangles

Every cell of a rectangle gets the formatting, and each empty cell gets an entry of its own, as Excel stores a cell it has formatted. A rectangle can have as many cells as one column (1,048,576); a larger one throws an `InputError`, while entire columns or rows are formatted at the cost of the cells they contain.

### Entire rows and columns

Entire rows and columns are formatted as Excel formats them:

- each line's own style takes the new fields (and its number format the new one), span by span;
- each cell in the lines that has a style or number format of its own gets the change on top of its own;
- each cell where the lines cross a line of the other direction that has a style (or a number format, when a number format is set) gets an entry of its own with what it displayed plus the patch. A row made bold across a filled column makes the cells where they cross bold and filled; a column made italic across a bold row makes them bold and italic.

A cell that displays its row's or column's formatting without formatting of its own gets nothing: the line's new style shows through. On entire rows and columns, `meta`, `note` and `link` can only be removed (null), from the cells in those lines.

## Typing text

`setCellText(sheet, cell, text, opts?)` parses text typed into a cell as `parseEntry` parses it ([Cell entry](entry.md)): in the workbook's locale and date system, with the session's clock (for the year of a date typed without one) and by the number format the cell displays, its own, otherwise its row's, otherwise its column's.

- A number format the cell displays from its row or column, and which the entry keeps, is not written into the cell: `5` typed into a column formatted as percentages becomes 5%, and the cell has no number format of its own.
- A cell's own General format stays its own.
- The cell keeps its style, note, hyperlink and `meta`.
- A number, a date or a formula's numeric result too wide for a column of the default width, or of a width Excel fitted, widens the column as Excel does (`entryWidth`, [Fitting columns](formats.md#fitting-columns)), in the same undo step; `width` in the result is the width, and `widen: false` returns it without applying it. A hidden column and a cell merged across columns are left unchanged.
- Text with a line break wraps the cell (its style, its own or otherwise its row's or column's, gets `alignment.wrap`, with the style found in or added to the table; the result's `styles` is the new table), and the row of a wrapping cell is fitted as AutoFit fits it, unless its height was set manually or it is hidden, in the same step ([A line break](entry.md#a-line-break), [Fitting rows](formats.md#fitting-rows)).
- A value the cell's data validation rule does not allow is still written, with `validation`: `{ valid: false, rule, severity, message }`, where the message is the rule's `error` or the library's text in the workbook's locale. With `refuseInvalid: true`, such a value is rejected instead (`code: 'validation'`, the cell, the rule and its style) and nothing changes: the host shows Excel's alert first and calls again without the option if the user keeps the value (Yes on a warning, OK on an information alert); for a rule with the style `stop`, Excel does not accept the value at all. Without the option, the host calls `undo()` if the user cancels. A rule that shows no alert (`showError: false`) returns none.
- A formula that does not parse is rejected with its error code, its position in the text and Excel's message in the workbook's locale; so are a change to part of a legacy array formula (`array`) and a formula in a table's header cell (`header`). Text typed into a header cell renames its column.

`previewCellText` returns the same information, the cell's input, its value, the width and the alert, and changes nothing.

`setCellTexts(entries, opts?)` accepts several texts at once, as pasting text or entering it into a selection does, each `{ sheet, cell, text }` parsed as `setCellText` would parse it in that cell:

- The texts are parsed in order, so a cell given twice keeps the later text, parsed with the number format the earlier one gave it: `15.03.2026` followed by `45` leaves the cell at 45 with the date format.
- All are written with one recalculation, one revision and one undo step.
- A column widens once, to the largest width that any one of the texts typed alone into one of its cells would give it, measured with the values all the texts produce: `widths` lists `{ sheet, column, width }`, and `widen: false` returns them without applying them.
- `validations` lists the cells whose values their data validation rules do not allow, each with its alert, like the `validation` of `setCellText`, after all the texts are written.
- A text that `setCellText` would reject rejects the whole call with its code, message and position, plus its `sheet` and `cell`; nothing is written.

`previewCellTexts` returns each cell's input and value, the widths and the alerts, and changes nothing.

`pasteText(sheet, cell, text, opts?)` pastes text from the clipboard as Excel pastes it, which in some places parses text differently from typing ([A paste and Control+Return](entry.md#a-paste-and-controlreturn)):

- The text fills a rectangle starting at the cell, with rows split at line breaks and fields at tabs (`parseClipboard`). A field that starts with a quote runs to its closing quote, can contain line breaks and wraps its cell. An empty field, the missing fields of a short row and the row after a final line break clear their cells. The result's `range` is the rectangle.
- Each field is parsed as if typed into its cell, except that a leading apostrophe stays part of the text, a formula Excel does not accept becomes text, a percentage cell has no percentage entry, and `€5` gets the euro format.
- No column widens and no data validation rule is checked.
- Part of a legacy array formula, a formula in a table's header cell and a rectangle extending past the edge of the sheet (`off_sheet`) are rejected with their `sheet` and `cell`; nothing is written.

`fillCellText(sheet, range, text, opts?)` enters text into every cell of a range, like Excel's Control+Return:

- The text is parsed once, in the active cell (the range's top-left cell, or `active`) by its number format, and every cell gets what the active cell contains afterwards: the value, or the formula adjusted as copying it from the active cell would adjust it, plus the active cell's number format if it has one.
- The columns widen as typing widens them (`widths`; `widen: false` returns them without applying them).
- Only the active cell's data validation rule is checked: `validation` is its alert, as for `setCellText`.
- A formula Excel does not accept is rejected with the active cell, and part of an array or a formula in a header cell with that cell; nothing is written.

`previewPasteText` and `previewFillCellText` return each cell's input and value (and, for a fill, the widths and the alert) and change nothing.

`validateEntry(sheet, cell, input)` returns `{ ok: true }` or `{ ok: false, code: 'validation', message, rule, severity }`.

## What an edit returns

Every call that changes the workbook (`setCells`, `replaceCells`, `clearCells`, `patchCells`, `setCellText`, `setCellTexts`, `pasteText`, `fillCellText`, `setNames`, `setClock`, `setStyles`, `setLayout`, `editStructure`, `undo`, `redo`) returns an `EditResult`:

- `revision`: the session's revision after the call. Each such call increments it by one, `undo` and `redo` included; a rejection, a failure and an undo with nothing to undo do not.
- `changes`: the cells whose value or displayed text changed, and those whose formula, own style or own number format changed, with `formula`, `style` and `numberFormat` in that case (null when the cell no longer has one), and a new number format also as a code, `numberFormatCode`: a built-in format by its code in the workbook's locale (14 is `dd.mm.yyyy` in ru-RU), as a viewport gives it. After a structure operation, also the cells whose value or formula text differs from that of the cell previously at their position: after a row is inserted, `SUM(A1:A4)` that became `SUM(A1:A5)` is listed with its new text.
- `layout`: the sheets whose layout changed (rows and columns, merged areas, tables, rules): a host calls `getLayout()` again.
- `styles`: the workbook's style table, when the call changed it.
- `moved`: after a structure operation and its undo, how addresses moved, in order: `{ kind: 'lines', sheet, axis, at, count }` (the rows or columns from `at` onward move by `count`, up or left when it is negative), `{ kind: 'range', from, to }`, and a sheet renamed, deleted, inserted or moved. Rows deleted on a filtered sheet are one `lines` move for each run of visible rows, from the bottom up. A host shifts its viewport by them rather than mapping each cell with `mapCell`.
- `repaint`: the ranges of the workbook's conditional formats, after a call that changed a value or a rule. A rule's formatting depends on values wherever its formulas refer, so a host reads its viewport again where it overlaps them. A large change of formatting comes the same way: a rectangle of more than 65,536 cells whose formatting changed is a range of `repaint`, not cells of `changes`, after `patchCells` and `clearCells` (Bold or Clear Formats over entire columns of formatted rows), a structure operation (Merge Cells on entire columns), and their undo and redo. A cell with an entry of its own (a value, a formula, a note) is still listed. Formatting stored by ranges changes by ranges, so such an edit costs what its ranges and its cells with entries hold rather than its area; the cells where entire rows cross a formatted column, or entire columns a formatted row, still get entries one by one.

## Revisions

`session.revision` is the current revision (for an asynchronous session, as of its last completed call), and `getViewport` returns the revision of its viewport. An edit given `baseRevision` checks it first: if the session's revision differs, the edit changes nothing and throws `CalcError` with the code `stale_revision`. Calls to one session run in order, so a host with one view needs no revisions; one that applies edits from several views does. A session restored on a new thread of a pool comes back at its revision.

## Undo

Each call is one step: `clearCells` and `patchCells` regardless of their parts (the cells, the rows and columns, the rules, the style table), `setCellText`, `setCellTexts` and `fillCellText` including the column widths they set, and `pasteText` regardless of its rectangle. A call with nothing to do (`setCells([])`) is not a step; a call given cells that ended up unchanged is. The undo's result is the call's `EditResult` with `call`, plus `cells` for the calls that change cells.

## Checked against Excel

Formatting and clearing were tested in Excel 16.113 for Mac and Excel 16.0.20430 for Windows on 34 workbooks (1,967 cells), with the results predicted before Excel ran them; the engine displays what both versions displayed in every cell. Typed text, pastes and Control+Return are checked as described in [Cell entry](entry.md).
