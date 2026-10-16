# Editing a grid

A grid over a book edits cells the way Excel does: a value typed over a cell keeps the cell's look, Delete clears the content, Clear Formats the look, Bold over a selection goes over what each cell shows, and typed text is read in the book's language. A session does each of these in one call, recomputes what depends on it, keeps it as one step of the undo and answers what the grid has to draw again, so a host needs no code of its own between what a person does and the book.

The rules of formatting and clearing below were measured in Excel 16.113 for Mac and Excel 16.0.20430 for Windows, en-US.

An edit's `changes` includes `note` and `link` when their fields change, and null when they are removed, including on an otherwise empty cell. `setCells`, `replaceCells`, `patchCells`, `clearCells`, `undo` and `redo` use the same rule. Equal notes or links do not produce these fields. `meta` remains host data: it is saved in the book and restored by history, but its changes are not listed in `CellChange`.

## In short

| to | call |
|---|---|
| put a value or a formula into cells, keeping their look (typing over a cell, a paste of values) | `setCells([{ sheet, cell, input: { value } }])` |
| replace cells whole, or remove them (`input: null`) | `replaceCells([{ sheet, cell, input }])` |
| set the look, a note or a link of cells, a range, whole rows or columns, without their values (Bold, Fill, a number format) | `patchCells([{ sheet, range, set: { style, numberFormat } }])` |
| clear the content, the look or both (Delete, Clear Formats, Clear All) | `clearCells([{ sheet, range }], 'contents')`, `'formats'` or `'all'` |
| take text a person typed into a cell, as Excel reads it in the book's locale | `setCellText(sheet, cell, text)` |
| take texts typed into several cells, each read as typed into its cell, in one step | `setCellTexts([{ sheet, cell, text }])` |
| paste text from the clipboard at a cell, as Excel pastes it | `pasteText(sheet, cell, text)` |
| enter a text into every cell of a selection (Control+Return) | `fillCellText(sheet, range, text, { active })` |
| change widths, heights, hidden lines, a line's style, merged areas, frozen panes, conditional formats, data validation, or AutoFit columns | `setLayout([{ sheet, columns, set: { width } }])` ([Layout](grid.md#layout)) |
| replace the book's table of styles | `setStyles(styles)` |
| insert or delete rows and columns, move or copy a range, fill it as the fill handle does, merge or unmerge cells, insert, rename, move, hide or delete a sheet, add, rename, convert or delete a table | `editStructure(op)` ([Structure operations](structure.md)) |

A style by its fields, `{ font: { bold: true } }`, is what `patchCells` takes: it goes over the style each cell shows, and the session finds or adds the style in the book's `styles`. `setCells` and `replaceCells` take a cell's input as the book stores it, a style by its number in `styles`, as `getCells`, `getViewport` and the changes give it. A host that keeps a book's styles itself finds a style's number by the same rule: `internStyle(styles, style)` gives the first equal style's number and the same array, or a new array with a copy of the style at the end; styles are equal when their fields are, in any order.

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

`setCells(edits)` sets the fields an edit gives and keeps the cell's others: null removes a field, and the content is one, so a new `value` takes the place of a formula and its `array`, a new `formula` that of a value. `input: null` clears the content and the hyperlink and keeps the style, the format, the note and `meta`, as Excel's Delete key. `replaceCells(edits)` replaces each cell whole, null removing it. `clearCells(ranges, what)` clears rectangles, whole rows and columns too, as Excel's ClearContents (`'contents'`: values, formulas and hyperlinks), ClearFormats (`'formats'`: each cell's style and number format, so that it shows the plain look whatever its row or column has, and the conditional formats over the cells) and Clear (`'all'`: both, and notes, `meta` and the rules of data validation over the cells); whole rows and columns lose their own style and format too.

`patchCells(patches)` sets the `style`, `numberFormat`, `meta`, `note` or `link` of a `cell` or a `range` (`B2:D9`, whole columns `B:D`, whole rows `2:5`; null removes the field) and keeps values and formulas without being given them: on a cell of an old-kind array formula's area or of a spill it leaves the array as it is, a cell left with nothing is removed, and it computes nothing. A style given by its fields goes over the style each cell shows, its own, else its row's, else its column's, as Excel formats a selection: `{ font: { bold: true } }` makes an italic cell bold and italic, null removes a field (`{ fill: null }`), and equal styles share one entry of the book's `styles`, found or added (`styles` in the answer is the table then); a style by its number replaces the cell's. An empty cell of a rectangle gets a cell with the look, up to 1,048,576 cells (one column's worth; more is an `InputError`). Whole rows and columns are formatted as Excel formats them: the lines' own style and format, the cells of theirs with a look of their own, and the cells where they cross a line of the other direction with one, which get the two merged.

Beside `revision` and `changes`, an answer has `layout`, the sheets whose layout changed (lines, merged areas, tables, rules), for a grid to read `getLayout()` again; `styles`; `moved`, after a structure operation and its undo; and `repaint`, the rectangles of conditional formats whose cells may look otherwise after a value or a rule changed. An edit made over another revision than the session's (`baseRevision`, from an answer, `session.revision` or a window, `getViewport(…).revision`) changes nothing and throws `CalcError` with the code `stale_revision`, so that a host does not apply an edit over one it has not seen. Each call is one step of the undo. [docs/edits.md](edits.md) has the rules and how they were checked against Excel.

## Editing a cell

`setCells(edits)` sets the fields an edit gives and keeps the cell's others: a new value keeps the style, the number format, the note, the hyperlink and `meta`.

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 2, numberFormat: '0.00', note: { text: 'n' } } } }] })
session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 5 } }])
// the number format and the note stay
console.log(session.getBook().sheets[0].cells.A1)
// → { numberFormat: '0.00', note: { text: 'n' }, value: 5 }
```

- A field set to null is removed: `{ style: null }` leaves the cell the look of its row or column.
- The content is one: a new `value` takes the place of a formula and its `array`, a new `formula` the place of a value. `{ formula: null }` removes the formula and its area.
- `input: null` clears the content and the hyperlink and keeps the rest, as Excel's Delete key and ClearContents do. A cell left with nothing is removed.
- A formula with a leading `=` is taken without it, as everywhere a book is given.

Everything else is as before: an edit that changes part of an old-kind array formula's area is an `InputError`, a value in a spill range blocks the array, a table's header cell typed over renames its column ([Tables](tables.md)).

`replaceCells(edits)` replaces each cell whole: what an input leaves out the cell loses, and null removes the cell.

## Clearing

`clearCells(ranges, what)` clears rectangles of a sheet (`B2:D9`), whole columns (`B:D`) and whole rows (`2:5`), as Excel's three commands do:

| `what` | Excel | the cells lose | they keep |
|---|---|---|---|
| `'contents'` | ClearContents | values, formulas, hyperlinks | styles, number formats, notes, `meta` |
| `'formats'` | ClearFormats | styles, number formats; the cells of conditional formats over them | values, formulas, notes, hyperlinks, `meta` |
| `'all'` | Clear | all of these, notes, `meta`; the cells of rules of data validation over them | nothing |

A cell whose formats are cleared shows the book's plain look, whatever its row or column has: Excel stores it with the Normal style. Where its row or column shows a style, the cell gets the plain style of its own (a style with nothing set, added to the book's table if it has none), and where they show a number format, `General`. An empty cell of a rectangle in a formatted line gets such a cell too, and so does an empty cell where a cleared column crosses a formatted row. Whole rows and columns cleared of their formats lose their own style and number format.

A rule of conditional formatting, or of data validation for `'all'`, loses the cells cleared: its ranges are what is left (`B2:D4` cleared of `B2:D2` is `B3:D4`), a rule left without cells goes, and one whose first cell goes has its formulas written for its new first cell (`B2>3` over `B2:D4` becomes `B3>3` over `B3:D4`), so each remaining cell keeps its condition.

A host's `meta` is the cell's identity rather than its content or its look, so it goes with `'all'` alone.

## Formatting

`patchCells(patches)` sets the look of a `cell` or a `range`: a rectangle, whole columns or whole rows. A patch sets `style`, `numberFormat`, `meta`, `note` and `link`; null removes a field; values and formulas stay.

### A style by its fields

A style given by its fields goes over the style each cell shows, as Excel formats a selection: the cell's own style, else its row's, else its column's.

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

- The font, the border's sides and the alignment change field by field; the fill changes whole. `{ font: { bold: null } }` removes bold, `{ fill: null }` the fill.
- Equal styles share one entry of the book's `styles`: the session finds an equal style or adds one at the end, and the answer's `styles` is the table then. Styles stay numbers in the book.
- A style by its number (`style: 3`) replaces the cell's whole, as before.
- A style left with nothing is the plain style, which keeps a cell plain in a formatted line.

### Rectangles

Every cell of a rectangle gets the look, the empty ones a cell of their own, as Excel stores a cell it formatted. A rectangle may have as many cells as one column (1,048,576); a larger one is an `InputError`, and whole columns or rows take the look at the cost of the cells they hold.

### Whole rows and columns

Whole lines are formatted as Excel formats them:

- each line's own style takes the fields (its number format the new one), span by span;
- each cell of the lines with a style or number format of its own takes it over its own;
- each cell where the lines cross a line of the other direction with a style (or a number format, for a number format set) gets a cell of its own with what it showed and the patch over it. A row made bold across a filled column gives the cells where they cross bold and filled; a column made italic across a bold row gives them bold and italic.

A cell that shows its lines' look without a look of its own takes none: the line's new style shows through. On whole lines `meta`, `note` and `link` can only be removed (null), from the lines' cells.

## Typing text

`setCellText(sheet, cell, text, opts?)` reads text typed into a cell as `parseEntry` reads it ([Typed entry](entry.md)), in the book's locale and date system, by the session's clock (the year of a date typed without one) and by the number format the cell shows: its own, else its row's, else its column's.

- A format the cell shows from its line and the entry keeps is not written into the cell: `5` typed into a column of percents is 5% and the cell has no format of its own.
- A cell's own General stays its own.
- The cell keeps its style, note, hyperlink and `meta`.
- A number, a date or a formula's number too wide for a column of the default width, or of one Excel fitted, widens the column as Excel widens it (`entryWidth`, [Fitting columns](formats.md#fitting-columns)), in the same step of the undo; `width` in the answer is the width, and `widen: false` gives it without setting it. A hidden column and a cell merged across columns are left as they are.
- Text with a line break wraps the cell (its style, its own or else its row's or column's, with `alignment.wrap`, a style of the table found or added; the answer's `styles` is the new table), and the row of a cell that wraps fits it as AutoFit does, unless its height was set by hand or it is hidden, in the same step ([A line break](entry.md#a-line-break), [Fitting rows](formats.md#fitting-rows)).
- A value the cell's rule of data validation does not take is written, with `validation`: `{ valid: false, rule, severity, message }`, the message the rule's `error` or the library's text in the book's locale. With `refuseInvalid: true` such a value is refused instead (`code: 'validation'`, the cell, the rule and its style) and nothing changes: the host shows Excel's window first and calls again without the option where the person keeps the value (a warning's Yes, an information's OK); for a rule of the style `stop` Excel does not take the value at all. Without the option the host calls `undo()` if the person cancels. A rule that shows no alert (`showError: false`) gives none.
- A formula that does not read is refused with its check code, its position in the text and Excel's message in the book's locale; so are a change of part of an old-kind array (`array`) and a formula in a table's header cell (`header`). Text typed into a header cell renames its column.

`previewCellText` answers the same, the cell's input, its value, the width and the alert, and changes nothing.

`setCellTexts(entries, opts?)` takes several texts at once, as a paste of text or an entry into a selection does, each `{ sheet, cell, text }` read as `setCellText` reads it in its cell:

- The texts are read in order, so a cell given twice keeps the later text, read by the format the earlier one gave it: `15.03.2026` then `45` leave the cell at 45 with the date format.
- All are written with one recomputation, one revision and one step of the undo.
- A column widens once, to the largest width a text typed alone into one of its cells would give it, measured with the values all the texts give: `widths` lists `{ sheet, column, width }`, and `widen: false` gives them without setting them.
- `validations` lists the cells whose values their rules of data validation do not take, each with its alert, as `setCellText`'s `validation`, after all the texts are written.
- A text `setCellText` refuses refuses the whole call with its code, message and position, and its `sheet` and `cell`; nothing is written.

`previewCellTexts` answers each cell's input and value, the widths and the alerts, and changes nothing.

`pasteText(sheet, cell, text, opts?)` pastes text from the clipboard as Excel pastes it, which reads text otherwise than typing in places ([A paste and Control+Return](entry.md#a-paste-and-controlreturn)):

- The text fills a rectangle from the cell, its rows at line breaks and its fields at tabs (`parseClipboard`), a field begun by a quote running to its closing quote, line breaks inside it, and wrapping its cell; an empty field, a short row's missing ones and the row after a last line break clear their cells. The answer's `range` is the rectangle.
- Each field is read as typed into its cell, but an apostrophe stays part of the text, a formula Excel does not take is text, a percent cell has no percent editor and `€5` takes the euro format.
- No column widens and no rule of data validation is checked.
- Part of an old-kind array, a formula into a table's header cell and a rectangle past the sheet's edge (`off_sheet`) are refused with their `sheet` and `cell`; nothing is written.

`fillCellText(sheet, range, text, opts?)` enters a text into every cell of a range, as Excel's Control+Return:

- The text is read once in the active cell, the range's top-left cell or `active`, by its format, and every cell takes what the active cell holds after it: the value, or the formula moved as a copy from the active cell moves it, and the active cell's number format where it has one.
- The columns widen as typing widens them (`widths`, `widen: false` to give them without setting them).
- Only the active cell's rule of data validation is checked: `validation` is its alert, as `setCellText`'s.
- A formula Excel does not take is refused with the active cell, part of an array and a header's formula with theirs; nothing is written.

`previewPasteText` and `previewFillCellText` answer each cell's input and value (and the fill's widths and alert) and change nothing.

`validateEntry(sheet, cell, input)` answers `{ ok: true }` or `{ ok: false, code: 'validation', message, rule, severity }`.

## What an edit answers

Every call that changes the book (`setCells`, `replaceCells`, `clearCells`, `patchCells`, `setCellText`, `setCellTexts`, `pasteText`, `fillCellText`, `setNames`, `setClock`, `setStyles`, `setLayout`, `editStructure`, `undo`, `redo`) answers an `EditResult`:

- `revision`: the session's revision after the call. Each such call adds one, `undo` and `redo` too; a refusal, a failure and an undo with nothing to undo add none.
- `changes`: the cells whose value or shown text changed, and those whose formula, own style or own number format changed, with `formula`, `style` and `numberFormat` then (null when the cell has none now), and a new format also as a code, `numberFormatCode`, a built-in format by the code of the book's locale (14 is `dd.mm.yyyy` in ru-RU), as a window gives it. After a structure operation, the cells whose value or formula text differ from those of the cell that was at their old place: after a row is inserted, `SUM(A1:A4)` that became `SUM(A1:A5)` is listed with its new text.
- `layout`: the sheets whose layout changed (lines, merged areas, tables, rules): a host reads `getLayout()` again.
- `styles`: the book's style table, when the call changed it.
- `moved`: after a structure operation and its undo, how the addresses moved, in order: `{ kind: 'lines', sheet, axis, at, count }` (the lines from `at` on move by `count`, up or left when it is negative), `{ kind: 'range', from, to }`, and a sheet renamed, deleted, inserted or moved. Rows deleted on a filtered sheet are one `lines` move for each run of visible rows, from the bottom up. A host moves its window by them rather than each cell by `mapCell`.
- `repaint`: the rectangles of the book's conditional formats after a call that changed a value or a rule. A rule's look reads values anywhere its formulas point, so a host reads its window again where it meets them.

## Revisions

`session.revision` is the revision now, an asynchronous session's as of its last finished call, and `getViewport` gives the revision its window is of. An edit given `baseRevision` checks it first: when the session's revision is another, the edit changes nothing and throws `CalcError` with the code `stale_revision`. Calls of one session run in order, so a host with one view needs no revision; one that applies edits from several views does. A session restored on a new thread of a pool comes back at its revision.

## Undo

Each call is one step: `clearCells` and `patchCells` whatever their parts (the cells, the lines, the rules, the style table), `setCellText`, `setCellTexts` and `fillCellText` with the widths they gave the columns, `pasteText` whatever its rectangle. A call given nothing to do (`setCells([])`) is no step; one given cells that changed nothing is. The undo's answer is the call's `EditResult` with `call`, and `cells` for the calls that change cells.

## Checked against Excel

Formatting and clearing were measured in Excel 16.113 for Mac and Excel 16.0.20430 for Windows on 34 books (1,967 cells), predicted before Excel ran them; the engine shows what both showed in every cell. Typed text, pastes and Control+Return are checked as [Typed entry](entry.md) says.
