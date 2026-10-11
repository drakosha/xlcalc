# Editing a grid

A grid over a book edits cells the way Excel does: a value typed over a cell keeps the cell's look, Delete clears the content, Clear Formats the look, Bold over a selection goes over what each cell shows, and typed text is read in the book's language. A session does each of these in one call, recomputes what depends on it, keeps it as one step of the undo and answers what the grid has to draw again, so a host needs no code of its own between what a person does and the book.

The rules of formatting and clearing below were measured in Excel 16.113 for Mac and Excel 16.0.20430 for Windows, en-US.

## Editing a cell

`setCells(edits)` sets the fields an edit gives and keeps the cell's others: a new value keeps the style, the number format, the note, the hyperlink and `meta`.

```js
session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 5 } }])
// A1 was { value: 2, style: 0, format: '0.00', note: { text: 'n' } }, now { value: 5, style: 0, format: '0.00', note: { text: 'n' } }
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

`patchCells(patches)` sets the look of a `cell` or a `range`: a rectangle, whole columns or whole rows. A patch sets `style`, `numberFormat` (or its old name `format`), `meta`, `note` and `link`; null removes a field; values and formulas stay.

### A style by its fields

A style given by its fields goes over the style each cell shows, as Excel formats a selection: the cell's own style, else its row's, else its column's.

```js
session.patchCells([{ sheet: 'S', range: 'A1:C3', set: { style: { font: { bold: true } } } }])
// A1 italic: bold and italic; B2 empty: bold; C1 in a filled column: bold and filled
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
- A value the cell's rule of data validation does not take is written, with `validation`: `{ valid: false, rule, severity, message }`, the message the rule's `error` or the library's text in the book's locale. The host shows Excel's window and calls `undo()` if the person cancels; for a rule of the style `stop` Excel does not take the value, which the host refuses itself. A rule that shows no alert (`showError: false`) gives none.
- A formula that does not read is refused with its check code, its position in the text and Excel's message in the book's locale; so are a change of part of an old-kind array (`array`) and a formula in a table's header cell (`header`). Text typed into a header cell renames its column.

`previewCellText` answers the same, the cell's input, its value, the width and the alert, and changes nothing.

`validateEntry(sheet, cell, input)` answers `{ ok: true }` or `{ ok: false, code: 'validation', message, rule, severity }`.

## What an edit answers

Every call that changes the book (`setCells`, `replaceCells`, `clearCells`, `patchCells`, `setCellText`, `setNames`, `setClock`, `setStyles`, `setLayout`, `restructure`, `undo`, `redo`) answers an `EditResult`:

- `revision`: the session's revision after the call. Each such call adds one, `undo` and `redo` too; a refusal, a failure and an undo with nothing to undo add none.
- `changes`: the cells whose value or shown text changed, and those whose formula, own style or own number format changed, with `formula`, `style` and `numberFormat` then (null when the cell has none now). After a structure operation, the cells whose value or formula text differ from those of the cell that was at their old place: after a row is inserted, `SUM(A1:A4)` that became `SUM(A1:A5)` is listed with its new text.
- `layout`: the sheets whose layout changed (lines, merged areas, tables, rules): a host reads `layout()` again.
- `styles`: the book's style table, when the call changed it.
- `moved`: after a structure operation and its undo, how the addresses moved, in order: `{ kind: 'lines', sheet, axis, at, count }` (the lines from `at` on move by `count`, up or left when it is negative), `{ kind: 'range', from, to }`, and a sheet renamed, deleted, inserted or moved. A host moves its window by them rather than each cell by `mapCell`.
- `repaint`: the rectangles of the book's conditional formats after a call that changed a value or a rule. A rule's look reads values anywhere its formulas point, so a host reads its window again where it meets them.

## Revisions

`session.revision` is the revision now, an asynchronous session's as of its last finished call, and `getViewport` gives the revision its window is of. An edit given `baseRevision` checks it first: when the session's revision is another, the edit changes nothing and throws `CalcError` with the code `stale_revision`. Calls of one session run in order, so a host with one view needs no revision; one that applies edits from several views does. A session restored on a new thread of a pool comes back at its revision.

## Undo

Each call is one step: `clearCells` and `patchCells` whatever their parts (the cells, the lines, the rules, the style table), `setCellText` with the width it gave the column. A call given nothing to do (`setCells([])`) is no step; one given cells that changed nothing is. The undo's answer is the call's `EditResult` with `call`, and `cells` for the calls that change cells.

## How it was checked

The lines harness (`bench/coverage/lines.ts`, `npx tsx bench/coverage/run.ts lines <set>`, and `bench/windows/run.ts lines <set>` for Excel for Windows) writes a book with `writeXlsx`; Excel takes the case's steps through its object model (a number format, bold, italic or a fill set on a range, a value typed into a cell, ClearContents, ClearFormats or Clear of a range), reads what each listed cell shows (`Range.Text`, the number format, the display format's fill and font) and saves the book. `test/lines-excel.test.ts` takes the same steps in a session (`patchCells`, `setCells`, `clearCells`) and compares every cell; for the clearing cases it also compares the notes, hyperlinks, conditional formats and rules of data validation of the file Excel saved, read back by `readXlsx`, with the session's book.

| set | cases | cells | what |
|---|---|---|---|
| `files`, `held` | 15 | 1,087 | measured before for the writer and the reader: columns and rows formatted across each other, values typed into them |
| `patch` | 7 | 124 | a rectangle over a filled column and across an italic row; a column over cells with their own style and format; a row's date format across a column of percents; a column filled, then bold, then a row italic across it; a row italic over a filled cell; a column filled across a row with a number format alone; a style over a book's own line styles |
| `clear` | 12 | 756 | ClearContents, ClearFormats and Clear of a rectangle of values, formulas, styles, formats, notes and hyperlinks at the top of a rule's cells; of cells of a formatted column; of a whole column crossing a bold row and of a whole row; a value typed into cells cleared of their formats |

The engine's answers to `patch` and `clear` were committed before Excel for Mac ran them. Every cell of both sets showed what was predicted, and so did the notes, conditional formats and rules of data validation of the files; the hyperlinks did not: Excel's ClearContents takes a cell's hyperlinks, which the prediction kept, and the session now takes them. Excel for Windows, run after, showed what Excel for Mac showed in every cell and every part of the files. On the cases of `files` and `held` Excel showed what a session's `patchCells` gives in every cell, the cells where Excel formats a line across a formatted one, which `setLayout` leaves to the line's style, among them.

The seeded test of sessions against full computations (`test/session-fuzz.test.ts`) replaces, clears, formats and types into cells beside its other edits, compares the session's hash with a full computation's after each, and at the end undoes every step and finds the book it opened. `test/session-edit.test.ts`, `test/session-styles.test.ts` and `test/session-text.test.ts` check each call in a synchronous session, in the same thread and in a thread, with closures and eval.
