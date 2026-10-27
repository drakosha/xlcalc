# Writing over the original file

`writeXlsx` writes a file anew from the workbook, so whatever the workbook cannot hold is gone from it: charts, pictures and shapes, pivot tables, sparklines, the page setup and print area, protection, the theme, formatted runs of text, macros. The `xlcalc/xlsx/keep` subpath writes the workbook over the file it was read from instead. What the workbook changed is written; everything else stays as the file has it, byte for byte. A user who opens a report with a chart, changes one number and saves it gets the report back with its chart.

```js
import { compute, createSession } from 'xlcalc'
import { readXlsx, writeXlsx } from 'xlcalc/xlsx'
import { openKept, writeKept } from 'xlcalc/xlsx/keep'

// a file from elsewhere, here one the writer makes: `await readFile(path)` gives the bytes of a real one
const made = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*10' } } }] }
const original = (await writeXlsx({ book: made, result: compute(made) })).data
const { book, cached } = await openKept(original)
const session = createSession(book, { cached })
session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
const { data } = await writeKept(original, session.getSnapshot())
console.log((await readXlsx(data)).cached.sheets[0].cells.A2) // → { value: 30 }

// nothing changed: the original bytes
const same = await writeKept(original, { book, result: cached })
console.log(same.data === original) // → true
```

## In short

`openKept(data, opts)` reads a file as `readXlsx` does and returns the same `{ book, cached, issues }`, with hidden names added to the workbook, the carriers ([Structure operations](#structure-operations)). `writeKept(original, { book, result }, opts)` takes the original file's bytes and a snapshot, as `writeXlsx` takes it (`session.getSnapshot()`), and resolves to `{ data, issues }`. It is a pure function of the two: nothing is kept between opening and writing, so a host stores the original file of a version and writes each new version over it.

| option | |
|---|---|
| `values` | which formula values the file stores, as for `writeXlsx`: `'require'` by default with a result, `'omit'` without one, `'preserve'` |
| `locale` | the language of the issues' messages, as for `writeXlsx` |
| `limits` | how large the original file's parts may be once inflated, as for `readXlsx` |

A change the file cannot take without pointing a part at other cells is refused: `writeKept` throws `KeptRefusal`, exported by the subpath, with the reason in its message and its kind in `code`, and writes no file ([Refusals](#refusals)).

## What is written and what stays

| part | |
|---|---|
| a part the book did not change: charts, drawings, pictures, pivot tables and caches, the theme, `vbaProject.bin`, printer settings, custom XML, document properties | as the file has it, byte for byte (the archive's entry is copied, not compressed again) |
| a sheet the book did not change | as the file has it |
| a sheet the book changed | its rows, cells, columns, merged areas, conditional formats, data validation, hyperlinks, the tables' list, the panes and selection of its first view, written from the workbook; the other elements (page setup, protection, the drawing, sparklines, printing options) stay in their place in the order of the schema, the references an operation of the structure moved written into them |
| a row or a cell the book did not change | the file's own XML: formatted runs of text, rich values, the cell format, a shared formula |
| a cell whose formula did not change but whose value did | the file's own `<f>` with the new value: a shared formula stays shared, and a formula the engine does not parse stays as written |
| a new or changed cell | written from the workbook; a cell format that looks as one of the file's takes its index, so its theme colors and the parts of a style the model does not hold stay; new text goes inline (`t="inlineStr"`) |
| `styles.xml` | added to: new cell formats, fonts, fills, borders, number formats and differential formats at the end of their lists, the rest as it was |
| `sharedStrings.xml` | as it was |
| `calcChain.xml` | left out when anything changed: Excel builds it again |
| `workbook.xml` | the sheets' names, order and states, the defined names, the active sheet and iterative calculation, where the book changed them; a full calculation on loading when cells changed, as `writeXlsx` asks for one |
| a table the book changed or added | written from the workbook in its part, the file's id kept |
| notes the book changed | the comments part and its VML drawing written again |
| a sheet the book added | written as `writeXlsx` writes one |
| a sheet the book deleted | removed with its own parts (tables, notes, printer settings, drawings with the charts and pictures only they show) |

A workbook that did not change, the values of its formulas aside, gives the original bytes: the values Excel stored stay. Once something changed, the formulas take their values from the result, as `writeXlsx` writes them, and Excel computes the workbook again when it opens the file.

What the file holds and the workbook does not is reported by `openKept`'s issues as by `readXlsx`'s: the workbook cannot hold it, but the file keeps it.

## Structure operations

Inserting or deleting rows and columns, moving a range, and renaming, moving or deleting a sheet move references in the parts the workbook does not hold: the series of a chart, the anchor of a picture, the source of a pivot table. `openKept` gives each such reference to the workbook as a hidden name with the reserved prefix `_xlcalc.keep.`, the carrier, and the engine moves these names as it moves every name, by Excel's rules. `writeKept` reads them back, writes each moved reference into its part, and does not write the carriers to the file; `writeXlsx` leaves them out too, so a workbook opened by `openKept` can also be saved as a new file. A host keeps them among the workbook's names (`setNames` replaces the whole list) and leaves hidden names out of its lists of names, as Excel's Name Manager does.

What follows the cells this way:

- the series, categories and labels of charts (every `<c:f>`, also of the charts of chart sheets), with the `$`s the file gave them; a reference whose cells were deleted becomes `Sheet!#REF!`, one whose sheet was deleted `#REF!`;
- the anchors of drawings (charts, pictures, shapes, slicers) by how each follows the cells. One that moves and sizes with them (`editAs="twoCell"`, the default) follows by both corners; a corner whose row or column was deleted lands where the deleted lines were, its offset 0; with all its rows or all its columns deleted the drawing is deleted, with the chart or picture only it shows, as Excel deletes it. One that moves without sizing (`oneCell`, and a one-cell anchor) follows by its top left corner and keeps its size, its bottom right corner counted again from the rows' heights and the columns' widths. One that is fixed (`absolute`) stays;
- the place of a pivot table, and the source of its cache: cells, or a table or a name, which takes a table's new name;
- sparklines, their data and their cells;
- a sheet's AutoFilter and its sort, whose sort goes when its cells are deleted;
- the names the workbook does not take: the print area and titles (`_xlnm.Print_Area`, `_xlnm.Print_Titles`), a filter's database (`_xlnm._FilterDatabase`) and the other built-in names;
- page breaks, a break going with the row or column it comes before; protected ranges;
- a pivot chart's pivot table, by its sheet's new name;
- the cells with what only the file holds: formatted runs of text, a rich value the writer does not write (a picture in a cell, a data type), phonetic text. Each has a carrier of its own, and moved, it is the file's cell at its new address where the workbook has it as it was. A sheet with more than 5,000 such cells (`KEPT_CELLS`) gives them none: when its rows or columns move, `writeKept` reports the count (`richText`), and the cells that moved are written from the workbook.

Each sheet has a carrier of its own on its last cell, which tells which sheet of the workbook the file's sheet became, and whether rows or columns were inserted or deleted on it. The parts whose references are not read are known only by it: OLE objects and controls, form controls in a legacy drawing, data consolidation, scenarios, custom views, watched cells, web publishing, smart tags, extensions of other kinds, and macro and dialog sheets (which refer to any sheet). Rows and columns inserted or deleted on their sheet, or its rename or deletion, are refused; a range moved over cells they refer to is not seen.

The carriers' names follow the file's parts in order, so a workbook takes `writeKept` over the file it was opened from, by the same version of `openKept`.

## Refusals

`writeKept` refuses, with `KeptRefusal`, a change it cannot write as Excel would, or that would leave the file pointing at other cells than the workbook does or lose what it holds. Its `code` (`KeptRefusalCode`) gives the kind for a program, in brackets:

- a range moved over a drawing (its cells move and its lines do not), or a change that would cut a pivot table, which Excel refuses (`moved`); a part that moves in a way `writeKept` does not write (`moved`);
- a drawing whose cells were deleted with all the rows below or the columns beside it; the source of a pivot cache on a sheet deleted (Excel links it to the file itself by its path, which `writeKept` does not know), or in cells deleted; a sparkline's cell deleted, an AutoFilter whose range was deleted (`deleted`);
- a sheet deleted that a pivot chart names, or a table or a name deleted that a pivot cache takes its data from; a sheet or a table deleted with parts of its own the workbook does not have (a pivot table, a slicer) (`deleted`);
- the cells of comment threads moved, or notes changed on a sheet with comment threads, which the workbook reads as notes (`threads`);
- rows or columns inserted or deleted on a sheet with parts whose references are not read, or such a sheet renamed or deleted (`unread`);
- the conditional formats or the data validation of a sheet changed, by an edit or by rows and columns moving, where the file has rules of that sheet the workbook could not read (those `readXlsx` reports as `conditionalFormat` and `dataValidation` issues): written again from the workbook, they would be lost (`rules`);
- the links to other workbooks changed (`links`);
- the workbook's default font (`font`) changed, which every cell without a font of its own would show (`font`);
- a carrier missing: a workbook not opened by `openKept`, or whose hidden names were dropped (`carriers`);
- a file that names a sheet as Excel would not, which the reader read under another name, or a Strict Open XML file, once anything changed; a worksheet in other namespaces (`file`).

Everything the workbook holds moves with these operations on its own: formulas, names, conditional formats, data validation, merged areas, tables, notes, hyperlinks, frozen panes.

## Known differences

- A series' category or name whose cells were all deleted keeps `#REF!` in its place; Excel moves it to an extension of the series (`c15:filteredCategoryTitle`, `c15:filteredSeriesTitle`). Both read as lost.
- A drawing that moves without sizing keeps its size by the rows' heights and the columns' widths the workbook has; Excel's answers were taken for rows of the default height. A drawing goes with all its columns deleted by the rule Excel showed for rows.
- The formats and strings no cell uses any more stay in `styles.xml` and `sharedStrings.xml`; Excel drops them when it saves.
- A cell that moved is the file's own only on its sheet: one moved by a range to another sheet is written from the workbook.
- A table the workbook changed is written from the workbook: the criteria and sort of its AutoFilter are lost, as when `writeXlsx` writes it.
- When a sheet's conditional formats change, its rules are written from the workbook; a sheet with rules the workbook could not read is refused (`rules`) rather than written without them. So with data validation.
- A row the workbook did not change keeps the collapsed mark the file has, where `writeXlsx` writes the mark as Excel keeps it.
- The ignored errors of a sheet with rows or columns inserted or deleted are left out: no carrier holds their cells.
- `docProps/app.xml`, which lists the sheets' names for file browsers, stays as it was.

## Checked against Excel

Excel 16.113 for Mac opened, without its repair window, each of 30 files written over files it had made with charts, a chart sheet, pictures, pivot tables, formatted runs, theme colours, the print area, a page break and notes, after edits of values, formulas, styles, merged areas, rules, validation, hyperlinks, notes, and rows and sheets where no part refers; it saved every one of those parts again and the values written. For 20 operations of the structure over such files (rows and columns inserted and deleted in and around a chart's data, under pictures of each `editAs`, around a pivot table and its source, the print area and a page break; a range moved; sheets renamed and deleted), Excel 16.114 did each operation itself and saved the file, and `writeKept` writes the references Excel saved, but for a lost category or series name, which Excel moves to an extension, and the one change it refuses (a pivot cache's sheet deleted). Excel opened each file `writeKept` wrote for them without its repair window and kept its parts and references. Without edits, each file of the test sets and files of real workbooks from other sources is written back as its own bytes; after random edits and operations of the structure, a file reads back as `writeXlsx` writes the same workbook, with the parts no edit concerns unchanged, or the change is refused. Excel for Windows was not run. [docs/development.md](development.md#writing-over-the-original-file) has the details.
