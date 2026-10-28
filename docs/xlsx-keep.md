# Writing over the original file

`writeXlsx` writes a file anew from the workbook, so whatever the workbook cannot hold is gone from it: charts, pictures and shapes, pivot tables, sparklines, the page setup and print area, a protection's password, the theme, formatted runs of text, macros. The `xlcalc/xlsx/keep` subpath writes the workbook over the file it was read from instead. What the workbook changed is written; everything else stays as the file has it, byte for byte. A user who opens a report with a chart, changes one number and saves it gets the report back with its chart.

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
| `locale` | the language of the issues' messages and of a refusal's, as for `writeXlsx` |
| `limits` | how large the original file's parts may be once inflated, as for `readXlsx` |

A change the file cannot take without pointing a part at other cells is refused: `writeKept` throws `KeptRefusal`, exported by the subpath, with the reason in its message and its kind in `code`, and writes no file ([Refusals](#refusals)).

## What is written and what stays

| part | |
|---|---|
| a part the book did not change: charts, drawings, pictures, pivot tables and caches, the theme, `vbaProject.bin`, printer settings, custom XML, document properties | as the file has it, byte for byte (the archive's entry is copied, not compressed again) |
| a sheet the book did not change | as the file has it |
| a sheet the book changed | its rows, cells, columns, merged areas, conditional formats, data validation, hyperlinks, the tables' list, the panes and selection of its first view, written from the workbook, with the conditional formats the workbook does not read among its own ([Conditional formats the workbook does not read](#conditional-formats-the-workbook-does-not-read)); the other elements (page setup, protected ranges, the drawing, sparklines, printing options) stay in their place in the order of the schema, the references an operation of the structure moved written into them |
| a row or a cell the book did not change | the file's own XML: formatted runs of text, rich values, the cell format, a shared formula |
| a cell whose formula did not change but whose value did | the file's own `<f>` with the new value: a shared formula stays shared, and a formula the engine does not parse stays as written |
| a cell, row or column whose look did not change | its own cell format (`s`, `style`), even where another of the file's looks the same to the workbook, and also where rows or columns inserted or deleted moved it ([Structure operations](#structure-operations)): theme colors, a cell style, a quote prefix and the parts of a style the model does not hold stay |
| a new or changed cell | written from the workbook; a new look takes the file's first cell format that looks the same, else a new one; new text goes inline (`t="inlineStr"`) |
| a sheet's protection | as the file has it while the workbook did not change it; changed, the attributes of what users may do are set in the file's element and its password stays; removed with the workbook's protection, written where the schema puts it for a sheet the workbook protects anew |
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
- the criteria and sort of a sheet's AutoFilter, and a sheet's sort, which goes when its cells are deleted ([AutoFilters](#autofilters));
- the names the workbook does not take: the print area and titles (`_xlnm.Print_Area`, `_xlnm.Print_Titles`), a filter's database (`_xlnm._FilterDatabase`) and the other built-in names;
- page breaks, a break going with the row or column it comes before; protected ranges;
- a pivot chart's pivot table, by its sheet's new name;
- the conditional formats the workbook does not read, their ranges and formulas ([Conditional formats the workbook does not read](#conditional-formats-the-workbook-does-not-read));
- the cells with what only the file holds: formatted runs of text, a rich value the writer does not write (a picture in a cell, a data type), phonetic text. Each has a carrier of its own, and moved, it is the file's cell at its new address where the workbook has it as it was. A sheet with more than 5,000 such cells (`KEPT_CELLS`) gives them none: when its rows or columns move, `writeKept` reports the count (`richText`), and the cells that moved are written from the workbook.
- the cell formats the workbook does not tell from another of the file's of the same look: a quote prefix, a cell style, a theme color beside the same color given as RGB, a pattern fill. Each such format has carriers of the blocks of cells it covers (`A6:A9`, `B12:C13`), and of its rows and columns. A cell, row or column that moved keeps its format where its look did not change; one that comes to the address of such a cell does not take it. Rows or columns inserted inside a block stretch it, and the new cells take its format where the workbook gives them its look, as Excel gives them the format of the row above. A sheet whose formats lie in more than 5,000 blocks (`KEPT_FORMATS`) gives them none: when its rows or columns move, `writeKept` reports the count of their cells (`style`), and a cell that moved takes the format of the file's cell at its new address where that looks the same, else the first of its look.

Each sheet has a carrier of its own on its last cell, which tells which sheet of the workbook the file's sheet became, and whether rows or columns were inserted or deleted on it. The parts whose references are not read are known only by it: OLE objects and controls, form controls in a legacy drawing, data consolidation, scenarios, custom views, watched cells, web publishing, smart tags, extensions of other kinds, and macro and dialog sheets (which refer to any sheet). Rows and columns inserted or deleted on their sheet, or its rename or deletion, are refused; a range moved over cells they refer to is not seen.

The carriers' names follow the file's parts in order, so a workbook takes `writeKept` over the file it was opened from, by the same version of `openKept`.

## Links to other workbooks

A link the workbook has after the file's (one a host added with `setExternalBooks`, or in `BookInput.externalBooks` before it opened the session) is written as `writeXlsx` writes it, in a part of its own (`xl/externalLinks/externalLinkN.xml`) with its relationship to the other file, its content type and an `<externalReference>` of the workbook after the file's. A link of the file whose sheets, names or stored values changed (more values from a host, or the mark Excel and a session put on a link a formula is typed over) is written again from the workbook into its part, through the relationship the part has; a link that did not change keeps its part as the file has it. The links keep their numbers, which the file's formulas read them by: a link removed, moved or pointed at another file is refused. Excel 16.114 for Mac opened, without its repair window, each of six files written over files it had made with a link added and a formula typed over it, kept the link and calculated the formula from the stored values; as with any of its workbooks, it saved the link's target as an absolute path and only the values formulas refer to.

## AutoFilters

`openKept` reads a sheet's AutoFilter into `autoFilter` with its range, as `readXlsx` does; its criteria and sort stay in the file. `writeKept` keeps the file's `<autoFilter>`, criteria and sort included, wherever the workbook's AutoFilter is where the file's moved with the operations: rows inserted inside it, say, move both. Where a host put the AutoFilter on another range, it is written from the workbook, without criteria, as a changed table is. Where a host removed it, or an operation took it away, it goes, with its filter database. A sheet that gets an AutoFilter the file did not have gets its filter database too, as `writeXlsx` writes it.

```js
const read = await openKept(file)            // Data's AutoFilter over A1:C21, criteria on Region
const s = createSession(read.book, { cached: read.cached })
s.editStructure({ op: 'insertRows', sheet: 'Data', row: 5 })
await writeKept(file, s.getSnapshot())       // <autoFilter ref="A1:C22"> with the criteria on Region
s.setLayout([{ sheet: 'Data', autoFilter: { range: 'A1:D22' } }])
await writeKept(file, s.getSnapshot())       // <autoFilter ref="A1:D22">, the file's criteria lost: the sheet, filtered, gets the writer's
```

## Conditional formats the workbook does not read

A file can hold conditional formats the workbook cannot: a color scale with an automatic color, a rule from another program, a rule whose fill is a gradient. `readXlsx` reports them (`conditionalFormat`). When a host adds, deletes or changes a rule on such a sheet, or rows and columns move under them, `writeKept` writes the workbook's rules and puts the file's other rules back among them as the file has them, Excel 2010's rules in the sheet's extension list too. Each keeps its place after the rule of the workbook it came after in the file; the priorities are counted from 1 over all of them.

```js
// S has, in priority order: a rule on A2:A21, a color scale with an automatic color on B2:B21 (not read), a rule on C2:C21
const read = await openKept(file)
const s = createSession(read.book, { cached: read.cached })
const [a, c] = s.getBook().sheets[0].conditionalFormats
s.setLayout([{ sheet: 'S', conditionalFormats: [c] }])
await writeKept(file, s.getSnapshot())   // priority 1: the color scale on B2:B21, 2: the rule on C2:C21
s.editStructure({ op: 'insertRows', sheet: 'S', row: 1 })
await writeKept(file, s.getSnapshot())   // the color scale on B3:B22, its threshold $G$2 now $G$3
```

Such a rule moves with rows and columns as Excel moves a rule:

- Rows or columns inserted inside its ranges, or right after them, join them; a rule after them moves. A range that ends on the sheet's last row or column keeps ending there.
- Deleted rows and columns leave its ranges, and a rule with no cells left goes. A rule goes after the others as Excel puts it back, by the rules of the workbook's ([Conditional formats](conditional-formats.md#structure-operations)): a color scale whose threshold `$G$2` became `$E$2` after columns C and D were deleted goes, though its cells stayed.
- Its formulas and thresholds move as a cell's formula at the rule's first cell; a reference whose cells were deleted becomes `#REF!`, and one to a sheet renamed or deleted follows it. A move of all its cells on the sheet takes it along.

A rule the workbook reads without a part of its look (a gradient fill, a fill that removes the cell's own) keeps the file's differential format while the workbook's look of it stays the same; a new look is the workbook's.

What cannot go as Excel would is refused: a pivot table's rules, which the pivot table names by priority, a rule the workbook read over part of its ranges, and a rule whose formulas have references relative to its first cell when a deletion took that cell (`rules`); a rule whose cells were moved to another sheet (`moved`). A rule whose ranges or formulas do not parse is known only by its sheet's carrier: rows and columns inserted or deleted there are refused (`unread`).

## Refusals

`writeKept` refuses, with `KeptRefusal`, a change it cannot write as Excel would, or that would leave the file pointing at other cells than the workbook does or lose what it holds. Its message is in the language of `writeKept`'s `locale`, as its issues are (`{ locale: 'ru-RU' }` once ru-RU is loaded, or the locale object), English by default; its `code` (`KeptRefusalCode`) gives the kind for a program, in brackets:

- a range moved over a drawing (its cells move and its lines do not), or a change that would cut a pivot table, which Excel refuses (`moved`); a part that moves in a way `writeKept` does not write (`moved`);
- a drawing whose cells were deleted with all the rows below or the columns beside it; the source of a pivot cache on a sheet deleted (Excel links it to the file itself by its path, which `writeKept` does not know), or in cells deleted; a sparkline's cell deleted (`deleted`);
- a sheet deleted that a pivot chart names, or a table or a name deleted that a pivot cache takes its data from; a sheet or a table deleted with parts of its own the workbook does not have (a pivot table, a slicer) (`deleted`);
- the cells of comment threads moved, or notes changed on a sheet with comment threads, which the workbook reads as notes (`threads`);
- rows or columns inserted or deleted on a sheet with parts whose references are not read, or such a sheet renamed or deleted (`unread`);
- the data validation of a sheet changed, by an edit or by rows and columns moving, where the file has rules of that sheet the workbook could not read (those `readXlsx` reports as `dataValidation` issues): written again from the workbook, they would be lost; the conditional formats of a sheet changed where the file has a rule of it that cannot go with them as Excel would ([Conditional formats the workbook does not read](#conditional-formats-the-workbook-does-not-read)) (`rules`);
- a link to another workbook removed, moved among the links or pointed at another file, which the file's formulas would read by another number, or a link of the file of another kind (DDE, OLE) changed (`links`);
- the workbook's default font (`font`) changed, which every cell without a font of its own would show (`font`);
- a carrier missing: a workbook not opened by `openKept`, or whose hidden names were dropped (`carriers`);
- a file that names a sheet as Excel would not, which the reader read under another name, or a Strict Open XML file, once anything changed; a worksheet in other namespaces (`file`).

Everything the workbook holds moves with these operations on its own: formulas, names, conditional formats, data validation, merged areas, tables, AutoFilters, notes, hyperlinks, frozen panes.

## Known differences

- A series' category or name whose cells were all deleted keeps `#REF!` in its place; Excel moves it to an extension of the series (`c15:filteredCategoryTitle`, `c15:filteredSeriesTitle`). Both read as lost.
- A drawing that moves without sizing keeps its size by the rows' heights and the columns' widths the workbook has; Excel's answers were taken for rows of the default height. A drawing goes with all its columns deleted by the rule Excel showed for rows.
- The formats and strings no cell uses any more stay in `styles.xml` and `sharedStrings.xml`; Excel drops them when it saves.
- Excel gives a cell, row or column it inserts the format of the line before it. `writeKept` gives a row or column inserted right after a block of a format, and its cells, the first format of their look; a new cell of the plain look inside a block of a quote prefix, a cell style or a pattern fill is not written, as the workbook has no cell there.
- A cell that looks as its row or column, which the workbook does not hold as a cell, is left out where rows or columns moved; Excel keeps it at its new address. It looks the same either way.
- A cell moved by a range takes the format of the file's cell at its new address where that looks the same, else the first of its look. So do the cells of a sheet whose carriers of formats a host dropped.
- A cell that moved is the file's own only on its sheet: one moved by a range to another sheet is written from the workbook.
- A table the workbook changed is written from the workbook: the criteria and sort of its AutoFilter are lost, as when `writeXlsx` writes it. So is a sheet's AutoFilter a host put on another range, as when `writeXlsx` writes it, and one removed goes with its filter database (`_xlnm._FilterDatabase`).
- The filter database goes with the workbook's AutoFilter where that moves. Excel moves it as a name: it keeps it over cells moved to another sheet, where the AutoFilter goes. Where the AutoFilter stays, `writeKept` moves it as Excel does, `#REF!` under cells moved over the AutoFilter, stretched over its last rows moved down.
- When a sheet's data validation changes, its rules are written from the workbook; a sheet with rules the workbook could not read is refused (`rules`) rather than written without them.
- The workbook's list of rules cannot say on which side of a rule it does not read a moved rule goes: that rule stays after the rule it came after. Excel, given a rule's `Priority`, puts it before the rule that had it, a carried one too.
- Excel keeps the priorities rules had after it deletes or moves rules, gaps and all (1, 3, 4, 10); `writeKept` counts them from 1, in the same order.
- A move over part of a carried rule's cells, a paste over them and a clearing of their formats are not seen: the rule keeps its cells, where Excel takes the moved ones along, or leaves the pasted and cleared ones.
- A carried rule that loses its last rows or columns while its formulas change goes after the others, where Excel keeps its place: no carrier tells which of its lines were deleted. After rows or columns both inserted and deleted on its sheet, which the carriers do not tell apart, a carried rule goes after the others only when it is over the statistics of its cells and lost some.
- A row the workbook did not change keeps the collapsed mark the file has, where `writeXlsx` writes the mark as Excel keeps it.
- The ignored errors of a sheet with rows or columns inserted or deleted are left out: no carrier holds their cells.
- `docProps/app.xml`, which lists the sheets' names for file browsers, stays as it was.

## Checked against Excel

Excel 16.113 for Mac opened, without its repair window, each of 30 files written over files it had made with charts, a chart sheet, pictures, pivot tables, formatted runs, theme colours, the print area, a page break and notes, after edits of values, formulas, styles, merged areas, rules, validation, hyperlinks, notes, and rows and sheets where no part refers; it saved every one of those parts again and the values written. For 20 operations of the structure over such files (rows and columns inserted and deleted in and around a chart's data, under pictures of each `editAs`, around a pivot table and its source, the print area and a page break; a range moved; sheets renamed and deleted), Excel 16.114 did each operation itself and saved the file, and `writeKept` writes the references Excel saved, but for a lost category or series name, which Excel moves to an extension, and the one change it refuses (a pivot cache's sheet deleted). Excel opened each file `writeKept` wrote for them without its repair window and kept its parts and references. Without edits, each file of the test sets and files of real workbooks from other sources is written back as its own bytes; after random edits and operations of the structure, a file reads back as `writeXlsx` writes the same workbook, with the parts no edit concerns unchanged, or the change is refused. Excel 16.114 opened, without its repair window, files `writeKept` wrote over protected files with a password, after a value changed, the protection changed or removed, rows inserted above an unlocked cell and a sheet protected anew, and read and saved the protection the workbook holds, the password kept ([Protection](xlsx.md#protection)). For 19 operations over a file with an AutoFilter, Excel 16.114 moved or removed the AutoFilter as the workbook does, and opened each file `writeKept` wrote for them without its repair window, with the AutoFilter where the workbook has it; so for 14 files written after edits that put, moved or removed an AutoFilter, or kept one with its criteria. For 13 operations of rows and columns over a file Excel made with a quote prefix, a theme color beside its RGB, a cell style, a pattern fill and rows and columns of each color, Excel 16.114 gave every cell, row and column that moved the cell format `writeKept` gives it, by the format's content, opened each file `writeKept` wrote without its repair window and kept the formats. Excel for Windows was not run for these. For the conditional formats the workbook does not read, Excel 16.0.20430 for Windows opened and saved books with a color scale with an automatic color, Excel 2010's color scale over another sheet and a rule with a gradient fill among rules the workbook reads, and did 35 cases itself over one of them (rules added, deleted, given another priority or look; rows and columns inserted and deleted; a sheet renamed and deleted; ranges moved): the rules it saved are in the order `writeKept` writes, with the same ranges and formulas, in 31; of the other four, `writeKept` refuses one (a rule moved to another sheet) and three are the differences above. Excel 16.114 for Mac opened each of the 34 files `writeKept` wrote for them without its repair window and kept their rules. [docs/development.md](development.md#writing-over-the-original-file) has the details.
