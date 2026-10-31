# PivotTables

A workbook read from a file holds its PivotTables: `readXlsx` and `openKept` read each PivotTable with its cache, its
fields on their axes, its value fields, filters, sorting, totals, layout, formats and the layout Excel last saved, and the
workbook's slicers and timelines. `GETPIVOTDATA` calculates over that saved layout, so a report that reads its numbers
out of a PivotTable gets Excel's values without the PivotTable being refreshed.

```js
const { book } = await readXlsx(data)
book.pivotCaches[0].fields.map((f) => f.name)
// ['Region', 'Product', 'Price', …]
book.sheets[1].pivotTables[0].values
// [{ field: 2, name: 'Sum of Price' }]
compute({ ...book, sheets: [...book.sheets, { name: 'G', cells: { A1: { formula: 'GETPIVOTDATA("Sum of Price",P!A3,"Region","North")' } } }] })
// G!A1: North's sum, the value of its cell in the PivotTable
```

The engine does not refresh a PivotTable or lay it out again yet: its cells hold the values the file saved, as
values of the sheet.

## The model

The caches are the workbook's, `BookInput.pivotCaches`, each by its `id`; the PivotTables are their sheets',
`SheetInput.pivotTables`, each naming its cache by that id. Slicer and timeline caches are the workbook's
(`slicerCaches`, `timelineCaches`), slicers and timelines their sheets' (`slicers`, `timelines`); a cache names the
PivotTables it filters by sheet and name. Everything is plain data, so a workbook with PivotTables goes to a pool's
threads as any other ([API](api.md)).

- `PivotCache`: the source (cells of a sheet, a table, a defined name, in this workbook or another; consolidated
  ranges, an external connection, the data model and scenarios are held as the file stores them), the fields with their
  items, the kinds of value each holds, grouping (dates by levels, numbers by intervals, groups made by hand) and
  calculated fields' formulas, the rows the cache saved (`records`), calculated items, `refreshOnLoad` and `saveData`.
  A value of a cache is a blank (`null`), a number, text, a boolean, an error or a date (`{ date: '2024-01-01T00:00:00' }`).
- `PivotTable`: its place (`location`: the range without the Filters area, where its headers and values begin), one
  `PivotField` per cache field (items with what is hidden and collapsed, subtotals, form, sorting, repeated labels, a
  blank line after each item), the fields on rows, columns and the Filters area, the value fields (`PivotDataField`:
  the function, Show Values As with its base field and item, the number format, the name), filters by label, value,
  Top 10 and date, grand totals, the form, captions, what errors and empty cells show, the style, the formats of its
  areas, and the saved layout: `rowLines` and `columnLines`, each line of the rows and of the columns as Excel laid it out.
- A PivotTable's conditional format stays among its sheet's rules, with `pivot` naming the PivotTable and the areas
  it applies to.

`rows` and `columns` hold `values` where several value fields lie along that axis (Σ Values). Fields and items are
named by index: a field by its place in the cache, an item by its place in its field (`PivotFieldItem.item` into the
cache field's items).

The engine checks what it reads of a PivotTable (its place, fields, value fields and saved lines, its cache's field
names and items) and what refers to what (a PivotTable's cache, a slicer's cache and PivotTables, unique ids and names);
the rest of the model's shape is checked by `readXlsx` and `writeXlsx`.

## What the file holds that the model does not

`readXlsx` reports in `issues` (`pivotTable`):

- a cache over an external connection, the data model, consolidated ranges or scenarios: read as the file stores it,
  never refreshed;
- the OLAP part of a cache (hierarchies, measures, dimensions);
- a PivotTable's chart formats, and its PivotCharts, which are charts of the sheet's drawing;
- a PivotTable or a timeline whose cache is missing from the file, which is left out;
- a PivotTable's conditional format that no PivotTable of its sheet names, which is left out.

`writeXlsx` does not write PivotTables, caches, slicers and timelines yet: the file holds the values of their cells, and
a PivotTable's conditional formats are left out; it reports so (`pivotTable`). Writing over the original file with
`writeKept` keeps every part of them as the file has it ([Writing over the original file](xlsx-keep.md)).

A sheet renamed in a session renames it in the sources of the caches and in the PivotTables slicers and timelines name.

## GETPIVOTDATA

`GETPIVOTDATA(data_field, pivot_table, [field1, item1], …)` returns the value a PivotTable shows for a value field at
the items the pairs name:

- `pivot_table` is a reference to any cell of the PivotTable, its Filters area too, or a range whose top left cell is
  one; without a Filters area, the cell just above its top left cell is the PivotTable's too (`A2` of a PivotTable at
  `A3`, not `B2` or `A1`). A cell outside every PivotTable is `#REF!`.
- `data_field` is the value field's name (`Sum of Price`) or its source field's (`Price`), in any case. Given as a
  reference to a cell holding the name, it is `#REF!`, as in Excel; text a formula makes (`TEXT(A1,"0")`) is taken.
- A field is named by its name in the PivotTable, the source's where it has none of its own (a field renamed `Area` is
  not found as `Region`), in any case. An item is matched by its label, in any case:
  `"North"`, `"(blank)"` for the blank item, a number's or a date's label as the PivotTable shows it (`"10"`, not
  `"10.0"`; `"1/1/2024"`, not `"2024-01-01"`), a text item before a number of the same label. A number matches a number
  or a date of that value as `=` compares them (`DATE(2024,1,1)`, `45292`, `DATE(2025,1,2)+TIME(9,33,0)`), a boolean a
  boolean.
- The pairs pick the line of the rows and the column that name exactly their fields: naming all of them gives a value
  cell, naming fewer a subtotal or the grand total where the PivotTable shows one, a collapsed item's total on its own
  line. An inner field named without a field between finds the one line that names it, where there is one (North and
  Code 10 when only one Flag of North has Code 10). A field of the Filters area is matched against the item it shows.
  What the layout does not show (an item filtered out, an inner field under several outer items, a field on no axis, a
  pair without its item, an item with subtotals of several functions) is `#REF!`.
- The answer is the value of that cell, an error as it is, an empty cell 0.
- A PivotTable whose definition changed without being laid out again (`PivotCache.stale`) is `#REF!`.
- A PivotTable in another workbook (`[1]P!$A$3`) is `#REF!` when calculated; a file opened keeps the value Excel saved
  for such a formula, as it keeps one over the functions of connections ([Values a file stores](formulas.md#values-a-file-stores)).

The formula learns the PivotTable's cells as its references, so an edit of them recalculates it; a change of the
PivotTable builds the workbook again. Its call takes up to 254 arguments (126 pairs).

## Layout of a PivotTable

When `xlcalc/pivot` lays a PivotTable out (on a refresh, or after its definition changed), its cells come out as Excel lays them: the same lines of rows and columns, the same place on the sheet and the same values in the same cells. The workbook keeps the layout with the PivotTable (`rowLines`, `columnLines`, `location`), so that a file written afterwards opens in Excel without a refresh and `GETPIVOTDATA` finds its cells.

The cells of the area hold values, not formulas, and change only when the PivotTable is laid out again: a change of the source data shows after a refresh, as in Excel.

### Forms

- **Compact** (the default): the row fields share the first column, each level indented one step further; the header of that column is `Row Labels`, of the columns `Column Labels`.
- **Outline**: each row field has a column of its own, headed by the field's name; an item's label stands on a line of its own above the items under it.
- **Tabular**: as outline, but an item's label stands on the first line of its group, and its subtotal on a line after the group.

A field can have a form of its own (`PivotField.layout`). A compact field shares the column of the field before it only when that field is compact too.

### Labels

An item shows its caption (`PivotFieldItem.caption`), or its value: text as text, a number as a number with its field's number format, a date as a date (the short date format without one of its own), `TRUE` and an error such as `#N/A` as text, an empty value as `(blank)`. With **Repeat Item Labels** (`repeatLabels`) an outline or tabular field repeats its label on every line of its group, the subtotal lines of inner fields included; a compact field never repeats.

### Subtotals and grand totals

- An item over a group shows its subtotal at the top, on its own line (`subtotalTop`, the default, in compact and outline forms), or on a line below the group, labelled `East Total`.
- Subtotals of several functions (`subtotals: ['sum', 'count']`) always go below, one line each: `East Sum`, `East Count`. A field's one function of its own goes at the top, as the automatic one does.
- A subtotal caption of the field's own stands `?` for the item's label: `? all` shows `East all`.
- A field has subtotals only when a field other than the value fields lies under it.
- The grand total of the rows is a line at the bottom (`grandTotals.rows`), of the columns a column at the right (`grandTotals.columns`); with value fields on the axis, a grand total per value field, `Total Sum of Amount`.
- **Insert Blank Line after Each Item** (`blankLine`) leaves an empty line after each item's group.

### Value fields

Each value cell shows its value field's value with the value field's number format; a value shown as a percentage (`showAs` of a share or a percentage difference) without a format of its own shows as `0.00%`. **For error values show** and **For empty cells show** (`errors`, `empty`) put their text in place of an error and of an empty value; a text that reads as a number is shown as the number (`0`).

Several value fields lie across the columns by default, or across the rows (`valuesOnRows`), at the place `valuesPosition` gives among the axis's fields. With **Show the Values Row** (`valuesRow`) a caption row stands over value fields alone on the columns.

### Filters area

The fields of the Filters area take a row each above the table, a field's name and its item (`(All)`, an item, or `(Multiple Items)`), with a blank row between them and the table; `pagesAcrossFirst` and `pagesWrap` lay them across and in runs, as Excel's options do. The table keeps its top-left cell, and moves down only when the rows above it are too few for the Filters area.

### When the PivotTable does not fit

A PivotTable laid out again over cells that hold something, over a table or over another PivotTable shrinks to its top-left cell, which shows `#SPILL!`, as Excel 365 does: its new layout is kept, its old cells are cleared, and the cells in the way are left as they are. A refresh of a cache lays out all its PivotTables in one step, undone in one step.

## Typing into a PivotTable

A session does not let an entry change a value of a PivotTable, as Excel does not: `setCells` throws an `InputError`, and `setCellText`, `setCellTexts`, `fillCellText` and `pasteText` return the refusal `pivot`. A cell's style and number format may change, and the blank row between the Filters area and the table takes anything.

```js
const r = session.setCellText('Pivot', 'B4', '100')
// r.ok === false, r.code === 'pivot'
```

With `xlcalc/pivot` loaded, text typed over a label names it, as in Excel, through `editPivotTable`:

- over an item's label it gives the item a caption; the name of another item of the same field moves that item there;
- over a field's name it names the field; the name of another field on the same axis swaps the two;
- over a value field's name it names the value field; a name of a field of the source, or of another value field, is rejected;
- over a subtotal's label it gives the field its subtotal caption, `?` standing for the item's label;
- over `Grand Total`, `Row Labels` or `Column Labels` it gives the PivotTable's caption;
- over a Filters area's item it shows that item (`(All)` shows all of them); a text that names no item is rejected.

A formula, a cleared label, and text over a value or over an empty cell of the area are rejected, with Excel's words.

## Refreshing and calculating

A PivotTable summarizes the rows of its pivot cache. With `xlcalc/pivot` loaded, a session refreshes the cache from its source and lays every PivotTable over it out again, as Excel's Refresh does:

```js
import { registerPivot, createSession } from 'xlcalc'
import { pivotEngine } from 'xlcalc/pivot'

registerPivot(pivotEngine)
const session = createSession(book)
session.editStructure({ op: 'refreshPivots' }) // Refresh All
session.editStructure({ op: 'refreshPivots', pivots: [{ sheet: 'Report', name: 'PivotTable1' }] })
```

The same lines serve Node and a browser bundle, the inline calculator and a pool, whose threads import the engine themselves. Without it, `refreshPivots` is refused (`pivot`) and the cells keep the values the file saved.

A PivotTable is not recalculated when the cells of its source change, only by a refresh, as in Excel: the cells of its area keep their values until then. Opening a workbook refreshes nothing, also a cache saved to refresh on opening (`refreshOnLoad`, which `getBook` returns as the file had it): the host refreshes it with `refreshPivots` after opening. One refresh is one step of undo. Refreshing one PivotTable refreshes its cache, and with it every PivotTable over that cache; Refresh All passes over the caches it cannot refresh, while naming a PivotTable of one is refused with the reason:

- the source is in another workbook, consolidates several ranges, is an external connection, the data model or scenarios;
- the cache does not allow refreshing (`enableRefresh: false`);
- the source is gone (a sheet, table or name no longer there);
- a header of the source is empty (Excel's "The PivotTable field name is not valid").

### The source

The source is a range of a sheet, a table (its header row and data rows, without the totals row), or a defined name over a range or a table. A name whose formula gives a reference (`OFFSET(Data!$A$1,0,0,COUNTA(Data!$A:$A),4)`) is calculated for the reference it gives at the moment of the refresh. A range of whole columns stops at the last row that holds something.

Each column is a field, named by its header; a repeated header is numbered (`Amount`, `Amount2`). A field keeps the values of its cells as calculated: formulas give their values, and a number in a cell with a date format is a date. Text in any case is one item: `North` and `north` are both `North`, the first spelling in the source, and after a refresh the item takes the first spelling the source has then.

A refresh keeps what the PivotTables were set to:

- a field whose header is gone leaves the PivotTable; a new column comes on no axis;
- an item gone from the source is kept in the cache as no longer there, as many per field as the cache's `retainItems` allows (all by default, none at 0);
- new items come after the others, and are hidden in a field where some items are hidden by hand (Excel's "Include new items in manual filter" off);
- a hidden item stays hidden, also when it comes back to the source;
- a page field whose item is gone shows the first item left.

### Summarizing

The 11 functions of a value field follow Excel's rules for blanks, text and errors:

| Values at the cell | `sum` | `count` | `countNums` | `average` | `max`, `min`, `product` | `stdDev`, `var` | `stdDevp`, `varp` |
|---|---|---|---|---|---|---|---|
| no record | empty | empty | empty | empty | empty | empty | empty |
| blanks only | empty | empty | empty | empty | empty | empty | empty |
| text or booleans only | 0 | their count | 0 | `#DIV/0!` | 0 | `#DIV/0!` | `#DIV/0!` |
| one number | the number | 1 | 1 | the number | the number | `#DIV/0!` | 0 |
| an error among them | the error | counted | not counted | the error | the error | the error | the error |

A date in a value field is its serial number. A total whose cells hold several errors returns the first, in the order the cache lists the items. Subtotals and grand totals summarize the values under them, not the values shown: the grand total of an average is the average of every record. A field's subtotals by other functions (Field Settings, Subtotals, Custom) summarize the same values by each function.

### Filters

Records count where every field lets them through:

- hidden items (unchecked in the filter list), on any field, also one on no axis (a slicer keeps its choice this way);
- a page field's item, or the items left checked with Select Multiple Items;
- label filters, by the item's caption: equals, begins with, ends with and contains take `?` and `*` (`~` before one takes it as it is), in any case; greater than and between compare text in Excel's order, a number item against a number by value;
- date filters on a field of dates, by whole days unless the filter takes the time: before, after, between, a specific date, a quarter or month of any year; the relative ones (today, this week from Sunday, last month, next quarter, year to date…) count from the session's clock, the day `TODAY` returns;
- value filters and Top 10, by a value field's values, among the items under the same item of the fields before: an empty value counts as 0, an error fails; Top 10 by items keeps the items tied with the last one it takes; by percent or sum, the items from the top (or the bottom) until their total reaches the share or the sum, no more (of two tied items, the first).

An axis shows the items its own filters leave: a value filter on the rows that leaves no row still shows the columns' items, empty.

A label or date filter and a value filter on one field both apply (Allow multiple filters per field). Totals summarize the records left: a grand total under Top 10 is the total of the items shown.

### Sorting

A field placed on an axis lists its items ascending. Sorting by label puts the texts of a custom list first in the list's order (Sun…Sat, Sunday…Saturday, Jan…Dec, January…December, and the workbook's locale's, with Use Custom Lists when Sorting on), then numbers and dates, text in Excel's order, booleans, errors, and the blank last; descending turns the whole order round, the blank first. Sorting by a value field's values puts an empty value first going up and last going down, numbers by value, and errors last either way, at the grand total of the other axis or at the items it names; ties keep their order, and under an item of a field before go by the items' values over all of them. A field without a sort (sorted by hand, or with "Sort automatically every time the report is updated" off) keeps its items' order, new items after them.

## Checked against Excel

Excel for Windows built PivotTables through its object model and saved them (the sets basic, layout, calc, edit,
filter, refresh and sort): for every PivotTable after each step, the PivotTable `readXlsx` reads lies where Excel showed
it, its values begin where Excel's did, and `GETPIVOTDATA` of every value Excel showed, by the items of its line and
column, returns that value (over 2,000 cells); the fields, value fields, filters and sorting read are those the steps
made. 101 `GETPIVOTDATA` formulas over these PivotTables in Excel for Windows (the set gpd-1), and formulas over a
PivotTable of another workbook open, closed and reopened, give Excel's values. [Development](development.md#pivottables)
has the details.

The layout was compared with Excel for Windows 16.0 (en-US) on every book of the sets in `test/excel/win/pivot/` (`basic`, `layout`, `filter`, `sort`, `edit`, `refresh`, `calc`, `gpd-1`): each PivotTable laid out again from what Excel saved gives Excel's lines, place, values and number formats (`test/pivot-layout-books.test.ts`), and each entry of the set `edit` gives Excel's refusal or Excel's new definition (`test/pivot-entry.test.ts`).

Excel for Windows built PivotTables over 48 rows of text, numbers with blanks, text, an error and a boolean among them, dates, numbers kept as text and long text, and saved them: 48 PivotTables of the 11 functions, columns, two value fields and kinds of items; 61 of filters; 27 of sorting; 34 of refreshing after edits of the source (new items, items gone with each Number of items to retain, values and types changed, an error appearing, a table grown and resized, a name made smaller, a header renamed, the source edited before a filter, a sort, a layout change, a field moved or added, two PivotTables of one cache, two caches). Each PivotTable laid out again from the cache Excel saved, and refreshed from its source, gives the cells Excel showed, with one known difference (`docs/limitations.md`). Held-out sets, predicted before Excel ran them, matched 20 of 20 of layout, 20 of 22 of filters and 17 of 18 of sorting; their three misses are rules now.
