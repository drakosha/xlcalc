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

Without `xlcalc/pivot` registered, the engine does not refresh a PivotTable or lay it out again: its cells hold the
values the file saved, as values of the sheet. A session creates and changes PivotTables either way
([Creating and changing a PivotTable](#creating-and-changing-a-pivottable)), and `writeXlsx` and `writeKept` write them
([Writing PivotTables](#writing-pivottables)).

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
  `"10.0"`; `"1/1/2024"`, not `"2024-01-01"`), a text item before a number of the same label. An item renamed in the
  PivotTable is named by its new name, and a text item by its value too (`"East"` renamed `Eastern` by either; `FALSE`
  renamed `No` by `"No"` or `FALSE`, not `"FALSE"`). A number matches a number
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

Each value cell shows its value field's value ([Show Values As](#show-values-as) applied) with the value field's number format; a value shown as a percentage (`showAs` of a share or a percentage difference) without a format of its own shows as `0.00%`. The line of an item over its group, which shows no value, takes the format too. **For error values show** and **For empty cells show** (`errors`, `empty`) put their text in place of an error and of an empty value; a text that reads as a number is shown as the number (`0`).

Several value fields lie across the columns by default, or across the rows (`valuesOnRows`), at the place `valuesPosition` gives among the axis's fields. With **Show the Values Row** (`valuesRow`) a caption row stands over value fields alone on the columns.

### Filters area

The fields of the Filters area take a row each above the table, a field's name and its item (`(All)`, an item, or `(Multiple Items)`), with a blank row between them and the table; `pagesAcrossFirst` and `pagesWrap` lay them across and in runs, as Excel's options do. The table keeps its top-left cell, and moves down only when the rows above it are too few for the Filters area.

### When the PivotTable does not fit

A PivotTable laid out again over cells that hold something, over a table or over another PivotTable shrinks to its top-left cell, which shows `#SPILL!`, as Excel 365 does: its new layout is kept, its old cells are cleared, and the cells in the way are left as they are. A refresh of a cache lays out all its PivotTables in one step, undone in one step.

### The look of its cells

The cells of the area keep what Excel itself keeps in them: a compact form's row labels are aligned left and indented
by their level (an item by its level, a subtotal at the bottom by its field's, the grand total not at all). With Merge
and center cells with labels (`mergeLabels`), the headers are centered, the outline form's row labels centered, and the
labels of the tabular form's rows and of the columns centered both ways and wrapped, each merged over the lines it
labels: an item over its items' rows or columns, a subtotal and the grand total across the labels' columns or down the
labels' rows. A PivotTable's style (`style`: `PivotStyleLight16` and the others, and its options) is kept and written,
for Excel to draw; the engine does not draw it, as it does not draw a table's style.

## Creating and changing a PivotTable

`editStructure` creates, changes, moves and deletes PivotTables, as Excel's PivotTable commands do, in every kind of
session, each operation one step of undo:

```js
session.editStructure({ op: 'addPivotTable', sheet: 'Report', cell: 'A3', source: { type: 'range', sheet: 'Data', range: 'A1:D200' }, definition: { rows: [0], values: [{ field: 3 }] } })
session.editStructure({ op: 'editPivotTable', sheet: 'Report', pivot: 'PivotTable1', set: { columns: [1], layout: 'tabular' } })
session.editStructure({ op: 'movePivotTable', sheet: 'Report', pivot: 'PivotTable1', to: { sheet: 'Summary', cell: 'B2' } })
```

- `addPivotTable` makes a PivotTable at a cell, or at A3 of a new sheet (`newSheet`). A source by range, table or name
  makes a new pivot cache, its fields named by the source's header row (a repeated name gets `2`, `3`, as Excel names
  it); `{ cache: id }` shares a cache of the workbook. Its name is the first free `PivotTable1`, `PivotTable2` without
  `name`; its style `PivotStyleLight16`. An empty PivotTable takes three columns and 18 rows, as Excel draws one. A place
  over cells that hold something, a table or another PivotTable, an empty header cell, a source without a row of data
  and a name already on the sheet are refused (`pivot`).
- `editPivotTable` sets fields of the definition (`set`, in the form `PivotTable` holds them): a value replaces a field,
  `null` removes it, `fields` edits the PivotTable's fields by their index. A field lies on one axis, and a value field
  without a name is named as Excel names it (`Sum of Amount`, `Сумма по полю Amount` in ru-RU, `Sum of Amount2` for a
  second). `source` is Change Data Source: the fields kept by name, a cache of its own when the old one is shared.
- `renamePivotTable`, `movePivotTable` (its cells move as cut and paste moves them, to another sheet too) and
  `deletePivotTable` (its cells cleared, its cache gone with its last PivotTable).
- `refreshPivots` refreshes the caches of the PivotTables named, or every cache that can be (Refresh All).

With `xlcalc/pivot` registered, a new PivotTable is refreshed from its source and laid out, and a changed or moved one
is laid out again from its cache, as Excel does without a refresh; a new or moved PivotTable laid over cells it cannot
take is refused. Without it, the definition changes, the cells stay, and the PivotTable's cache is marked `stale`; the
operation's answer names those PivotTables (`stale`), and a file written then asks Excel to refresh them on opening.

The operations of the structure treat a PivotTable as Excel does: rows and columns inserted or deleted before it, a
move of all its cells and a sheet renamed move it; an operation that would cut it (rows inserted into it, part of it
deleted or moved, a paste, a fill, a merge or a table over part of it) is refused (`pivot`); all its cells deleted, or
its sheet, delete it, and the slicers and timelines over it keep their caches, naming it no longer. A source by range
moves with its cells as a reference does, and stays as it was when they or their sheet are deleted.

## Writing PivotTables

`writeXlsx` writes the workbook's pivot caches with their rows, PivotTables, slicers and timelines as Excel 365 writes
them: a cache's fields with the kinds of value they hold and their items, a PivotTable's fields, axes, value fields,
filters, sorting, totals, form and options, the layout it was last laid out with, the formats of its areas and the
areas of its conditional formats, its style; slicers and timelines with their caches and frames. Excel opens the file
and shows each PivotTable as written, without refreshing it, and lays it out the same after a refresh.

`writeKept` keeps a PivotTable, a cache, a slicer or a timeline the session did not change as the file has it, byte for
byte; one moved with its cells has its place written back. One the session changed or refreshed is written again from
the workbook, a new one gets parts of its own, a deleted one goes with its parts. Slicers and timelines added or
deleted are refused (`slicers`).

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

Source formula values and calculated fields keep their binary64 precision during refresh. The public result's 15-significant-digit rounding is applied after the pivot calculation. Totals follow Excel's accumulation order, including its record partitioning when filters remove records; Show Values As reads those same totals. The measured numeric rules and their exact comparisons are in [development.md](development.md#pivottables-refreshing-and-calculating).

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

A new field lists its items ascending. Numbers and dates precede text, then booleans, errors and the blank. With Use Custom Lists when Sorting enabled, texts belonging to a common custom list use that list's order. Texts belonging to disjoint lists follow Excel's partition sort: comparing them is not a total ordering, so sorting the same field again can change their order. A label sort changes the field's `items`; reading an existing layout keeps the order saved there.

Sorting by a value field puts empty values first going up and last going down, numbers by value, and errors last either way. Equal values are ordered by Excel's middle-pivot partitioning over the field's item order, independently of the last saved row or column layout. A preceding label sort can therefore change the order of equal values on a later value sort.

`registerPivot(pivotEngine, { sortThreads: 2 })` selects the Excel ordering profile. The default is 2, as on the Windows reference machine. Above 512 items Excel sorts partitions and merges them; this can change ties. Set `sortThreads: 1` for one partition. The setting changes the ordering profile for subsequent pivot calculations, including in existing sessions and a pool's threads; it creates no worker threads. Recovery replays each completed operation with the profile that operation used.

The classic layout (fields dragged in the grid) keeps the row of the columns' area with no column field: the value field's name above it, `Total` over its values. A field without a sort (sorted by hand, or with "Sort automatically every time the report is updated" off) keeps its items' order, new items after them.

### Show Values As

A value field shows its values as Excel's Show Values As does (`showAs`, with `baseField` and `baseItem` where the calculation takes them), over the values of the items shown:

| `showAs` | A cell shows | Its totals |
|---|---|---|
| `percentOfTotal`, `percentOfRow`, `percentOfColumn` | its share of the grand total, of its row's total, of its column's total | their shares, the grand total 100% |
| `percentOfParentRow`, `percentOfParentColumn` | its share of the item above it on the rows (columns) | the grand total 100% |
| `percentOfParent` | its share of the base field's item it lies under | the base field's items 100%, totals above them empty |
| `index` | value × grand total / (row total × column total) | the same |
| `difference`, `percent`, `percentDiff` | its difference from the base item's value at the same items, its share of it, the difference as a share | the base item's own cells empty (100% for `percent`), totals above the base field empty |
| `runTotal`, `percentOfRunningTotal` | the sum of the base field's items up to it, in the order shown, and that sum's share of all of them | totals above the base field empty |
| `rankAscending`, `rankDescending` | its rank among the base field's items, 1 for the smallest (largest); ties share a rank, and the next value takes the next one (1, 2, 2, 3) | totals above the base field empty |

- A cell with no record counts as 0: it shows 0% of a total, a running total so far, a difference of 0 less the base. A share of a base item and a percentage difference show `#NULL!` there, as Excel does.
- A share of an empty value (a parent with no record, a base item's empty cell, a running total of empty cells) is empty; of 0, or of an empty row or column total, `#DIV/0!`.
- The previous and next items (`baseItem: 'previous'`, `'next'`) are those with a line at the same items under them: `West`'s previous for `Banana` is the region before it that has a `Banana` line. The first item's previous, and the last's next, is the item itself.
- A base item with no line at the cell's items (`East` has no `Banana`) is `#N/A`; a base item hidden or filtered out, a base field on no axis (on the Filters area too) and no base field are `#N/A` in every cell, the totals too.
- An error among the values is the cell's error, or the error of the value it is set against; a running total takes the last error it passes; an error has no rank; a share of a parent is empty where the value or its parent is an error.
- Hidden and filtered items take no part: shares are of the totals of the items shown, and running totals, ranks and the previous and next items pass over the hidden ones. A Filters area's item or items limit the records as for any value.
- Sorting by a value field sorts by its values before Show Values As: a field sorted by a rank goes by the sums ranked.

## Grouping

A field of dates can show them by periods, a field of numbers by intervals, and any field its items gathered by hand into groups, as Excel's Group dialog does. A grouping belongs to the cache: every PivotTable over it sees the groups, and a refresh makes them again over the new values.

```js
session.editStructure({ op: 'groupPivotField', sheet: 'Report', pivot: 'PivotTable1', field: 2, group: { by: 'dates', periods: ['months', 'years'] } })
session.editStructure({ op: 'groupPivotField', sheet: 'Report', pivot: 'PivotTable1', field: 5, group: { by: 'numbers', interval: 10 } })
session.editStructure({ op: 'groupPivotField', sheet: 'Report', pivot: 'PivotTable1', field: 0, group: { by: 'items', items: [0, 3] } })
session.editStructure({ op: 'renamePivotGroup', sheet: 'Report', pivot: 'PivotTable1', field: 10, item: 0, name: 'Coast' })
session.editStructure({ op: 'ungroupPivotField', sheet: 'Report', pivot: 'PivotTable1', field: 10 })
```

Fields are named by their index in the cache and items by their place in the PivotTable's field, as elsewhere. Each operation lays out every PivotTable of the cache again and is one step of undo; without `xlcalc/pivot` it is refused (`pivot`).

- **Dates by periods.** Each period makes a field of its own after the source's, `Months (Date)`, `Years (Date)`, as Excel 365 does: the PivotTable grouped shows them, coarsest first, in place of the dates; another PivotTable of the cache that shows the dates shows the periods before them. The labels are Excel's: `:00`…`:59` for seconds and minutes, `12 AM`…`11 PM`, `1-Jan`…`31-Dec` (29 February among them), `Jan`…`Dec`, `Qtr1`…`Qtr4`, the years. Days by a number of days (`days: 7`) make groups from the start, labelled `1/1/2024 - 1/7/2024`, as many as whole intervals reach the end, the last one cut short at it.
- **Numbers by an interval.** The field itself shows the groups: `1-2`, `3-4` over whole numbers, `1.03-11.03` over others; the last group of whole numbers takes the end in (`7-9` of 1 to 9 by 2).
- **Start and end.** Without them they are the least and greatest value, and a date field of whole days ends a day after its last date (12/22/2025 ends at 12/23/2025), as Excel takes it. A value before the start goes to the first group, `<1/1/2024`, one after the end to the last, `>12/23/2025`. A blank date goes to the first group; among numbers, the first group is `(blank)`, or `<10 or (blank)` with a start of one's own.
- **Items by hand.** The items chosen make a group of a new field, `Region2` over `Region`, which every PivotTable showing `Region` shows before it; each item not chosen is a group of its own, and further groups are `Group2`, `Group3`. The groups follow the order of their first items. Renaming a group changes its label for every PivotTable of the cache; text typed over the label changes it on that PivotTable alone, as a caption ([Typing into a PivotTable](#typing-into-a-pivottable)). Ungrouping takes away the groups given (`items`) or all of them, the field with the last.
- **A refresh** takes an automatic start and end from the new values again and makes the labels again: a new group (a year before or after the others) comes after the groups the PivotTable lists, a new item of a field grouped by hand as a group of its own.
- **Refused, as in Excel:** grouping text among dates or numbers ("Cannot group that selection"), grouping a field with calculated items, and ungrouping a field that is not grouped.

## Calculated fields and items

A calculated field is a formula over other fields by name; a calculated item a formula over other items of its field. Both belong to the cache, and every PivotTable over it can show them.

```js
session.editStructure({ op: 'setPivotCalculatedField', sheet: 'Report', pivot: 'PivotTable1', name: 'Unit price', formula: '=Amount/Qty' })
session.editStructure({ op: 'setPivotCalculatedItem', sheet: 'Report', pivot: 'PivotTable1', field: 0, name: 'Coast', formula: '=East+West' })
session.editStructure({ op: 'setPivotCalculatedField', sheet: 'Report', pivot: 'PivotTable1', name: 'Unit price', formula: null }) // deleted
```

A name the cache has changes the formula (Excel's Modify), and a null formula deletes the field or item. Adding a calculated field does not place it on Values; `editPivotTable` does.

- **A calculated field** is calculated in each cell of the PivotTable, totals too, over the sums of the fields it names, not row by row: `Amount/Qty` over two rows of 100 and 100, Qty 1 and 9, is 200/10 = 20. A field is named as it is (`Amount`), in quotes with spaces in its name (`'Unit price'`), and may name another calculated field. The sum of a field with errors is the last of them; of a field with no number at the cell, 0 (`MIN(Price,10*Qty)` there is 0).
- Every cell of the rows and columns shown gets the formula, one no record reaches too, over sums of 0: `Price/Qty` there is `#DIV/0!`. A subtotal by a function of its own (Field Settings, Custom) is empty.
- **A calculated item** is calculated in each cell over the values of the items it names at the same place, by the value field's function, and is counted in the totals over it: `=East+West` shows East and West added, and the grand total takes them once more, as Excel does. It shows under every item of the field before it, with every item of the fields after it that any item of its field has (`7`, under West alone, shows under `North+South` too, as 0). An item is named as it is (`East`), in quotes when it is not a plain name (`'Dragonfruit from the far south'`, `'10'`); the cache keeps the formula as Excel's file does, `Region[East]+Region[West]`.
- The formulas are calculated by the engine: any function of the sheet Excel allows there (`IF`, `ROUND`, `MAX`, `IFERROR`, `AND`), constants, booleans, errors and division by zero as in a cell. A reference to a cell or a range is refused ("References, names, and arrays are not supported in PivotTable formulas"), so is a formula that does not read or names itself.
- **Refused, as in Excel:** a calculated item in a grouped field, and in a PivotTable whose value fields average or take a deviation or variance ("Averages, standard deviations, and variances are not supported when a PivotTable report has calculated items").

## Slicers and timelines

A slicer filters PivotTables by the items of one field, a timeline by a range of dates. A session adds, changes and deletes them through `editStructure`, as Insert Slicer, Insert Timeline, Report Connections and the slicer's buttons do in Excel:

```js
session.editStructure({ op: 'addSlicer', sheet: 'Report', slicer: {}, cache: { field: 'Region', pivots: [{ sheet: 'Report', name: 'PivotTable1' }] } })
// the slicer Region over the new cache Slicer_Region, captioned Region, placed right of the PivotTable
const items = session.getBook().slicerCaches[0].items // [{ item: 2 }, { item: 0 }, …] in the slicer's order
session.editStructure({ op: 'editSlicer', slicer: 'Region', cache: { items: items.map((x) => ({ item: x.item, selected: x.item === 0 })) } })
// PivotTable1 shows North alone
session.editStructure({ op: 'addTimeline', sheet: 'Report', timeline: {}, cache: { field: 'Date', pivots: [{ sheet: 'Report', name: 'PivotTable1' }] } })
session.editStructure({ op: 'editTimeline', timeline: 'Date', cache: { selection: { start: '2025-01-01T00:00:00', end: '2025-06-30T00:00:00' } } })
// the first half of 2025
```

- A slicer chooses items of its field (`SlicerCache.items`, an index into the cache field's `items` with `selected: false` for an item not chosen); at least one stays chosen. Its choice is the filter of every PivotTable its cache names: the items not chosen are hidden on that field of each, whether the field lies on an axis or on none. A PivotTable's hidden items are what the calculation reads, so a PivotTable filtered by a slicer is filtered the same once the file is opened again, and a refresh keeps the choice: an item new in the source comes unchosen, as in a manual filter.
- A slicer over a field the cache lists no items of (numbers on no axis) lists them in the cache, as Excel does.
- A new slicer over a field whose items a PivotTable already hides, or whose Filters area shows one item, starts with that choice, and the PivotTable stays as it is. A choice on a page field hides the items not chosen and shows the field's items by Select Multiple Items, unless the one item it shows is the one chosen. Connecting another PivotTable (`cache: { pivots }`) gives it the slicer's choice; a PivotTable disconnected shows every item of the field again.
- A field on no axis filters a PivotTable only while a slicer or timeline over the PivotTable names it. Excel saves the hidden items of a slicer disconnected, and a deleted timeline's filter, and applies neither, also on a refresh; so does the calculation here.
- A slicer without cross filtering (Visually indicate items with no data off, `noData.mark: false`) marks no item, and Excel's object model tells none has data.
- After a choice, the PivotTables are laid out again from their cache, without reading the source, as Excel does. Each slicer then knows which of its items have data: an item has data when a record of it passes the other fields' filters (other slicers, page fields, label and date filters) and has a value for a value field. A region whose only record has no price has no data under Sum of Price. Items with no data come last, marked, unless the cache says otherwise (`noData`); items are ordered by label, ascending by default, the texts of custom lists first (`sort`, `customListSort`).
- A timeline's dates (`TimelineCache.selection`) filter its field on every PivotTable it names, by whole days: between the two dates, or on the one day when they are the same. The field stays where it is, on no axis as a rule. The level shown (`Timeline.level`) does not change the filter; `selectionLevel` is the widest level whose periods the dates span whole (2025-01-01 to 2025-06-30: quarters). `selection: null` is Clear Filter.
- Deleting a slicer or a timeline deletes its cache with the last one over it. The PivotTables keep a slicer's choice, its items stay hidden, and lose a timeline's: its dates no longer filter. A slicer or timeline whose PivotTables are all gone stays, filtering none, with its choice, as Excel keeps it when the sheet of its PivotTable is deleted.
- A name is unique among the workbook's slicers and timelines; without one, a slicer takes its field's (`Region`, `Region 1`…). A new cache is `Slicer_Region`, a timeline's `NativeTimeline_Date`. A field that is not in the source, a timeline over a field without dates and a name already used are rejected (`pivot`).
- Each choice is one step of undo. Without `xlcalc/pivot` the choice is written all the same, and the PivotTables keep their cells with their cache marked `stale`.

A table's slicer (`SlicerCache.table`) filters its table by the table's AutoFilter. It is read and kept as the file has it, and its look, place and order change; its choice is the AutoFilter's, which the workbook does not hold (`docs/limitations.md`), so `editSlicer` rejects items for it.

## Checked against Excel

Excel for Windows built PivotTables through its object model and saved them (the sets basic, layout, calc, edit,
filter, refresh and sort): for every PivotTable after each step, the PivotTable `readXlsx` reads lies where Excel showed
it, its values begin where Excel's did, and `GETPIVOTDATA` of every value Excel showed, by the items of its line and
column, returns that value (over 2,000 cells); the fields, value fields, filters and sorting read are those the steps
made; the slicers and timelines of the set slicer read with the items chosen and with data, and the dates and level,
Excel showed. 101 `GETPIVOTDATA` formulas over these PivotTables in Excel for Windows (the set gpd-1), and formulas over a
PivotTable of another workbook open, closed and reopened, give Excel's values; of the held-out sets gpd-held (55
formulas), gpd-held-2 (61) and gpd-held-3 (31), predicted before they ran, 51, 57 and 30 held, and the rest were
fitted. [Development](development.md#pivottables)
has the details.

Excel for Windows grouped and calculated the PivotTables of the sets group (dates by every period and by numbers of days, times, numbers with Excel's and their own start and end, items by hand, ungrouping, refusals) and calc (calculated fields with functions, errors and division by zero, on rows, columns and subtotals; calculated items; both over two PivotTables of a cache): each PivotTable laid out again from the cache Excel saved and refreshed from its source gives the cells Excel showed, calculated-field totals included (`test/pivot-excel.test.ts` and the exact-number checks in `test/pivot-precision.test.ts`), and the operations played over the rows give the same PivotTables.

The layout was compared with Excel for Windows 16.0 (en-US) on every book of the sets in `test/excel/win/pivot/` (`basic`, `layout`, `filter`, `sort`, `edit`, `refresh`, `calc`, `gpd-1`): each PivotTable laid out again from what Excel saved gives Excel's lines, place, values and number formats (`test/pivot-layout-books.test.ts`), and each entry of the set `edit` gives Excel's refusal or Excel's new definition (`test/pivot-entry.test.ts`).

Excel for Windows built PivotTables over 48 rows of text, numbers with blanks, text, an error and a boolean among them, dates, numbers kept as text and long text, and saved them: 48 PivotTables of the 11 functions, columns, two value fields and kinds of items; 61 of filters; 27 of sorting; 34 of refreshing after edits of the source (new items, items gone with each Number of items to retain, values and types changed, an error appearing, a table grown and resized, a name made smaller, a header renamed, the source edited before a filter, a sort, a layout change, a field moved or added, two PivotTables of one cache, two caches). Each PivotTable laid out again from the cache Excel saved, and refreshed from its source, gives the cells Excel showed, with one known difference (`docs/limitations.md`). Held-out sets, predicted before Excel ran them, matched 20 of 20 of layout, 20 of 22 of filters and 17 of 18 of sorting; their three misses are rules now.

Show Values As was compared on 36 PivotTables of the 15 calculations (over two row fields and a column field, the base field inner and outer, on rows and on columns, fixed, previous and next base items, a blank one, values with errors and empty cells, beside a plain value field, tabular, sorted), 22 more of hidden and filtered items, Filters area items, Top 10, items with no data shown, a base item hidden, a base field off the axes, sorting by the values shown, value fields on the rows, an average and a count, and a held-out set of 20 over other rows. Laid out again from the saved cache and refreshed from the source, each gives the cells and number formats Excel showed, a share to within one part in 10¹⁴ (`docs/limitations.md`).

Excel for Windows added slicers and timelines through its object model and chose their items and dates (the set slicer, 16 cases): a slicer on a row field, on fields on no axis (booleans, numbers with one kept as text), one item, several, the blank, all of them again; over two PivotTables and over one of two; two slicers, one leaving the other's items without data; over a field with items hidden; with a page field; a timeline's range, level and single day, over dates on the rows, beside a slicer. The same operations through a session give the cells Excel showed in every PivotTable and each slicer's items, chosen and with data, in Excel's order, and each timeline's dates and level (`test/pivot-slicers-excel.test.ts`). Each workbook Excel saved holds the choice as the session writes it. Deleting a slicer, a timeline, and the sheet of a PivotTable they filter (the set slicer-delete, 5 cases), and refreshing the workbooks saved after (slicer-file), gave the rules of deletion above. The held-out set slicer-held over other rows (16 cases: slicers over a page field and fields on no axis, over two PivotTables, crossing, sorted, items without data placed or hidden, connected and disconnected; timelines by levels and over one day) held 47 of 50 reads as predicted; its three misses are the rules of disconnecting and of no cross filtering above.

Excel for Windows opened the files `writeXlsx` wrote of every book of the sets basic, layout, filter, sort and slicer
(the books read and written again), and the files `writeKept` wrote after each entry of the set edit made by a session
over the book of its first step (234 files): each PivotTable read at once, without a refresh, shows the cells, place and
values Excel showed for its own book, and after `RefreshTable`, which answers true, the cells Excel's own book shows
after one; Excel saved the slicers and timelines read back as written. The held-out sets of layout, filter and sort,
predicted before Excel ran them, matched 31 of 37, 57 of 61 and 27 of 27; the misses were two bugs of reading and the
way Excel writes a subtotal's functions, a field's repeated labels and a label filter, rules now. The look of the cells
was compared with every book of basic, layout, filter, sort and edit and with Merge and center cells with labels in
three forms (the set merge-1).
