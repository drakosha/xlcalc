# PivotTables

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

A PivotTable is not recalculated when the cells of its source change, only by a refresh, as in Excel: the cells of its area keep their values until then. One refresh is one step of undo. Refreshing one PivotTable refreshes its cache, and with it every PivotTable over that cache; Refresh All passes over the caches it cannot refresh, while naming a PivotTable of one is refused with the reason:

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
- value filters and Top 10, by a value field's values, among the items under the same item of the fields before: an empty value counts as 0, an error fails; Top 10 keeps the items tied with the last one it takes; by percent or sum, the items from the top (or the bottom) until their total reaches the share or the sum.

A label or date filter and a value filter on one field both apply (Allow multiple filters per field). Totals summarize the records left: a grand total under Top 10 is the total of the items shown.

### Sorting

A field placed on an axis lists its items ascending. Sorting by label puts the texts of a custom list first in the list's order (Sun…Sat, Sunday…Saturday, Jan…Dec, January…December, and the workbook's locale's, with Use Custom Lists when Sorting on), then numbers and dates, text in Excel's order, booleans, errors, and the blank last; descending turns the whole order round, the blank first. Sorting by a value field's values puts an empty value first going up and last going down, numbers by value, and errors last either way, at the grand total of the other axis or at the items it names; ties keep their order. A field sorted by hand keeps its order, new items after it; with "Sort automatically every time the report is updated" off, a sort is not applied again.

### Checked against Excel

Excel for Windows built PivotTables over 48 rows of text, numbers with blanks, text, an error and a boolean among them, dates, numbers kept as text and long text, and saved them: 48 PivotTables of the 11 functions, columns, two value fields and kinds of items; 61 of filters; 27 of sorting; 34 of refreshing after edits of the source (new items, items gone with each Number of items to retain, values and types changed, an error appearing, a table grown and resized, a name made smaller, a header renamed, the source edited before a filter, a sort, a layout change, a field moved or added, two PivotTables of one cache, two caches). Each PivotTable laid out again from the cache Excel saved, and refreshed from its source, gives the cells Excel showed, with three known differences (`docs/limitations.md`).
