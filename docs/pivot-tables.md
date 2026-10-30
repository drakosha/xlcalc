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
  one; a cell outside every PivotTable is `#REF!`.
- `data_field` is the value field's name (`Sum of Price`) or its source field's (`Price`), in any case. Given as a
  reference to a cell holding the name, it is `#REF!`, as in Excel; text a formula makes (`TEXT(A1,"0")`) is taken.
- A field is named as the PivotTable or the source names it, in any case. An item is matched by its label, in any case:
  `"North"`, `"(blank)"` for the blank item, a number's or a date's label as the PivotTable shows it (`"10"`, not
  `"10.0"`; `"1/1/2024"`, not `"2024-01-01"`), a text item before a number of the same label. A number matches a number
  or a date of that value (`DATE(2024,1,1)`, `45292`), a boolean a boolean.
- The pairs pick the line of the rows and the column that name exactly their fields: naming all of them gives a value
  cell, naming fewer a subtotal or the grand total where the PivotTable shows one. A field of the Filters area is
  matched against the item it shows. What the layout does not show (an item filtered out, an inner field without its
  outer one, a field on no axis, a pair without its item) is `#REF!`.
- The answer is the value of that cell, an error as it is, an empty cell 0.
- A PivotTable whose definition changed without being laid out again (`PivotCache.stale`) is `#REF!`.
- A PivotTable in another workbook (`[1]P!$A$3`) is `#REF!` when calculated; a file opened keeps the value Excel saved
  for such a formula, as it keeps one over the functions of connections ([Values a file stores](formulas.md#values-a-file-stores)).

The formula learns the PivotTable's cells as its references, so an edit of them recalculates it; a change of the
PivotTable builds the workbook again. Its call takes up to 254 arguments (126 pairs).

## Checked against Excel

Excel for Windows built PivotTables through its object model and saved them (the sets basic, layout, calc, edit,
filter, refresh and sort): for every PivotTable after each step, the PivotTable `readXlsx` reads lies where Excel showed
it, its values begin where Excel's did, and `GETPIVOTDATA` of every value Excel showed, by the items of its line and
column, returns that value (over 2,000 cells); the fields, value fields, filters and sorting read are those the steps
made. 101 `GETPIVOTDATA` formulas over these PivotTables in Excel for Windows (the set gpd-1), and formulas over a
PivotTable of another workbook open, closed and reopened, give Excel's values. [Development](development.md#pivottables)
has the details.
