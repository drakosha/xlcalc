# PivotTables

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

## Checked against Excel

The layout was compared with Excel for Windows 16.0 (en-US) on every book of the sets in `test/excel/win/pivot/` (`basic`, `layout`, `filter`, `sort`, `edit`, `refresh`, `calc`, `gpd-1`): each PivotTable laid out again from what Excel saved gives Excel's lines, place, values and number formats (`test/pivot-layout-books.test.ts`), and each entry of the set `edit` gives Excel's refusal or Excel's new definition (`test/pivot-entry.test.ts`).
