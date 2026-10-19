# Outlines

Rows and columns of a sheet can be grouped, as with Excel's Data > Group: the bars beside the row numbers and above the column letters, a `−` or `+` button on each group's summary row or column, and the level buttons in the corner. The library returns a sheet's groups as Excel draws them, with `outlineGroups`, without the engine, and implements Excel's commands on them (Group, Ungroup, Show Detail, Hide Detail and the level buttons) as structure operations: `editStructure(book, op)`, or `session.editStructure(op)` with one revision and one undo step, and on a sheet alone, without the engine, `outlineEdit`. A grid draws the groups from `getLayout()` and sends the commands; the rules below are those of Excel for Windows, verified through its object model.

```js
import { createSession, outlineGroups } from 'xlcalc'

const session = createSession({
  sheets: [{
    name: 'S',
    cells: { A3: { value: 3 }, A6: { value: 6 }, A9: { value: 9 }, B1: { formula: 'SUBTOTAL(109,A3:A10)' } },
    rows: { '3:4': { outlineLevel: 1 }, '5:7': { outlineLevel: 2 }, '8:10': { outlineLevel: 1 } },
  }],
})
session.editStructure({ op: 'hideDetail', sheet: 'S', row: 8 })
session.editStructure({ op: 'hideDetail', sheet: 'S', row: 11 })
const r = session.editStructure({ op: 'showDetail', sheet: 'S', row: 11 })
console.log(r.changes.map((c) => [c.cell, c.value])) // → [ [ 'B1', 12 ] ]
console.log(outlineGroups(session.getLayout().sheets[0]).rows)
// → [
//     { first: 3, last: 10, level: 1, summary: 11, collapsed: false },
//     { first: 5, last: 7, level: 2, summary: 8, collapsed: true }
//   ]
```

## The groups

A row's outline level (`RowInput.outlineLevel`, 1 to 7) says how deeply it is grouped, and likewise for a column. `outlineGroups(sheet)` accepts a sheet or its layout and returns each axis's groups, each as `{ first, last, level, summary, collapsed }`, with columns by number:

- A group is the longest run of rows (or columns) at its level or deeper. Rows 3:4 and 8:10 at level 1 with 5:7 at level 2 form the group 3:10 at level 1 and 5:7 at level 2. Two adjacent runs form one group: grouping rows 3:5 and then rows 6:8 makes one group 3:8; a row without a level between them keeps them separate.
- The summary row, which holds the group's button, is the row after the group: row 11 for 3:10, below by default, or, with summaries above (`SheetInput.outline`, `summaryBelow: false`), the row before it, row 2. Groups that end together share their summary row. A group at the edge of the sheet, where its summary row would fall outside it, has none (`summary: null`), and Excel draws no button for it.
- `collapsed` is the marker Excel keeps on the summary row (see below); a group without a summary row is collapsed when all its rows are hidden.

The groups are listed in order of their first row, the outer group before the inner one where two start together.

## The markers Excel keeps

Excel marks a summary row as collapsed (xlsx `collapsed`, `RowInput.collapsed`) whenever all the rows of its detail, the deeper rows next to it, are hidden, however they were hidden: by its button, by a level button, or manually. It removes the marker as soon as one of them becomes visible. Rows 2:4 hidden through `setLayout`, with no marker on row 5, form a collapsed group: Excel opens such a file with the group collapsed and saves it with the marker.

One more thing keeps an inner group collapsed: a summary row inside a group whose rows are all hidden keeps the marker it had. Hide 5:7 (row 8 gets the marker), then 3:10 (row 11 gets the marker, and row 8 is hidden with its marker), then show 3:10: rows 3, 4, 8, 9 and 10 become visible, and 5:7 stay hidden under the marker on row 8, as in the example above. If 3:10 is hidden first, row 8 gets no marker, and showing 3:10 shows all of it.

A group at the edge of the sheet keeps its marker on its row or column at the edge: row 1 with summaries above, row 1,048,576, or column A with summaries to the left.

`outlineGroups`, the commands and `writeXlsx` interpret the markers this way, whatever markers a workbook contains; `setLayout` sets the fields as given, and `readXlsx` reads a file's markers as they are.

## The commands

Each command is a `StructureOp` on one sheet; a session's `editStructure` recalculates what `SUBTOTAL` and `AGGREGATE` count over the rows hidden or shown, and nothing else. No cell moves: `mapCell` maps every cell to its own position.

- `{ op: 'group', sheet, rows: '5:7' }` (or `columns: 'B:D'`, or a single row or column, `'5'` or `'C'`) moves each row or column one level deeper. One at level 7, the deepest, stays there, and Group is rejected (`outline_deep`) when all of them are at level 7. Grouping hides and shows nothing.
- `{ op: 'ungroup', sheet, rows }` removes one level from each row or column that has one, and is rejected (`not_grouped`) when none has. Hidden rows stay hidden: ungrouping a collapsed group leaves its rows hidden.
- `{ op: 'hideDetail', sheet, row: 8 }` and `showDetail` (or `column: 'H'`) are a group's buttons, applied at one row or column, as Excel's object model applies them. At a summary row they act on its entire detail: with rows 3:7 at level 1 and 5:7 at level 2 sharing row 8, Hide Detail at row 8 hides 3:7. At any other row of a group they act on the group of that row's level that contains it, through its summary row: row 6 hides 5:7, and row 4 hides 3:10. Hiding hides the detail and marks the summary row; showing shows every row of the detail, including rows hidden manually, except those of inner groups whose summary row keeps its marker, and removes the marker.
- They are rejected at a row in no group (`not_grouped`: row 12 above), at a group with no summary row (`not_grouped`), and where Excel's button would do nothing (`detail`): Hide Detail on a group already marked, or Show Detail on one that is not marked.
- `{ op: 'showLevels', sheet, rows: 2 }` (and `columns`) is a level button: it hides the rows at level 2 and deeper, shows the other grouped rows, including rows hidden manually among them, leaves rows without a level unchanged, and then sets the markers as Excel keeps them. With the rows above, level 2 hides 5:7 and marks row 8; level 1 hides 3:10 and marks only row 11; level 8 shows every group.

A rejection is `{ ok: false, code, message }`, with the message in the workbook's locale (`Детали строки 11 уже показаны` in ru-RU), and changes nothing.

### Without the engine

A grid that keeps its own rows and columns runs the same commands with `outlineEdit(sheet, op)` from `xlcalc/a1`: it takes the sheet's `rows`, `columns` and `outline` (a `SheetInput` or its layout) and the operation without `sheet`. It returns the full axes after the command, as a sheet holds them (an axis with no rows or columns is left out), or the same rejection `editStructure` returns, in English.

```js
import { outlineEdit } from 'xlcalc/a1'

const rows = { '3:4': { outlineLevel: 1 }, '5:7': { outlineLevel: 2 }, '8:10': { outlineLevel: 1 } }
console.log(outlineEdit({ rows }, { op: 'hideDetail', row: 8 }))
// → {
//     ok: true,
//     rows: {
//       '8': { outlineLevel: 1, collapsed: true },
//       '3:4': { outlineLevel: 1 },
//       '5:7': { outlineLevel: 2, hidden: true },
//       '9:10': { outlineLevel: 1 }
//     }
//   }
console.log(outlineEdit({ rows }, { op: 'showDetail', row: 11 }))
// → {
//     ok: false,
//     code: 'detail',
//     message: 'the detail of row 11 is shown already'
//   }
```

## In xlsx files

`writeXlsx` writes the markers as Excel keeps and saves them. When all the rows of a group are hidden manually (`Hidden`), Excel itself immediately marks the summary row and saves it with `collapsed="1"`; a file that contains such rows without the marker opens collapsed in Excel (its `ShowDetail` is False) and is saved with the marker. So rows hidden through `setLayout` are written with the marker too, and a marker over visible rows is left out. `readXlsx` keeps a file's markers as they are; levels, hidden rows and columns, and summary settings round-trip.

## Known differences

- Rows hidden and shown manually (`setLayout`): Excel re-evaluates the markers only for the groups whose rows change, so a summary row that is shown manually after being hidden along with its group stays without a marker, and showing the outer group later shows the inner one too. The engine marks every summary row whose group is entirely hidden, as Excel interprets a file it opens.
- Columns grouped in a span of several columns that reaches XFD, as a file's `<col max="16384">` stores them: Excel's level buttons leave them unchanged and its object model rejects Ungroup over any of them. The engine treats them like any other columns, as Excel treats XFD grouped on its own.

## Checked against Excel

Tested in Excel for Windows through COM, en-US and ru-RU, each set predicted before Excel ran it.
