# Outlines

Rows and columns of a sheet can be grouped, as Excel's Data > Group does: the bars beside the row numbers and above the column letters, a `−` or `+` button on each group's summary line and the level buttons in the corner. The library gives a sheet's groups as Excel draws them, `outlineGroups`, without the engine, and Excel's commands on them, Group, Ungroup, Show Detail, Hide Detail and the level buttons, as structure operations: `editStructure(book, op)`, or `session.editStructure(op)` with one revision and one step of the undo, and on a sheet alone, without the engine, `outlineEdit`. A grid draws the groups from `getLayout()` and sends the commands; the rules below are Excel for Windows' own, measured through its object model.

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

A row's outline level (`RowInput.outlineLevel`, 1 to 7) says how deep it is grouped; a column's likewise. `outlineGroups(sheet)` takes a sheet or its layout and gives each axis's groups, each `{ first, last, level, summary, collapsed }`, columns by number:

- A group is the longest run of lines at its level or deeper. Rows 3:4 and 8:10 at level 1 with 5:7 at level 2 make the group 3:10 of level 1 and 5:7 of level 2. Two runs side by side are one group: rows 3:5 grouped and then rows 6:8 make one group 3:8; a line without a level between them keeps them apart.
- The summary line, which holds the group's button, is the line past the group: row 11 for 3:10, below by default, and with summaries above (`SheetInput.outline`, `summaryBelow: false`) the line before it, row 2. Groups that end together share their summary line. A group on the sheet's edge, past which its summary line would be, has none (`summary: null`), and Excel draws no button for it.
- `collapsed` is the mark Excel keeps on the summary line (below); a group without a summary line is collapsed when all its lines are hidden.

The groups come in order of their first line, the outer before the inner where two begin together.

## The marks Excel keeps

Excel marks a summary row collapsed (xlsx `collapsed`, `RowInput.collapsed`) whenever all the rows of its detail, the deeper rows next to it, are hidden, however they got hidden: by its button, by a level button, or by hand. It drops the mark as soon as one of them shows. Rows 2:4 hidden through `setLayout`, with no mark on row 5, are a collapsed group: Excel opens such a file with the group collapsed and saves it with the mark.

One thing more keeps an inner group closed: a summary row inside a group whose rows are all hidden keeps the mark it had. Hide 5:7 (row 8 marked), then 3:10 (row 11 marked, row 8 hidden with its mark), then show 3:10: rows 3, 4, 8, 9 and 10 show, and 5:7 stay hidden under the mark on row 8, as in the example above. Hide 3:10 first, and row 8 gets no mark: showing 3:10 shows all of it.

A group on the sheet's edge keeps its mark on its line at the edge: row 1 with summaries above, row 1,048,576, column A with summaries left.

`outlineGroups`, the commands and `writeXlsx` read the marks so, whatever marks a book holds; `setLayout` sets the fields as given, and `readXlsx` reads a file's marks as they are.

## The commands

Each is a `StructureOp` on one sheet; a session's `editStructure` recomputes what `SUBTOTAL` and `AGGREGATE` count over the rows hidden or shown, and nothing else. No cell moves: `mapCell` gives every cell its own place.

- `{ op: 'group', sheet, rows: '5:7' }` (or `columns: 'B:D'`, one line `'5'` or `'C'`) puts each line one level deeper. A line at level 7, the deepest, stays there, and Group is refused (`outline_deep`) when all the lines are at 7. Grouping hides and shows nothing.
- `{ op: 'ungroup', sheet, rows }` takes one level off each line that has one, refused (`not_grouped`) when none has. Hidden lines stay hidden: a collapsed group ungrouped leaves its rows hidden.
- `{ op: 'hideDetail', sheet, row: 8 }` and `showDetail` (or `column: 'H'`) are a group's buttons, at one line, as Excel's object model takes them. At a summary line they act on its whole detail: with rows 3:7 at level 1 and 5:7 at level 2 sharing row 8, Hide Detail at row 8 hides 3:7. At another line of a group they act on the group of that line's level that holds it, through its summary line: row 6 hides 5:7, row 4 hides 3:10. Hiding hides the detail and marks the summary line; showing shows every row of the detail, rows hidden by hand too, but those of the groups inside whose summary line keeps its mark, and drops the mark.
- They are refused at a line in no group (`not_grouped`: row 12 above), at a group with no summary line (`not_grouped`), and where Excel's button would not act (`detail`): Hide Detail of a group already marked, Show Detail of one not marked.
- `{ op: 'showLevels', sheet, rows: 2 }` (and `columns`) is a level button: the lines of level 2 and deeper hidden, the other grouped lines shown, rows hidden by hand among them too, lines without a level left as they are; then the marks as Excel keeps them. With the rows above, level 2 hides 5:7 and marks row 8; level 1 hides 3:10 and marks row 11 alone; 8 shows every group.

A refusal is `{ ok: false, code, message }`, the message in the book's locale (`Детали строки 11 уже показаны` in ru-RU), and changes nothing.

### Without the engine

A grid that keeps its own rows and columns runs the same commands with `outlineEdit(sheet, op)` of `xlcalc/a1`: the sheet's `rows`, `columns` and `outline` (a `SheetInput` or its layout) and the operation without `sheet`. It answers the axes after the command, whole, as a sheet holds them (an axis with no lines left out), or the refusal `editStructure` gives, in English.

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

`writeXlsx` writes the marks as Excel keeps and saves them. Excel itself, hiding every row of a group by hand (`Hidden`), marks its summary row at once and saves it with `collapsed="1"`; a file that holds such rows without the mark opens collapsed in Excel (its `ShowDetail` is False) and is saved with the mark. So rows hidden through `setLayout` are written with it too, and a mark over rows shown is left out. `readXlsx` keeps a file's marks as they are; the levels, hidden lines and summary settings go both ways.

## Known differences

- Rows hidden and shown by hand (`setLayout`): Excel reads the marks again only for the groups whose rows change, so a summary row shown by hand over a group it was hidden with stays without a mark there, and showing the outer group later shows the inner one too. The engine marks every summary row whose group is all hidden, as Excel reads a file it opens.
- Columns grouped in a span of several columns that reaches XFD, as a file's `<col max="16384">` holds them: Excel's level buttons leave them as they are and its object model refuses Ungroup over any of them. The engine takes them as any other columns, as Excel takes XFD grouped alone.

## Checked against Excel

Measured in Excel for Windows through COM, en-US and ru-RU, each set predicted before Excel ran it.
