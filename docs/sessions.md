# Sessions

A session keeps a computed book and recomputes only what an edit touches. It answers each change with what to draw again, keeps undo and redo, tells listeners of every change and takes its clock for the volatile functions.

## Opening a session

Mixed cell, name and formatting operations can be applied atomically with `applyChanges(operations, { baseRevision, clock })`. `previewChanges` returns bounded before/after value impacts, including dependent cells and errors, without changing the book or its history. Use the preview's revision and clock when applying it. `readCells`, `findCells`, `getDependents` and `summarizeRanges` read at one revision where the book lives, with projection and bounded pages instead of a full snapshot. `findCells` finds as Excel's Find with `wildcards`, `wholeCell`, `byColumns`, `skipHidden` and `in: 'edit'` (Look in Formulas) or `'text'` (Values), measured in Excel for Windows; `rows` keeps the rows a grid shows, `after` begins after a cell and goes round as Find Next does, `editText` gives the text a cell opens with, and the session keeps the texts it searched, so a search typed letter by letter costs a match of kept texts (0.02 s to 0.04 s over 510,000 cells). `replaceInCells` is Excel's Replace All over the same options, measured in Excel for Windows too, in one step of the undo. These methods also work in inline sessions and pools. See [live workbook operations and queries](live-books.md) for the operation shapes, clocks, cursors and limits.

A session holds a computed book and recomputes only what an edit changes. `setCells` applies edits atomically and answers the session's revision and the cells whose values changed; `explain` shows where a value comes from, as a tree and as text in the book's locale ([Localization](locale.md)).

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{
    name: 'Plan',
    cells: {
      A1: { value: 100 }, A2: { value: 250 },
      B1: { formula: 'SUM(A1:A2)' },
      B2: { formula: 'ROUND(B1*1.2,0)' },
    },
  }],
})

console.log(session.setCells([{ sheet: 'Plan', cell: 'A2', input: { value: 300 } }]))
// → {
//     revision: 1,
//     changes: [
//       { sheet: 'Plan', cell: 'A2', value: 300, text: '300' },
//       { sheet: 'Plan', cell: 'B1', value: 400, text: '400' },
//       { sheet: 'Plan', cell: 'B2', value: 480, text: '480' }
//     ]
//   }
console.log(session.explain('Plan', 'B2').text)
// → Plan!B2 = 480  ← ROUND(B1*1.2,0)
//     Plan!B1 = 400  ← SUM(A1:A2)
//       Plan!A1:A2 (2 cells)
console.log(session.getResult().hash === session.getResult().hash) // → true
session.close()
```

Every call that changes the book answers an `EditResult`: the session's `revision` after it (one more for each such call, `session.revision` reads it) and `changes`, the cells whose value or shown text changed, with a cell's `formula`, own `style`, own `numberFormat`, `note` or `link` when they changed (null when the cell has none now). Changes to `meta` are not listed: the host keeps this data, and history restores it. `getResult()` returns the same shape as `compute`, `getSnapshot()` independent copies of the book and its result at one moment in every kind of session (what `writeXlsx` takes), `getHash()` its values hash alone, `getBook()` the book with every change so far (the session's own objects in the same thread: do not modify them), and `close()` releases the book. Later edits leave a snapshot unchanged, and changing a snapshot cannot change the session. A book an answer gives has the type `Book`, a `BookInput`: it goes back into any call that takes a book. A full `compute` of the edited book with `session.clock` gives the same hash as `session.getResult()`; the test suite checks this after each of 200 random edits.

A session changes its book as a grid does, each kind of change in a section of its own: the cells and their look ([Editing a grid](edits.md)), typed text ([Cell entry](entry.md)), widths, heights and the sheet's other settings ([Layout](grid.md#layout)), rows, columns, ranges and sheets ([Structure operations](structure.md)), the names (`setNames`) and the style table (`setStyles`); it reads a window without the whole book ([A window for a grid](grid.md#a-window-for-a-grid)) and undoes every change ([Undo and redo](#undo-and-redo)).

## Undo and redo

A session undoes and redoes its changes, as Excel's Undo and Redo do:

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{ name: 'Plan', cells: { A1: { value: 100 }, A2: { value: 250 }, A3: { formula: 'SUM(A1:A2)' } } }],
})
session.setCells([{ sheet: 'Plan', cell: 'A1', input: { value: 120 } }])
session.editStructure({ op: 'insertRows', sheet: 'Plan', row: 2 })
console.log(session.history) // → { undo: 2, redo: 0 }
const back = session.undo()
console.log(back.call, back.mapCell({ sheet: 'Plan', cell: 'A4' }), back.changes.map((c) => c.formula)) // → editStructure { sheet: 'Plan', cell: 'A3' } [ 'SUM(A1:A2)' ]
console.log(session.undo())
// → {
//     call: 'setCells',
//     revision: 4,
//     changes: [
//       { sheet: 'Plan', cell: 'A1', value: 100, text: '100' },
//       { sheet: 'Plan', cell: 'A3', value: 350, text: '350' }
//     ],
//     cells: [ { sheet: 'Plan', cell: 'A1' } ]
//   }
console.log(session.undo(), session.history) // → null { undo: 0, redo: 2 }
console.log(session.redo().changes.length, session.getBook().sheets[0].cells.A1) // → 2 { value: 120 }
```

`undo()` takes back the last call that changed the book, whole however many cells it changed: `setCells`, `replaceCells`, `clearCells`, `patchCells`, `setCellText`, `setCellTexts` and `fillCellText` (with the widths they gave the columns), `pasteText`, `setNames`, `setStyles`, `setLayout` or `editStructure`. Its answer says which call it was and is that call's answer (`EditResult`), its revision one more: the cells whose value, text, formula or look changed, the `layout`, the `styles`, the `moved` addresses. For the calls that change cells it also lists the cells whose input came back (`cells`), whose note or link a host reads again. For a structure operation it has `mapCell`, where a cell of the book before the undo is now (`A4` above is `A3` again once the inserted row is gone): a host moves its values by it and applies the changes, as after `editStructure`, and reads `getLayout()`, `getUsedRange(sheet)` and its window again, as it does after a change of the layout, the names or the styles. `redo()` does the last undone step again; a new change forgets what `redo` could do. With nothing to take back or do again they answer null, and `history` counts the steps each way, for a toolbar's buttons. `setClock` and a live session's new clock are no steps, and neither is a call that fails: an undo stopped by a deadline or an abort changes nothing.

A session keeps 100 steps, as Excel does, in 64 MB at most by the engine's estimate, or in a pool with `maxHeapMb` its share of an eighth of a thread's heap (8 MB for 384 MB and `maxSessionsPerThread: 6`): the option `undo: { depth, maxMb }` sets both, and `depth: 0` keeps none. A step keeps what its call replaced and no more: the inputs of the cells it changed, the layout of the sheets it changed (the widths, `bestFit` and hidden columns of an AutoFit too), the style table or the names. A structure operation keeps the cells it does not carry back as they were (the rows it deleted, the formulas it rewrote) and the layout of the sheets it changed, not a copy of the book, and undoing it builds the book's state again, as the operation did. A deleted sheet comes back with its cells, which Excel cannot do; a step larger than `maxMb` alone is not kept and the steps before it go, as Excel forgets what it cannot undo. On the synthetic book of 34,000 rows, an edit and its undo cost what the edit costs (107 ms for an amount on the data sheet, which its 2,800 SUMIFS read), `insertRows` keeps 0.4 MB and takes 1.5 s where it takes 1.3 s without a history, its undo and its redo 1.4 to 1.5 s each; `deleteRows` takes 1.5 s where it takes 1.4 s. Deleting the sheet of 506,594 cells would keep 59 MB, over the default, so the session forgets its history then; with `maxMb: 100` it is undone in 1.4 to 1.7 s.

## Following changes

`session.subscribe(listener)` calls the listener after each call that changes the book, whoever makes it: an edit, typed text, a paste, a structure operation, a change of layout or styles, `undo` and `redo`. The listener gets what the call answered of `EditResult` (the revision, the changed cells, the sheets whose layout changed, the style table, how addresses moved, where to repaint) and the call's name in `method`, so a grid, a total shown elsewhere and a second window can all follow one session without passing the answers around:

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{ name: 'Plan', cells: { A1: { value: 100 }, A2: { value: 250 }, A3: { formula: 'SUM(A1:A2)' } } }],
})
const unsubscribe = session.subscribe((change) => {
  console.log(change.method, change.revision, change.changes.map((c) => `${c.cell}=${c.value}`).join(' '))
})
session.setCells([{ sheet: 'Plan', cell: 'A1', input: { value: 150 } }])
session.undo()
session.previewCellText('Plan', 'A2', '300')
unsubscribe()
session.setCells([{ sheet: 'Plan', cell: 'A2', input: { value: 300 } }])
console.log(session.revision)
// → setCells 1 A1=150 A3=400
// → undo 2 A1=100 A3=350
// → 3
```

- A preview, a read, a refusal and a call that throws are not told; the listener runs when the call is done, before it returns, so the session it reads is the one after the change.
- Listeners run in the order they subscribed, and changes reach them in the order of their revisions: a change a listener makes reaches every listener after the one it answers.
- A listener that throws stops neither the call nor the other listeners; its error is thrown again in a microtask.
- `subscribe` answers the function that unsubscribes; `close()` unsubscribes every listener. An asynchronous session (`createInlineCalc`, `createCalcPool`) takes `subscribe` the same way, synchronously: a listener runs when a call's answer has come, before the call's promise settles.

## The clock and volatile functions

```js
import { compute } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { formula: 'TODAY()' }, A2: { formula: 'NOW()' }, A3: { formula: 'RANDBETWEEN(1,6)' } } }] }

const fixed = compute(book, { clock: { now: '2026-01-31T18:00:00Z', seed: 42 } })
console.log(fixed.sheets[0].cells.A1.value, fixed.sheets[0].cells.A2.value) // → 46053 46053.75

const now = compute(book) // the current moment and a random seed
console.log(compute(book, { clock: now.clock }).hash === now.hash) // → true
console.log(now.volatile.map((c) => c.cell).join(',')) // → A1,A2,A3
```

`volatile` lists the cells that depend on volatile functions, directly or through other cells, and on the clock's year through date text without a year ([Localization](locale.md)). A session opened without a clock keeps the moment it was opened until `setClock`; `createSession(book, { live: true })` takes a fresh clock on every recomputation instead.
