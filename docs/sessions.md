# Sessions

A session keeps a calculated workbook and recalculates only what an edit affects. It returns, for each change, what needs to be redrawn, keeps an undo and redo history, notifies listeners of every change and holds the clock for the volatile functions.

## Opening a session

Mixed cell, name and formatting operations can be applied atomically with `applyChanges(operations, { baseRevision, clock })`. `previewChanges` returns bounded before/after impacts on values, including dependent cells and errors, without changing the workbook or its history. Use the preview's revision and clock when applying it. `readCells`, `findCells`, `getDependents` and `summarizeRanges` read at a single revision where the workbook lives, with projection and bounded pages instead of a full snapshot.

`findCells` searches like Excel's Find, with `wildcards`, `wholeCell`, `byColumns`, `skipHidden` and `in: 'edit'` (Look in: Formulas) or `'text'` (Look in: Values), as verified in Excel for Windows. `rows` restricts the search to the rows a grid displays, `after` starts after a given cell and wraps around as Find Next does, and `editText` returns the text a cell opens with in edit mode. The session caches the texts it searched, so a search typed letter by letter only costs a match over cached texts (0.02 s to 0.04 s over 510,000 cells). `replaceInCells` is Excel's Replace All with the same options, also verified in Excel for Windows, in a single undo step. These methods work in inline sessions and pools too. See [live workbook operations and queries](live-books.md) for the shapes of operations, clocks, cursors and limits.

A session holds a calculated workbook and recalculates only what an edit changes. `setCells` applies edits atomically and returns the session's revision and the cells whose values changed; `explain` shows where a value comes from, as a tree and as text in the workbook's locale ([Localization](locale.md)).

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

Every call that changes the workbook returns an `EditResult`: the session's `revision` after the call (incremented by one for each such call; `session.revision` reads it) and `changes`, the cells whose value or displayed text changed, with a cell's `formula`, own `style`, own `numberFormat`, `note` or `link` when they changed (null when the cell no longer has one). Changes to `meta` are not listed: the host owns this data, and the history restores it.

- `createSession(book, opts)` takes the workbook's cell objects and the options as given, not copies of them, in every kind of session: they belong to the session from then on, and changing one in place afterwards gives results that depend on the kind of session. Change the workbook through the session's calls instead.
- `getResult()` returns the same shape as `compute`.
- `getSnapshot()` returns independent copies of the workbook and its result at a single point in time, in every kind of session (what `writeXlsx` accepts). Later edits do not change a snapshot, and changing a snapshot cannot change the session.
- `getHash()` returns only the values hash.
- `getBook()` returns the workbook with every change so far (in the same thread, these are the session's own objects: do not modify them).
- `close()` releases the workbook.

A workbook returned by a call has the type `Book`, which is a `BookInput`: it can be passed back to any call that accepts a workbook. A full `compute` of the edited workbook with `session.clock` produces the same hash as `session.getResult()`; the test suite checks this after each of 200 random edits.

A session changes its workbook the way a grid does, and each kind of change has its own section: cells and their formatting ([Editing a grid](edits.md)), typed text ([Cell entry](entry.md)), widths, heights and the sheet's other settings ([Layout](grid.md#layout)), rows, columns, ranges and sheets ([Structure operations](structure.md)), the names (`setNames`) and the style table (`setStyles`). It reads a viewport without the whole workbook ([A viewport for a grid](grid.md#a-viewport-for-a-grid)) and can undo every change ([Undo and redo](#undo-and-redo)).

## Compiling and recalculating

By default a formula is compiled to a tree of JavaScript closures without generating code, so it works under a strict Content Security Policy. With `execution: 'eval'` it is compiled to JavaScript source, and a single `new Function` call compiles the whole workbook. A formula that exceeds the JS compiler's nesting limit falls back to closures, preserving evaluation order. Both modes produce the same values and hashes, and the tests calculate every test workbook both ways. Cells are evaluated in topological order of the reference graph. A session keeps the compiled workbook and recalculates only the cells that depend on an edit. On a synthetic workbook with 3,400 data rows and 3,907 formulas, mostly SUMIFS over entire columns, a cold calculation takes about 0.3 s including the cells' displayed text (0.28 s without it), and an edit of a criteria cell takes under 10 ms (p95) in either mode on an Apple M5 Pro with Node 26. `test/budgets.test.ts` enforces limits of 1.5 s and 100 ms.

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

`undo()` reverts the last call that changed the workbook, as a whole, however many cells it changed: `setCells`, `replaceCells`, `clearCells`, `patchCells`, `setCellText`, `setCellTexts` and `fillCellText` (including the column widths they set), `pasteText`, `setNames`, `setStyles`, `setLayout` or `editStructure`. It returns which call it reverted, in the form of that call's result (`EditResult`) with the revision incremented: the cells whose value, text, formula or formatting changed, the `layout`, the `styles` and the `moved` addresses. For calls that change cells it also lists the cells whose input was restored (`cells`), so a host can read their notes and hyperlinks again. For a structure operation it has `mapCell`, which tells where a cell of the workbook before the undo is now (`A4` above is `A3` again once the inserted row is gone): a host moves its own data with it and applies the changes, as after `editStructure`, and reads `getLayout()`, `getUsedRange(sheet)` and its viewport again, as it does after a change of the layout, the names or the styles.

`redo()` repeats the last undone step; a new change clears the redo history. With nothing to undo or redo they return null, and `history` counts the steps in each direction, for a toolbar's buttons. `setClock` and a live session's new clock are not steps, and neither is a call that fails: an undo stopped by a deadline or an abort changes nothing.

A session keeps 100 steps, as Excel does, in at most 64 MB by the engine's estimate, or in a pool with `maxHeapMb`, in its share of an eighth of a thread's heap (8 MB for 384 MB and `maxSessionsPerThread: 6`). The option `undo: { depth, maxMb }` sets both, and `depth: 0` disables the history. A step keeps only what its call replaced: the inputs of the cells it changed, the layout of the sheets it changed (including the widths, `bestFit` and hidden columns of an AutoFit), the style table or the names. A structure operation keeps the cells it cannot restore by itself as they were (the rows it deleted, the formulas it rewrote) and the layout of the sheets it changed, not a copy of the workbook, and undoing it rebuilds the workbook's state the way the operation built it. A deleted sheet comes back with its cells, which Excel cannot do. A step larger than `maxMb` on its own is not kept, and the steps before it are dropped, just as Excel clears what it cannot undo.

On the synthetic workbook of 34,000 rows, an edit and its undo cost what the edit costs (107 ms for an amount on the data sheet, which its 2,800 SUMIFS read). `insertRows` keeps 0.4 MB and takes 1.5 s, against 1.3 s without a history, and its undo and redo take 1.4 to 1.5 s each; `deleteRows` takes 1.5 s against 1.4 s. Deleting the sheet of 506,594 cells would keep 59 MB, over the default, so the session clears its history in that case; with `maxMb: 100` the deletion is undone in 1.4 to 1.7 s.

## Tracking changes

`session.subscribe(listener)` calls the listener after each call that changes the workbook, whoever makes it: an edit, typed text, a paste, a structure operation, a change of layout or styles, `undo` and `redo`. The listener receives the call's `EditResult` (the revision, the changed cells, the sheets whose layout changed, the style table, how addresses moved, where to repaint) and the call's name in `method`, so a grid, a total displayed elsewhere and a second window can all follow one session without passing results around:

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

- Previews, reads, rejected calls and calls that throw are not reported. The listener runs when the call is done, before it returns, so the session it reads is already the one after the change.
- Listeners run in the order they subscribed, and changes reach them in the order of their revisions: a change a listener makes reaches every listener after the change it is responding to.
- A listener that throws stops neither the call nor the other listeners; its error is rethrown in a microtask.
- `subscribe` returns the function that unsubscribes; `close()` unsubscribes every listener. An asynchronous session (`createInlineCalc`, `createCalcPool`) accepts `subscribe` the same way, synchronously: a listener runs when a call's result has arrived, before the call's promise settles.

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

`volatile` lists the cells that depend on volatile functions, directly or through other cells, and on the clock's year through date text without a year ([Localization](locale.md)). A session created without a clock keeps the moment it was created until `setClock` is called; `createSession(book, { live: true })` takes a fresh clock on every recalculation instead.
