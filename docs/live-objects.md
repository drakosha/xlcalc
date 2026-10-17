# Live sheet and cell objects

`liveWorkbook(session)` gives objects of a session's sheets and cells. Reading a cell's property reads what the session computed; assigning one is the session's edit, recomputed before the next line runs, with a step of the undo; and a listener of a cell, a sheet or the book is called when a change touched it, whoever made the change. The objects hold no copy of the book and need nothing besides the package, so a framework's reactivity (Vue's `customRef`, a signal) can sit on top of them.

```js
import { createSession, liveWorkbook } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'Plan', cells: { A1: { value: 100 } } }] })
const book = liveWorkbook(session)
const plan = book.sheet('Plan')
plan.cells.B1.formula = 'A1*1.2'
plan.cells.A1.value = 200
console.log(plan.cells.B1.value, plan.cells.B1.text, session.history.undo) // → 240 240 2
```

## Reading and assigning

A cell is `sheet.cells.A1` or `sheet.cell('A1')`, by its address in any case and without `$`. Its properties read the session as it is now, through `getCell`; nothing is computed apart from the session:

- `value`: the computed value, null for an empty cell.
- `text`: the text the cell shows; undefined for an empty cell, and in a session opened with `display: false`.
- `formula`: the formula without `=`, as the session spells it (sheets quoted, functions in upper case).
- `numberFormat` and `style`: the cell's own, as `getCell` gives them; not those it shows from its row or column.

An assignment is the session's call:

- `value = x` is `setCells` with a value: the content is replaced and a formula goes, the look stays, as when a value is typed over a cell in Excel. `value = null` clears the content and the hyperlink, as `clearCells` with `contents`.
- `formula = 'A1*2'` is `setCells` with a formula; `formula = null` removes it.
- `numberFormat` and `style` are `patchCells`: the value and formula stay. A style by its fields goes over the style the cell shows ([Editing a grid](edits.md)).
- `cell.type('5%')` is text typed into the cell, read as Excel reads it in the book's locale (`setCellText`), and answers what `setCellText` answers.

What the session refuses throws its own error from the assignment, and the book stays as it was: part of an array of the old kind, a value of the wrong shape, an edit over the session's `limits`. The objects always edit the book as it is now: they take no `baseRevision`.

```js
import { createSession, liveWorkbook } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'Plan', cells: { A1: { value: 0.25, numberFormat: '0%' }, A2: { formula: 'A1*4' } } }] })
const a = liveWorkbook(session).sheet('Plan').cells
a.A1.value = 0.5
console.log(a.A1.text, a.A2.value) // → 50% 2
a.A2.value = 'total'
console.log(a.A2.formula, a.A2.value) // → undefined total
a.A1.style = { font: { bold: true } }
console.log(a.A1.style, session.getLayout().styles) // → 0 [ { font: { bold: true } } ]
console.log(a.B1.type('5%').ok, a.B1.value, a.B1.text) // → true 0.05 5%
```

## Batches

`book.batch(fn)` applies the assignments `fn` makes as one `applyChanges` ([Batch changes, previews and paged queries in a live session](live-books.md)): one recomputation, one revision, one step of the undo and one change told to listeners.

- Until `fn` returns, the cells read as they were before the batch: the assignments wait for its end.
- A batch inside another is part of the outer one.
- When `fn` throws, none of its assignments is applied and the error goes on; an inner batch that throws loses only its own.
- Typed text (`type`) cannot be part of a batch, since `applyChanges` takes no typed text, and `fn` must not answer a promise: an assignment after an `await` would be outside the batch. Both throw `InputError`.

```js
import { createSession, liveWorkbook } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'Plan', cells: { A3: { formula: 'A1+A2' } } }] })
const book = liveWorkbook(session)
const plan = book.sheet('Plan')
book.batch(() => {
  plan.cells.A1.value = 1
  plan.cells.A2.value = 2
  console.log(plan.cells.A3.value) // → 0
})
console.log(plan.cells.A3.value, session.revision, session.history.undo) // → 3 1 1
```

## Following changes

`cell.watch(listener)`, `sheet.watch(listener)` and `book.watch(listener)` call the listener with the session's `SessionChange` ([Sessions](sessions.md), "Following changes") and answer the function that unsubscribes. They hear every change of the book, made through the objects or past them: `session.setCells`, typed text, a structure operation, `undo` and `redo`.

- A cell's listener is called when the change touched the cell: its value, shown text, formula or own look changed (`changes`, which lists the cells a recomputation changed too), or a structure operation may have brought another cell to its address (`moved`). A rule of conditional formatting that paints the cell otherwise (`repaint`) does not call it: the objects do not read a cell's conditional look.
- A sheet's listener is called for a change of its cells, its layout, a structure operation on it, its rename or move, the cells to repaint on it, and a change of the book's style table.
- The book's listener is called for every change.
- A listener given to several watches is called once a change; each watch unsubscribes alone. A listener that throws stops neither the edit nor the other listeners, and its error is thrown again in a microtask, as `subscribe` does.

The workbook subscribes to its session once and finds a cell's listeners by its address, so an edit costs the cells it changed, not the number of cells watched. A structure operation wakes the watched cells of its sheet from the inserted or deleted line on, and of both areas of a move, even those empty before and after.

```js
import { createSession, liveWorkbook } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'Plan', cells: { A1: { value: 1 }, A2: { formula: 'A1*2' }, C1: { value: 'x' } } }] })
const plan = liveWorkbook(session).sheet('Plan')
const stop = plan.cells.A2.watch((change) => console.log(change.method, plan.cells.A2.value))
plan.cells.C1.watch(() => console.log('C1'))
plan.cells.A1.value = 5
session.setCells([{ sheet: 'Plan', cell: 'A1', input: { value: 6 } }])
session.undo()
stop()
plan.cells.A1.value = 7
// → setCells 10
// → setCells 12
// → undo 10
```

## Cells keep their address, sheets follow their name

A cell object is its address: `const c = plan.cells.A5`, then a row inserted above it, and `c` is the cell now at A5, its listener called because what is there changed. A host that follows a cell as it moves maps it with `mapCell` of `editStructure`'s answer or with `moved`. The same address gives the same object while it is held; objects not held are collected, and a listener stays with its address all the same.

A sheet object follows its sheet through a rename or a move: `sheet.name` is the name now, and `book.sheet` takes the new name, in any case. After the sheet is deleted, every use of its object and of its cells throws `InputError`, and their listeners are called no more; a sheet brought back by `undo` is a new object, from `book.sheet`.

```js
import { createSession, liveWorkbook } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'Plan', cells: { A5: { value: 5 } } }] })
const book = liveWorkbook(session)
const plan = book.sheet('Plan')
const a5 = plan.cells.A5
a5.watch(() => console.log('A5 changed'))
session.editStructure({ op: 'insertRows', sheet: 'Plan', row: 2 })
console.log(a5.address, a5.value, plan.cells.A6.value)
session.editStructure({ op: 'renameSheet', sheet: 'Plan', name: 'Budget' })
console.log(plan.name, book.sheet('budget') === plan)
// → A5 changed
// → A5 null 5
// → Budget true
```

## Which sessions

`liveWorkbook` takes a synchronous session (`createSession`) and gives the same object for the same session again. An asynchronous session (`createInlineCalc`, `createCalcPool`) cannot answer an assignment with the recomputed value before the next line, so it throws `InputError`; its objects would need writes that answer promises.

## With Vue

The objects know nothing of Vue. A `customRef` tracks a cell's reads and triggers on its `watch`; the cell's own value stays the session's:

```ts
import { customRef, onScopeDispose } from 'vue'
import type { LiveCell } from 'xlcalc'

export function useCell(cell: LiveCell) {
  return customRef((track, trigger) => {
    const stop = cell.watch(trigger)
    onScopeDispose(stop)
    return {
      get: () => (track(), cell.value),
      set: (value) => (cell.value = value),
    }
  })
}
```

`<input v-model.number="a1">` with `const a1 = useCell(plan.cells.A1)` then writes into the session, and a `computed` over another cell's ref follows the recomputation.

## How it is tested

Not against Excel: the objects are a rule of this library over a session, and the session's own calls are what is checked against Excel. The tests (`test/live-objects.test.ts`, both modes of execution) read after assignments and batches, compare a refused assignment's error and the book's hash with the session's call, count the steps of the undo and the changes told, and check which listeners each kind of change calls: through the objects, through the session, `undo`, `redo`, inserted rows and columns, a move, a rename and a deletion of a sheet. The Vue example is not run.
