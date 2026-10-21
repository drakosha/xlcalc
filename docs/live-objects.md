# Live sheet and cell objects

`liveBook(session)` returns objects for a session's sheets and cells. Reading a cell's property reads what the session calculated; assigning to one is an edit of the session, recalculated before the next line runs and recorded as an undo step; and a listener on a cell, a sheet or the workbook is called whenever a change affects it, whoever made the change. The objects keep no copy of the workbook and depend on nothing outside the package, so a framework's reactivity (Vue's `customRef`, a signal) can be built on top of them.

```js
import { createSession, liveBook } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'Plan', cells: { A1: { value: 100 } } }] })
const book = liveBook(session)
const plan = book.sheet('Plan')
plan.cells.B1.formula = 'A1*1.2'
plan.cells.A1.value = 200
console.log(plan.cells.B1.value, plan.cells.B1.text, session.history.undo) // → 240 240 2
```

## Reading and assigning

A cell is `sheet.cells.A1` or `sheet.cell('A1')`, by its address in any case and without `$`. Its properties read the session's current state through `getCell`; nothing is calculated outside the session:

- `value`: the calculated value, null for an empty cell.
- `text`: the text the cell displays; undefined for an empty cell, and in a session created with `display: false`.
- `formula`: the formula without `=`, in the session's canonical form (sheets quoted, functions in upper case).
- `numberFormat` and `style`: the cell's own, as `getCell` returns them, not those it displays from its row or column.

An assignment is a call to the session:

- `value = x` is `setCells` with a value: the contents are replaced and a formula is removed, while the formatting stays, as when a value is typed over a cell in Excel. `value = null` clears the contents and the hyperlink, like `clearCells` with `contents`.
- `formula = 'A1*2'` is `setCells` with a formula; `formula = null` removes it.
- `numberFormat` and `style` are `patchCells`: the value and formula stay. A style given by its fields is applied on top of the style the cell displays ([Editing a grid](edits.md)).
- `cell.type('5%')` is text typed into the cell, parsed as Excel parses it in the workbook's locale (`setCellText`), and returns what `setCellText` returns.

An assignment the session rejects throws the session's own error, and the workbook stays unchanged: part of a legacy array formula, a value of the wrong shape, an edit beyond the session's `limits`. The objects always edit the workbook in its current state: they take no `baseRevision`.

```js
import { createSession, liveBook } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'Plan', cells: { A1: { value: 0.25, numberFormat: '0%' }, A2: { formula: 'A1*4' } } }] })
const a = liveBook(session).sheet('Plan').cells
a.A1.value = 0.5
console.log(a.A1.text, a.A2.value) // → 50% 2
a.A2.value = 'total'
console.log(a.A2.formula, a.A2.value) // → undefined total
a.A1.style = { font: { bold: true } }
console.log(a.A1.style, session.getLayout().styles) // → 0 [ { font: { bold: true } } ]
console.log(a.B1.type('5%').ok, a.B1.value, a.B1.text) // → true 0.05 5%
```

## Batches

`book.batch(fn)` applies the assignments `fn` makes as a single `applyChanges` ([Batch changes, previews and paged queries in a live session](live-books.md)): one recalculation, one revision, one undo step and one change reported to listeners.

- Until `fn` returns, cells read as they were before the batch: the assignments are deferred to its end.
- A batch inside another batch is part of the outer one.
- When `fn` throws, none of its assignments are applied and the error propagates; an inner batch that throws discards only its own assignments.
- Typed text (`type`) cannot be part of a batch, since `applyChanges` does not accept typed text, and `fn` must not return a promise: an assignment after an `await` would fall outside the batch. Both throw `InputError`.

```js
import { createSession, liveBook } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'Plan', cells: { A3: { formula: 'A1+A2' } } }] })
const book = liveBook(session)
const plan = book.sheet('Plan')
book.batch(() => {
  plan.cells.A1.value = 1
  plan.cells.A2.value = 2
  console.log(plan.cells.A3.value) // → 0
})
console.log(plan.cells.A3.value, session.revision, session.history.undo) // → 3 1 1
```

## Tracking changes

`cell.watch(listener)`, `sheet.watch(listener)` and `book.watch(listener)` call the listener with the session's `SessionChange` ([Sessions](sessions.md#tracking-changes)) and return the function that unsubscribes. They receive every change of the workbook, whether made through the objects or not: `session.setCells`, typed text, a structure operation, `undo` and `redo`.

- A cell's listener is called when the change affected the cell: its value, displayed text, formula or own formatting changed (`changes`, which also lists the cells a recalculation changed, but not those of a rectangle of formatting returned as `repaint`), or a structure operation may have moved another cell to its address (`moved`). A conditional formatting rule that changes how the cell is displayed (`repaint`) does not call it: the objects do not read a cell's conditional formatting.
- A sheet's listener is called for a change of its cells or its layout, a structure operation on it, its rename or move, cells to repaint on it, and a change of the workbook's style table.
- The workbook's listener is called for every change.
- A listener passed to several `watch` calls is called once per change; each `watch` unsubscribes separately. A listener that throws stops neither the edit nor the other listeners, and its error is rethrown in a microtask, as with `subscribe`.

The workbook object subscribes to its session once and finds a cell's listeners by address, so an edit costs as much as the cells it changed, not the number of cells watched. A structure operation notifies the watched cells of its sheet from the inserted or deleted row or column onward, and of both ranges of a move, even cells that were empty before and after.

```js
import { createSession, liveBook } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'Plan', cells: { A1: { value: 1 }, A2: { formula: 'A1*2' }, C1: { value: 'x' } } }] })
const plan = liveBook(session).sheet('Plan')
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

A cell object stands for its address: with `const c = plan.cells.A5` and a row then inserted above it, `c` is the cell now at A5, and its listener is called because what is there changed. A host that wants to follow a cell as it moves maps it with the `mapCell` of `editStructure`'s result or with `moved`. The same address returns the same object while the object is held; objects that are not held are garbage-collected, and a listener stays attached to its address regardless.

A sheet object follows its sheet through a rename or a move: `sheet.name` is the current name, and `book.sheet` accepts the new name, in any case. After the sheet is deleted, every use of its object and of its cells throws `InputError`, and their listeners are no longer called; a sheet restored by `undo` is a new object, obtained from `book.sheet`.

```js
import { createSession, liveBook } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'Plan', cells: { A5: { value: 5 } } }] })
const book = liveBook(session)
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

`liveBook` accepts a synchronous session (`createSession`) and returns the same object for the same session each time. An asynchronous session (`createInlineCalc`, `createCalcPool`) cannot deliver the recalculated value of an assignment before the next line runs, so `liveBook` throws `InputError` for it; its objects would need assignments that return promises.

## With Vue

The objects do not depend on Vue. A `customRef` tracks a cell's reads and triggers on its `watch`; the cell's value itself stays in the session:

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

`<input v-model.number="a1">` with `const a1 = useCell(plan.cells.A1)` then writes into the session, and a `computed` over another cell's ref follows the recalculation.

## How it is tested

Not against Excel: the objects are this library's own layer over a session, and the session's calls are what is checked against Excel. The tests (`test/live-objects.test.ts`, in both execution modes) read cells after assignments and batches, compare the error and the workbook's hash of a rejected assignment with those of the session's call, count the undo steps and reported changes, and check which listeners each kind of change calls: through the objects, through the session, `undo`, `redo`, inserted rows and columns, a move, and a sheet's rename and deletion. The Vue example is not run.
