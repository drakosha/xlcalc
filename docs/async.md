# Asynchronous and threaded computation

A server that holds workbooks for many users, or a page that must stay responsive, cannot let one large workbook block its thread. The engine calculates the same way in three places, with the same values and hashes:

- synchronously, as `compute` and `createSession` always have;
- on the same thread, in slices: `createInlineCalc` yields to the event loop between formulas, and a call stops at its soft deadline or when its `AbortSignal` fires;
- in a thread pool: `createCalcPool` from `xlcalc/worker` keeps each session's workbook in a thread (Node `worker_threads`, or Web Workers in a browser), terminates a thread that runs past a hard deadline, bounds a thread's heap, and restores a session after its thread is lost.

Both calculators have the same API: `compute(book, opts)`, `createSession(book, opts)` and `close()`, all returning promises. A session has the methods of `Session`, returning promises.

```js
import { createInlineCalc } from 'xlcalc'
import { createCalcPool } from 'xlcalc/worker'

const book = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*10' } } }] }
const calc = createInlineCalc() // the same thread
const pool = createCalcPool({ threads: 2, hardTimeoutMs: 10_000, maxHeapMb: 512 })

const result = await calc.compute(book, { timeoutMs: 2_000 })
const session = await pool.createSession(book)
const { changes } = await session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }], { signal: AbortSignal.timeout(5_000) })
console.log(result.sheets[0].cells.A2.value, changes.map((c) => [c.cell, c.value]))
// → 20 [ [ 'A1', 3 ], [ 'A2', 30 ] ]
await Promise.all([calc.close(), pool.close()])
```

## In short

`createInlineCalc()` calculates on the same thread, in slices that let timers and input run between them; `createCalcPool()` from `xlcalc/worker` calculates in a thread pool, where a session's workbook lives in its thread. Both have the same API: `compute`, `createSession` and the session's methods return promises, and every call accepts `timeoutMs`, a soft deadline for calculation time, and `signal`, an `AbortSignal`. A call that fails changes nothing. `limits` bound the formulas of a workbook, the values of one array, the cells of one spill and the cells one session edit writes (`writes`: copying one cell onto an entire column writes 1,048,576), in the synchronous `compute` and `createSession` too; beyond them the call throws `CalcError` rather than calculating or writing.

```js
import { createInlineCalc } from 'xlcalc'
import { createCalcPool } from 'xlcalc/worker'

const book = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*10' }, A3: { formula: 'SUM(SEQUENCE(1000,1000))' } } }] }

const calc = createInlineCalc({ limits: { cells: 100_000 } })
const e = await calc.compute(book).catch((e) => e)
console.log(e.code, e.limit, e.cell) // → limit cells { sheet: 'S', cell: 'A3' }

const pool = createCalcPool({ threads: 2, hardTimeoutMs: 10_000 })
const session = await pool.createSession(book, { clock: { now: '2026-10-07T12:00:00Z', seed: 1 } })
console.log((await session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }], { timeoutMs: 1_000 })).changes)
// → [
//     { sheet: 'S', cell: 'A1', value: 3, text: '3' },
//     { sheet: 'S', cell: 'A2', value: 30, text: '30' }
//   ]
const stop = new AbortController()
const first = session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 4 } }])
const second = session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 5 } }], { signal: stop.signal }).catch((e) => e)
stop.abort() // the second call waits for the first: it is refused at once
await first
console.log((await second).code, (await session.getResult()).sheets[0].cells.A2.value) // → aborted 40
const bounded = await pool.createSession(book, { limits: { writes: 100_000 } })
const w = await bounded.editStructure({ op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', range: 'B:B' } }).catch((e) => e)
console.log(w.code, w.limit, w.size, bounded.revision) // → limit writes 1048576 0
await pool.close()
```

A task that runs past the pool's `hardTimeoutMs` terminates its thread, and in Node `maxHeapMb` bounds each thread's heap. The pool keeps a snapshot and a journal of each session's workbook, so a session whose thread was lost is restored on its next call to the state before the call that failed, including its undo and redo history. With `maxHeapMb`, a session's history holds `maxHeapMb / (8 × maxSessionsPerThread)` MB by default, so the histories of a full thread take an eighth of its heap; a session's own `undo.maxMb` overrides this. [docs/async.md](async.md) has the rules, the error codes and the costs.

## Calls

A session processes its calls in order: each starts when the previous one has finished. Two `setCells` calls made without waiting each get their own changes, as two synchronous calls would. A call that fails changes nothing: the session stays as it was before the call, and the next call works on that state.

`editStructure` returns `{ ok, mapCell, changes }` without the workbook, as a synchronous session does, so a thread does not send the workbook back on every operation; `getBook()` fetches it. From a thread, `getBook()`, `getResult()`, `explain()` and `getViewport()` return copies. `getSnapshot()` returns independent copies of the workbook and its result in one call in every kind of session: no edit comes between them, later edits leave both unchanged, and modifying the snapshot cannot change the session. On the same thread, its workbook is copied in slices that yield to the event loop; a worker's transport already copies it. `getBook()` and `getLayout()` on the same thread still expose the session's objects, which must not be modified, and so does a cell's `meta` in `getCells()`, `getCell()`, `getViewport()` and `getResult()`; a cell's note and link there are copies.

`explain(sheet, cell, opts?)` accepts its tree options together with `signal` and `timeoutMs` in both asynchronous calculators. A signal that is already aborted, a waiting call aborted in the queue, and a running traversal aborted between slices all reject with `CalcError` code `aborted`. A soft deadline rejects with `soft_timeout`. The workbook, revision and undo history stay unchanged, and the session remains usable for the next call.

```js
import { createInlineCalc } from 'xlcalc'

const calc = createInlineCalc()
const session = await calc.createSession({ sheets: [{ name: 'Plan', cells: { A2: { value: 4 }, B2: { formula: 'A2*2' } } }] })
const controller = new AbortController()
const pending = session.explain('Plan', 'B2', { depth: 5, signal: controller.signal, timeoutMs: 500 })
controller.abort('active cell moved')
const e = await pending.catch((e) => e)
console.log(e.code, e.cause)
// → aborted active cell moved
// the session goes on
console.log((await session.explain('Plan', 'B2')).root.value)
// → 8
await calc.close()
```

Each local tree node carries `location: { sheet, range }` for navigation. Its `value`, invariant `formula`, `children`, `cells`, `spillOf` and `origin` are available directly, so a panel can traverse `root` without splitting `text`. Use `scanReferences(node.formula)` for formula links and fill in a reference without a sheet from `node.location.sheet`; no rendered explanation needs parsing. Nodes in another workbook have no local `location`.

A session's `clock` is its clock as of its last completed call.

`getHash()` returns the values hash that `getResult()` would produce, without building the result: from a thread, on the synthetic workbook of 34,000 rows, it takes 0.34 s with the event loop free (`getResult()` takes 1.4 s and blocks it for 260 ms while the result is read).

`getCells(sheet, range, opts?)` returns the cells of one rectangle as `getResult()` gives them (`signal` and `timeoutMs` stop it as they stop a calculation), with each cell's own `style`, so a viewport over a sheet needs no copy of the whole result. From a thread, only the rectangle comes back: 20,000 cells of the synthetic workbook of 34,000 rows take 27 ms there and 16 ms in a synchronous session, while the whole `getResult()` from the thread takes 1.7 s. Given several rectangles, `getCells(sheet, ranges)` reads them in one call, so the rows of a filtered viewport, even when far apart, are one round trip to the thread. Ten rectangles of 50 rows of the synthetic workbook of 34,000 rows take 11 ms from a thread, in one call or in ten; with a height set for each of its 34,000 rows, 15 ms in one call and 20 ms in ten: a session reads a sheet's row and column settings once, and again after they change (it was 42 ms in one call and 82 ms in ten when each call read them).

What a viewport needs besides its cells is also available without `getBook()`. `getLayout()` returns the workbook without its cells, and `getUsedRange(sheet)` returns where a sheet's data ends, including spills; neither reads the cells. From a thread, on the synthetic workbook of 34,000 rows, `getLayout()` takes 0.1 ms and `getUsedRange()` of its largest sheet 0.1 ms, while `getBook()` takes 460 ms; with a height for each of the 34,000 rows of that sheet and 500 merged areas, `getLayout()` takes 7 ms. A host that had to find the extent itself read `getBook()` and a viewport over every sheet's formulas for their spills, about 700 ms there.

`patchCells` changes cells' styles, formats, meta, notes and hyperlinks without touching their values. A host that formats a rectangle sends just the patch, whereas with `setCells` it had to read the cells first to send their values back: 2,000 cells of the synthetic workbook of 34,000 rows take 9 ms from a thread as a patch, and 15 ms read and written back. It goes into the session's journal like the other calls that change the session, so a session whose thread was lost comes back with its patches.

On the same thread, `getResult()`, `getHash()` and `getCells()` yield to the event loop between slices, as a calculation does. On the synthetic workbook of 34,000 rows, `getResult()` takes 0.95 s there and the event loop waits at most 50 ms: listing the 506,594 cells of its largest sheet is a single step.

A call that stopped while recalculating (`setCells`, `setClock` or `setLayout` past its deadline, aborted, or over the limits) leaves the session needing to recalculate its workbook, and the session's next call does that first. That recalculation does not count toward the next call's `timeoutMs`; only its `signal` and a pool's hard deadline bound it, so a call after a stopped one can exceed its `timeoutMs` by the duration of a cold calculation of the workbook.

## Deadlines and cancellation

`timeoutMs` of a call is a soft deadline: milliseconds of calculation, counted from when the call starts calculating, not while it waits for its turn. `signal` stops a call at any time: a waiting call is rejected immediately, and a running one at its next check. For a deadline measured from the moment of the call, pass `AbortSignal.timeout(ms)`.

The engine checks the deadline and the signal before each formula and inside long loops: a range read in full (entire columns), operators and functions over arrays, a spill and its cells, the indexes and selection of `SUMIFS`, `COUNTIFS` and related functions, the matrix functions, iterative calculation, the result's cells and its hashes, and the cells of a structure operation (each cell it moves or reads and each formula it rewrites). A slice can run past `sliceMs` by the duration of the formula in progress when it ends. Some steps have no check inside: listing or copying the cells of a sheet (40 ms for the 506,594 cells of the largest sheet of the synthetic workbook of 34,000 rows), sorting them, and preparing a workbook's state before its formulas are calculated, which a structure operation and new names do for the whole workbook. In a pool, `hardTimeoutMs` covers these: a task that runs longer, counted from when it reaches its thread, terminates the thread.

A structure operation stops at its checks but does not yield to the event loop. On the same thread, a signal that fires while it runs is seen when it ends, and its deadline stops it partway: on the synthetic workbook of 34,000 rows, `insertRows` (0.3 s without a deadline) stops after 45 ms with `timeoutMs: 20`. A deadline that passes later, while the new workbook's state is being prepared, can be detected up to 0.4 s late there.

A call stopped by its deadline or its signal throws `CalcError` with `soft_timeout` or `aborted` (the signal's reason is the error's `cause`).

## Limits

`limits` bound what a workbook may produce: `formulas`, the number of formulas in the workbook; `cells`, the values of one range read in full or one array; and `spill`, the values one spill fills; plus `writes`, what one edit writes (below). `SUM`, `COUNT`, `MIN` and `MAX` of one range read it cell by cell without materializing it, so `cells` does not bound them. These are also options of `compute` and `createSession`, not only of the calculators. A calculator's `limits` apply to every call; a call's own `limits` override them field by field, so a call that gives only `formulas` keeps the calculator's `cells` and `spill`.

```js
import { compute } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { formula: 'SUM(SEQUENCE(3000,3000))' } } }] }
try {
  compute(book, { limits: { formulas: 100_000, cells: 5_000_000, spill: 500_000 } })
} catch (e) {
  console.log(e.code, e.limit, e.max, e.size, e.cell)
// → limit cells 5000000 9000000 { sheet: 'S', cell: 'A1' }
}
```

A workbook over a limit is not calculated: the call throws `CalcError` with code `limit`, `limit` (`formulas`, `cells` or `spill`), `max`, `size` and, for `cells` and `spill`, the formula's `cell`. A limit is not a cell error such as `#NUM!`, because then the values and their hash would depend on the limits; a result is either exact or absent. Arrays are checked before they are created: without `cells`, `SUM(SEQUENCE(20000,20000))` requests 400 million values and crashes a Node process.

A spill costs more than its array, because each of its cells gets entries in the session's tables: the sheet's cells, the changes of a call, the cells of a result. So it has a limit of its own, checked before its cells are written. `spill` counts the cells a spill stores. A large array whose cells beyond some row and column all hold one value stores only the part before them ([spill.md](spill.md)): `A:A*2` over 500 rows of data stores 500 of its 1,048,576 cells, so such a formula is not rejected. Without `spill`, `cells` bounds spills too.

`writes` bounds an edit rather than a workbook: the cells one session call writes a value or a formula to. Copying one cell onto an entire column, as Excel pastes into a selected column, writes 1,048,576 cells: 2.5 s and 205 MiB of retained heap with a value, and 6.7 s and 1.4 GiB with `=ROW()*2`, which `formulas` stops only after 2.5 s and does not stop at all for the value. The cells are counted before anything is written or built, from the sizes of the ranges and the cells the source contains, never from the destination's cells, so `writes: 100_000` rejects that copy in under a millisecond. On the synthetic workbook of 34,000 rows the rejection takes 75 ms, the time to read the 506,594 cells of the source's sheet once; a copy that proceeds reuses that read. What counts:

- a copy (`editStructure` with `copyRange`): the cells of its source that contain a value or a formula, once for each copy of the source it places in its destination: `A1` onto `B:B` 1,048,576 times, `A1:A3` onto `B:B` once, as a destination that is not a multiple of the source gets one copy at its top. Each cell in the range of a legacy array formula that it partly takes counts, as it is pasted as an array of its own. A source over a filtered sheet's hidden rows counts the values of its visible rows, which it pastes, from the scratch copy the operation makes of them (0.2 s for an entire column, made again by a copy that proceeds); a destination over hidden rows places its copies in each run of visible rows, read from the sheet's spans of hidden rows;
- a move (`moveRange`): the cells of its source that contain a value or a formula;
- `fillCellText`: the cells of its range; `pasteText`: the cells of the clipboard's rectangle, with short rows padded, counted before the rectangle is built;
- `replaceInCells`: the cells it found and re-enters, which are known only once Find has scanned them, so they are counted after the search and before anything is written;
- `setCells`, `replaceCells`, `setCellTexts` and `applyChanges`: their cell edits, whose number the host provided;
- the previews of these (`previewFillCellText`, `previewPasteText`, `previewCellTexts`, `previewChanges`), like the calls they preview.

Formatting alone goes to the sheet's `looks` and costs little, so `patchCells` and the formatting a copy or a move carries do not count. What writes no more than the workbook already holds is not bounded: inserting and deleting rows and columns, the other structure operations, `clearCells`, `undo` and `redo`. An edit over `writes` throws `CalcError` with `limit: 'writes'`, `max` and `size` (the cells it would write), and changes nothing: the workbook, its revision, its history and its hash stay as they were. `compute` writes nothing, so `writes` applies to sessions only.

Since a user's edit triggers it, its message is in the workbook's language: `one edit writes 1048576 cells, over the limit of 100000`, or in a ru-RU workbook `Правка записывает слишком много ячеек: 1048576 при пределе 100000`. The other limits' messages, which come from a host's own settings, are in English.

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] }, { limits: { writes: 100_000 } })
try {
  session.editStructure({ op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', range: 'B:B' } })
} catch (e) {
  console.log(e.code, e.limit, e.max, e.size)
// → limit writes 100000 1048576
}
// nothing changed
console.log(session.revision, session.getBook().sheets[0].cells)
// → 0 { A1: { value: 1 } }
```

## Errors

`CalcError` has a `code`:

| code | reason |
|---|---|
| `soft_timeout` | the call's `timeoutMs` passed |
| `aborted` | the call's `signal` fired |
| `hard_timeout` | the task ran past `hardTimeoutMs`; its thread was terminated |
| `out_of_memory` | the thread reached `maxHeapMb` |
| `thread_failed` | the thread exited or failed outside a calculation |
| `restore_failed` | the session could not be restored on a thread; it is closed |
| `limit` | the workbook or the edit exceeds `limits` |
| `queue_full` | `maxQueue` tasks are already waiting for a thread |
| `too_many_sessions` | the pool holds `maxSessions` sessions |
| `closed` | the calculator or the session is closed |
| `stale_revision` | an edit's `baseRevision` is not the session's revision; it changed nothing ([Editing a grid](edits.md#revisions)) |

Invalid input still throws `InputError`, from a thread too. A workbook or an argument that structured cloning cannot transfer to a thread (a function in a cell's `meta`, which the contract types as JSON) is rejected by the pool with an `InputError` naming the argument (`book`, `edits`); the synchronous and same-thread calls pass `meta` through untouched, as before. A `meta` that holds a BigInt or a cycle, which neither the input hash nor a thread can take, is refused with `InputError`.

## On the same thread

`createInlineCalc({ sliceMs, timeoutMs, limits })` calculates in slices of `sliceMs` milliseconds (10 by default) and yields to the event loop between them, through `setImmediate` in Node and a `MessageChannel` message in a browser, so timers, input and other messages run while a workbook calculates. `timeoutMs` and `limits` are the defaults for every call. Several sessions and calculations of one calculator share the thread, slice by slice.

`close()` stops what is running and rejects pending calls with `closed`.

## In threads

`createCalcPool(opts)` starts `threads` threads (2 by default). A session's workbook is sent to its thread once, when the session is created; after that, only edits are sent there and only changes come back.

A workbook is transferred to and from a thread as JSON text, its cells in chunks of 16,384, written and read between turns of the event loop, so the calling side never blocks for a whole workbook. `createSession` and `compute` take the workbook immediately (its cells, not copies of them: what the caller changes in the workbook afterwards does not reach the pool, as with a synchronous session) and serialize it while waiting; a task keeps its place in the queue meanwhile. On the synthetic workbook of 34,000 rows (506,594 data cells), creating a session blocks the event loop for 62 ms, where copying the workbook and cloning it to the thread took 324 to 395 ms. A workbook that JSON cannot represent as is (a non-finite number or `-0`, a function, a date object or an undefined value in `meta`) is transferred as objects, by structured clone, as before.

- A session is placed on the thread that holds the fewest sessions and stays there. One-off calculations wait in a single queue for any thread. A thread that becomes idle takes the next call of its own sessions first, then the next one-off calculation.
- `maxQueue` bounds the tasks waiting for a thread (`queue_full`); a task an idle thread takes immediately does not wait, so `maxQueue: 0` rejects only tasks that would wait. `maxSessions` bounds the open sessions of the pool (`too_many_sessions`).
- `maxSessionsPerThread` bounds the sessions a thread holds at once. Beyond it, the least recently called session leaves the thread: the thread sends its workbook back as the session's new snapshot and forgets it, and the session is restored on its next call by opening that snapshot, with no journal to replay (see below). Sessions that take turns on a full thread are restored again and again; give the pool enough capacity for the sessions in use.
- `hardTimeoutMs` terminates a thread whose task runs longer. Restoring a session opens its workbook and replays its journal, each of which completed within the deadline once, so a restore gets a deadline of `1 +` the number of calls it replays. `maxHeapMb` is each thread's heap in Node (`resourceLimits.maxOldGenerationSizeMb`); Node terminates a thread that fills it, and its task fails with `out_of_memory`. Node terminates a thread cleanly when many small objects fill its heap, but a single allocation larger than what remains of the heap crashes the whole process: a large array, or the tables of a large spill's cells while its result, hash or changes are built. So a pool with `maxHeapMb` bounds both by the heap where its `limits` set none: `cells` is `maxHeapMb · 2^20 / 64`, one value per 64 bytes of heap, and `spill` is the smaller of `cells` and `maxHeapMb · 2^20 / 1024`, one value per KB. For 384 MB that is 6,291,456 values in one array and 393,216 in one spill, so `SUM(SEQUENCE(500000))` and `TOCOL(A:A)` calculate there. A host that raises `cells` keeps the bound on spills; one that sets `spill` sets it explicitly. The range caches of a thread's sessions and calculations share a quarter of `maxHeapMb`, 96 MiB for 384 MB, and drop the least recently read ranges beyond it ([cache.md](cache.md#cache-budget)).
- The two divisors were measured on Node 24 and 26, each run in a separate process. Without limits, an array (`SUM(SEQUENCE(n))`, `SEQUENCE(n)*2`, `MMULT`, `MUNIT`, `TOCOL` of entire columns, text built with `&`) crashed the process starting at 30 bytes of heap per value (35 million values in 1,024 MB) on heaps of 256 to 2,048 MB; below that, a thread ran out of memory cleanly. 64 bytes is twice that. A spill read by `compute`, `getResult()`, `getHash()` or `setCells` crashed the process starting at 286 bytes of heap per value (1.4 million values in 384 MB, 2.8 million in 768 MB), and 1,024 bytes is over three times that. With the defaults, these formulas, at the limits and beyond, were calculated or rejected with `limit` on heaps of 256, 384, 512, 768 and 1,024 MB on both Node versions (720 runs); at twice the defaults they were calculated, rejected or ran out of memory cleanly (660 runs). The process never crashed.
- An idle thread keeps no process alive: an open pool does not prevent Node from exiting. A Node thread starts without the process's own flags (`execArgv`), which are meant for the host's entry point.

Live sessions and heavy one-off calculations (building versions of a workbook, for example) are best kept in two separate pools, so that a long calculation does not occupy a thread the sessions need.

### When a thread is lost

For each session, the pool keeps a snapshot of its workbook and a journal of the calls that changed it since the snapshot, with the clock the session had after each one; the snapshot carries the session's revision, and the journal each call's edit options, so a restored session has its revision and checks a `baseRevision` as it would have. The first snapshot is the workbook passed to `createSession`; when the journal reaches `snapshotEvery` calls (64 by default), the pool takes a new snapshot from the thread and drops the journal up to it.

When a thread is terminated or fails, its task fails, and a new thread starts in its place when a task needs one; a thread that cannot start fails one task at a time. The sessions it held are restored on their next call: the new thread opens the snapshot and replays the journal with the recorded clocks (for a live session too), and then handles the call. The call that was running is not in the journal, so the session is as it was before that call, with the same hash as a session that never lost its thread. Each changing call carries a sequence number; a thread returns the result of the last one again rather than applying it twice, and rejects one out of order. A session whose restore fails is closed with `restore_failed`, and the thread discards what it had of it; the thread's other sessions continue. A session closed while its thread is lost is still closed, and so is one whose `close()` is pending when the pool closes.

This has a cost on the calling side: it holds each session's workbook as text. On the synthetic workbook of 34,000 rows the text takes 29 MB (the workbook's objects took 68 MB), a new snapshot from the thread blocks the event loop for 3 ms (it was 234 to 290 ms), and `getBook()` takes 450 ms, read one chunk at a time, blocking the event loop for at most 20 ms (it was 390 ms, with a single 230 ms block).

### Undo and redo

`undo()` and `redo()` of an asynchronous session accept `timeoutMs` and `signal` as the other calls do, and one stopped by them changes nothing: the workbook, its hash and the history stay as they were. `history` is the session's history as of its last completed call, so a toolbar reads it without a call.

In a pool, the history lives in the thread, alongside the workbook. Undo and redo change the workbook, so they go into the journal like the other changes, and a snapshot carries the history the session had at that point (as JSON text, like the workbook, when JSON can represent it). A session restored from its snapshot and journal, after its thread was lost or when it returns to a full thread (`maxSessionsPerThread`), has the history it would have had without leaving: the same steps to undo and redo, and an undo after the restore produces the same workbook as an undo without the loss. The history is not truncated at a snapshot. For a structure operation, whether done, undone or redone, the thread sends where it moved the cells (as numbers: rows or columns, a range, cells pasted over, a sheet) and the sheets of both workbooks, and the calling side builds `mapCell` from them.

A pool with `maxHeapMb` bounds the histories of a thread's sessions by its heap, as it bounds their range caches: a session that sets no `undo.maxMb` keeps `maxHeapMb / (8 × maxSessionsPerThread)` MB of steps, rounded down and at least one, so the histories of a full thread together take an eighth of its heap: 8 MB per session for 384 MB and 6 sessions per thread. Without `maxSessionsPerThread`, the pool's `maxSessions` divided by its threads gives the number of sessions per thread, and without either, one session. A session's own `undo` overrides this, and without `maxHeapMb` a session keeps 64 MB, as elsewhere. Beyond the budget the oldest steps are dropped: on 20,000 rows of 10 values, after four deletions of 1,000 rows (about 1.2 MB each), a deletion of 4,000 rows (5 MB) leaves three steps, 7.6 MB in total.

A snapshot grows by what the history holds, which `maxMb` bounds. On the synthetic workbook of 34,000 rows, an `insertRows` adds 0.14 MB of text to it and a `deleteRows` of 1,000 rows 0.53 MB; cell edits add their old inputs, a few dozen bytes each.

### In a browser

The same pool runs on Web Workers with the same messages, started as `new Worker(new URL('./thread.js', import.meta.url), { type: 'module' })`, the form bundlers recognize. A bundler that keeps that URL as written (esbuild) takes the thread's module, `xlcalc/worker/thread`, as a second entry point named `thread` in the same directory (`entryPoints: { app: 'src/app.js', thread: 'xlcalc/worker/thread' }`), for a Node app too; in a browser, where the thread's script lands elsewhere, `spawn` starts it: `spawn: () => new Worker(new URL('./xlcalc-thread.js', import.meta.url), { type: 'module' })`. A hard deadline terminates the worker. There is no memory limit: `maxHeapMb` throws `InputError` there.

`xlcalc` imports none of the thread code: a page that calculates on its own thread does not load it.

### Locales

A pool's threads calculate in the locales loaded where the pool runs (`registerLocale`): a thread starts with the locales loaded at that time, and a locale loaded later is sent to every thread, so a host loads ru-RU once, before or after it creates the pool. A locale is sent as its data, which structured cloning can transfer; its messages are functions, which no message between threads can carry, so a thread imports the package's module for the tag itself to get them (`src/worker/locales.ts`), a separate chunk for a bundler. A locale outside the package, or one whose module a thread cannot import, calculates there as sent and produces English messages there, as it does where the pool runs. A thread loads what was sent before it runs its next task. A thread rejects a workbook in a locale that is not loaded with the same `InputError` as the synchronous engine.

The pool is tested in headless Chromium, Firefox and WebKit through Playwright, and produces the same hashes as Node in each. WebKit stands in for Safari, which Playwright does not drive.

## What it costs the synchronous path

The evaluation loops are generators that yield only when a calculator's slice is used up, and each check is a single comparison when no calculator is running. On the full synthetic workbook (3,907 formulas), compared with the engine before them, five runs each in alternation, medians: a cold calculation took 322 and 327 ms with closures, 341 and 336 ms with eval; an edit of a criteria cell took 1.9 and 1.8 ms at p95 with closures, 1.8 and 1.9 ms with eval. The values hash is the same.

## How it is tested

Not against Excel, but against the synchronous engine: the test workbooks, edits, structure operations, undo and the loss of a thread produce the synchronous engine's values and hashes, on the same thread and in threads, with closures and eval, in Node and in Chromium, Firefox and WebKit.
