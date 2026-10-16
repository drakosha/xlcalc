# Asynchronous and threaded computation

A server that holds books for many users, or a page that must stay responsive, cannot let one large book hold its thread. The engine computes the same way in three places, with the same values and hashes:

- synchronously, as `compute` and `createSession` always did;
- in the same thread, in portions: `createInlineCalc` gives way to the event loop between formulas, and a call stops at its soft deadline or when its `AbortSignal` fires;
- in a pool of threads: `createCalcPool` from `xlcalc/worker` keeps each session's book in a thread (Node `worker_threads`, or Web Workers in a browser), terminates a thread that runs past a hard deadline, bounds a thread's heap, and brings a session back after its thread is lost.

Both calculators have one API: `compute(book, opts)`, `createSession(book, opts)` and `close()`, all promises. A session has the methods of `Session`, as promises.

```js
import { createInlineCalc } from 'xlcalc'
import { createCalcPool } from 'xlcalc/worker'

const calc = createInlineCalc()                                  // the same thread
const pool = createCalcPool({ threads: 4, hardTimeoutMs: 10_000, maxHeapMb: 512 })

const result = await calc.compute(book, { timeoutMs: 2_000 })
const session = await pool.createSession(book)
const changes = await session.setCells(edits, { signal: AbortSignal.timeout(5_000) })
```

## In short

`createInlineCalc()` computes in the same thread, in slices that let timers and input run between them; `createCalcPool()` from `xlcalc/worker` computes in a pool of threads, where a session's book lives in its thread. Both have one API: `compute`, `createSession` and the session's methods return promises, and every call takes `timeoutMs`, a soft deadline of computing time, and `signal`, an `AbortSignal`. A call that fails changes nothing. `limits` bound the formulas of a book, the values of one array, the cells of one spill and the cells one edit of a session writes (`writes`: a copy of one cell onto a whole column writes 1,048,576), in the synchronous `compute` and `createSession` too; over them the call throws `CalcError` rather than computing or writing.

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

A task that runs past the pool's `hardTimeoutMs` terminates its thread, and in Node `maxHeapMb` bounds each thread's heap. The pool keeps a snapshot and a journal of each session's book, so a session whose thread was lost comes back on its next call as it was before the call that failed, its history of undo and redo too. With `maxHeapMb` a session's history holds `maxHeapMb / (8 × maxSessionsPerThread)` MB by default, so the histories of a full thread take an eighth of its heap; a session's own `undo.maxMb` takes its place. [docs/async.md](async.md) has the rules, the error codes and what this costs.

## Calls

A session takes its calls in order: each starts when the one before has finished. Two `setCells` called without waiting give each promise its own changes, as two synchronous calls would. A call that fails changes nothing: the session is as it was before the call, and the next call works on that state.

`editStructure` answers `{ ok, mapCell, changes }` without the book, as a synchronous session's does, so a thread does not send the book back on every operation; `getBook()` fetches it. From a thread, `getBook()`, `getResult()`, `explain()` and `getViewport()` answer copies. `getSnapshot()` answers independent copies of the book and its result in one call in every kind of session: no edit comes between them, later edits leave both unchanged, and modifying the snapshot cannot change the session. In the same thread its book is copied in portions that yield to the event loop; a worker's transport already copies it. `getBook()` and `getLayout()` in the same thread still expose the session's objects and must not be modified.

`explain(sheet, cell, opts?)` takes its tree options together with `signal` and `timeoutMs` in both asynchronous calculators. A signal already aborted, a waiting call aborted in the queue and a running traversal aborted between portions reject with `CalcError` code `aborted`. A soft deadline rejects with `soft_timeout`. The book, revision and undo history stay unchanged, and the next call can use the session.

```js
const controller = new AbortController()
const pending = session.explain('Plan', 'B2', { depth: 5, signal: controller.signal, timeoutMs: 500 })
controller.abort('active cell moved')
```

Each local tree node carries `location: { sheet, range }` for navigation. Its `value`, invariant `formula`, `children`, `cells`, `spillOf` and `origin` are available directly, so a panel can traverse `root` without splitting `text`. Use `scanReferences(node.formula)` for formula links and fill a reference without a sheet from `node.location.sheet`; no rendered explanation needs parsing. Nodes of another workbook have no local `location`.

A session's `clock` is its clock as of its last finished call.

`getHash()` answers the values hash `getResult()` would give, without building the result: from a thread, on the synthetic book of 34,000 rows, 0.34 s with the event loop free (`getResult()` takes 1.4 s and stops it for 260 ms while the result is read).

`getCells(sheet, range, opts?)` answers the cells of one rectangle as `getResult()` gives them (`signal` and `timeoutMs` stop it as they stop a calculation), with each cell's own `style`, so a window over a sheet needs no copy of the whole result. From a thread only the rectangle comes back: 20,000 cells of the synthetic book of 34,000 rows take 27 ms there and 16 ms in a synchronous session, where the whole `getResult()` from the thread takes 1.7 s. Given several rectangles, `cells(sheet, ranges)` reads them in one call, so the rows of a filtered window, far apart, are one message to the thread and back. Ten rectangles of 50 rows of the synthetic book of 34,000 rows take 11 ms from a thread in one call or in ten; with a height set for each of its 34,000 rows, 15 ms in one call and 20 ms in ten: a session reads a sheet's row and column settings once, and again after they change (it was 42 ms in one call and 82 ms in ten when each call read them).

What a window needs beside its cells comes without `getBook()` too. `getLayout()` answers the book without its cells, and `getUsedRange(sheet)` where a sheet's cells end, spills included; neither reads the cells. From a thread, on the synthetic book of 34,000 rows, `getLayout()` takes 0.1 ms and `extent()` of its largest sheet 0.1 ms, where `getBook()` takes 460 ms; with a height for each of the 34,000 rows of that sheet and 500 merged areas, `getLayout()` takes 7 ms. A host that found the extent itself read `getBook()` and a window over every sheet's formulas for their spills, about 700 ms there.

`patchCells` changes cells' styles, formats, meta, notes and links without their values. A host that formats a rectangle sends the patch alone, where with `setCells` it read the cells first to send their values back: 2,000 cells of the synthetic book of 34,000 rows take 9 ms from a thread as a patch and 15 ms read and written back. It goes into the session's journal like the other calls that change the session, so a session whose thread was lost comes back with its patches.

In the same thread `getResult()`, `getHash()` and `getCells()` give way to the event loop between portions, as a computation does. On the synthetic book of 34,000 rows `getResult()` takes 0.95 s there and the event loop waits 50 ms at most: listing the 506,594 cells of its largest sheet is one step.

A call that stopped while it recomputed (`setCells`, `setClock` or `setLayout` past its deadline, aborted, or over the limits) leaves the session to compute its book again, and the session's next call does that first. That computation does not count in the next call's `timeoutMs`, only its `signal` and a pool's hard deadline bound it, so a call after a stopped one can take longer than its `timeoutMs` by the time of a cold computation of the book.

## Deadlines and cancellation

`timeoutMs` of a call is a soft deadline: milliseconds of computing, counted from when the call starts computing, not while it waits for its turn. `signal` stops a call at any time: a call that waits is refused at once, a running one at its next check. For a deadline from the moment of the call, pass `AbortSignal.timeout(ms)`.

The engine checks the deadline and the signal before each formula and inside the long loops: a range read whole (whole columns), operators and functions over arrays, a spill and its cells, the indexes and the selection of `SUMIFS`, `COUNTIFS` and their kin, the matrix functions, iterative calculation, the result's cells and its hashes, and a structure operation's cells (each cell it moves or reads and each formula it rewrites). A slice can run past `sliceMs` by the time of the formula in flight when it ends. Some steps have no check inside: listing or copying the cells of a sheet (40 ms for the 506,594 cells of the largest sheet of the synthetic book of 34,000 rows), sorting them, and preparing a book's state before its formulas compute, which a structure operation and new names do for the whole book. In a pool, `hardTimeoutMs` covers these: a task that runs longer, counted from when it reaches its thread, terminates the thread.

A structure operation stops at its checks but does not give way to the event loop. In the same thread a signal that fires while it runs is seen when it ends, and its deadline stops it on the way: on the synthetic book of 34,000 rows `insertRows` (0.3 s without a deadline) stops after 45 ms with `timeoutMs: 20`. A deadline that passes later, while the new book's state is prepared, is seen up to 0.4 s late there.

A call that stops by its deadline or its signal throws `CalcError` with `soft_timeout` or `aborted` (the signal's reason is the error's `cause`).

## Limits

`limits` bound what a book may make: `formulas`, the formulas of the book, `cells`, the values of one range read whole or one array, and `spill`, the values one spill fills; and `writes`, what one edit writes (below). `SUM`, `COUNT`, `MIN` and `MAX` of one range read it cell by cell without making it whole, so `cells` does not bound them. They are options of `compute` and `createSession` too, not only of the calculators. A calculator's `limits` hold for every call; a call's own `limits` take their place field by field, so a call that gives `formulas` alone keeps the calculator's `cells` and `spill`.

```js
compute(book, { limits: { formulas: 100_000, cells: 5_000_000, spill: 500_000 } })
```

A book over a limit is not computed: the call throws `CalcError` with code `limit`, `limit` (`formulas`, `cells` or `spill`), `max`, `size` and, for `cells` and `spill`, the formula's `cell`. A limit is no cell error such as `#NUM!`, because then the values and their hash would depend on the limits; a result is either exact or absent. Arrays are checked before they are made: without `cells`, `SUM(SEQUENCE(20000,20000))` asks for 400 million values and brings a Node process down.

A spill costs more than its array, because each of its cells gets entries in the session's tables: the sheet's cells, the changes of a call, the cells of a result. So it has a limit of its own, checked before its cells are written. `spill` counts the cells a spill stores. A large array whose cells past some row and column all hold one value stores only the part before them ([spill.md](spill.md)): `A:A*2` over 500 rows of data stores 500 cells of its 1,048,576, so such a formula is not refused. Without `spill`, `cells` bounds spills too.

`writes` bounds an edit rather than a book: the cells one call of a session writes with a value or a formula. A copy of one cell onto a whole column, as Excel pastes into a selected column, writes 1,048,576 cells: 2.5 s and 205 MiB of heap kept with a value, 6.7 s and 1.4 GiB with `=ROW()*2`, which `formulas` stops only after 2.5 s and the value not at all. The cells are counted before anything is written or built, from the sizes of the ranges and the cells the source holds, never from the target's cells, so `writes: 100_000` refuses that copy in under a millisecond. On the synthetic book of 34,000 rows the refusal takes 75 ms, the time to read the 506,594 cells of the source's sheet once; a copy that goes on takes them from that reading. What counts:

- a copy (`editStructure` with `copyRange`): the cells of its source that hold a value or a formula, once for each copy of the source it lays in its target: `A1` onto `B:B` 1,048,576 times, `A1:A3` onto `B:B` once, as a target that is no multiple of the source takes one copy at its top. Each cell of an old-kind array formula's area it takes part of counts, as it pastes as an array of its own. A source over a filtered sheet's hidden rows counts the values of its visible rows, which it pastes, on the scratch copy of them the copy makes (0.2 s for a whole column, made again by a copy that goes on); a target over hidden rows lays its copies in each run of visible rows, read from the sheet's spans of hidden rows;
- a move (`moveRange`): the cells of its source that hold a value or a formula;
- `fillCellText`: the cells of its range; `pasteText`: the cells of the clipboard's rectangle, its short rows filled out, counted before the rectangle is made;
- `replaceInCells`: the cells it found and enters again, which are known only once Find has gone over them, so they are counted after the search and before anything is written;
- `setCells`, `replaceCells`, `setCellTexts` and `applyChanges`: their edits of cells, whose number the host gave;
- the previews of these (`previewFillCellText`, `previewPasteText`, `previewCellTexts`, `previewChanges`), as the calls they preview.

A look alone goes to the sheet's looks and costs little, so `patchCells` and the looks a copy or a move carries do not count. What writes no more than the book holds is not bounded: inserting and deleting rows and columns, the other structure operations, `clearCells`, `undo` and `redo`. An edit over `writes` throws `CalcError` with `limit: 'writes'`, `max` and `size`, the cells it would write, and changes nothing: the book, its revision, its history and its hash stay as they were. `compute` writes nothing, so `writes` bounds sessions alone.

As a person's edit brings it, its message is in the book's language: `1048576 cells written by one edit, over the limit of 100000`, in a ru-RU book `Правка записывает слишком много ячеек: 1048576 при пределе 100000`. The other limits' messages, which a host's own settings bring, are in English.

```js
const session = createSession(book, { limits: { writes: 100_000 } })
session.editStructure({ op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', range: 'B:B' } }) // CalcError: limit writes, size 1048576
```

## Errors

`CalcError` has a `code`:

| code | why |
|---|---|
| `soft_timeout` | the call's `timeoutMs` passed |
| `aborted` | the call's `signal` fired |
| `hard_timeout` | the task ran past `hardTimeoutMs`; its thread was terminated |
| `out_of_memory` | the thread reached `maxHeapMb` |
| `thread_failed` | the thread exited or failed outside a computation |
| `restore_failed` | the session could not be brought back on a thread; it is closed |
| `limit` | the book or the edit is over `limits` |
| `queue_full` | `maxQueue` tasks already wait for a thread |
| `too_many_sessions` | the pool holds `maxSessions` sessions |
| `closed` | the calculator or the session is closed |
| `stale_revision` | an edit's `baseRevision` is not the session's revision; it changed nothing ([Editing a grid](edits.md#revisions)) |

Wrong input still throws `InputError`, from a thread too. A book or an argument that a structured clone cannot carry to a thread (a function in a cell's `meta`, which the contract types as JSON) is refused by the pool with `InputError` naming the argument (`book`, `edits`); the synchronous and the same-thread calls pass `meta` through untouched, as before.

## In the same thread

`createInlineCalc({ sliceMs, timeoutMs, limits })` computes in slices of `sliceMs` milliseconds (10 by default) and gives way to the event loop between them: through `setImmediate` in Node and a `MessageChannel` message in a browser, so timers, input and other messages run while a book computes. `timeoutMs` and `limits` are the defaults of every call. Several sessions and computations of one calculator share the thread, slice by slice.

`close()` stops what runs and waits with `closed`.

## In threads

`createCalcPool(opts)` starts `threads` threads (2 by default). A session's book goes to its thread once, when the session opens; afterwards only edits go there and only changes come back.

A book crosses to and from a thread as JSON text, its cells in chunks of 16,384 written and read between turns of the event loop, so the calling side never stops for a whole book. `createSession` and `compute` take the book at once (its cells, not copies of them: what the caller changes in the book afterwards does not reach the pool, as with a synchronous session) and write it while it waits; a task keeps its place in the queue meanwhile. On the synthetic book of 34,000 rows (506,594 cells of data) opening a session stops the event loop for 62 ms, where copying the book and cloning it to the thread took 324 to 395 ms. A book that JSON cannot hold as it is (a non-finite number or `-0`, a function, a date object or an undefined value in `meta`) goes as objects, by structured clone, as before.

- A session is placed on the thread that holds the fewest sessions and stays there. One-off computations wait in one queue for any thread. A thread that becomes idle takes the next call of its own sessions first, then the next one-off computation.
- `maxQueue` bounds the tasks that wait for a thread (`queue_full`); a task an idle thread takes at once does not wait, so `maxQueue: 0` refuses only what would wait. `maxSessions` bounds the open sessions of the pool (`too_many_sessions`).
- `maxSessionsPerThread` bounds the sessions a thread holds at once. Past it, the session called least recently leaves the thread: the thread sends its book back as the session's new snapshot and forgets it, and the session comes back on its next call by opening that snapshot, with no journal to replay (see below). Sessions that take turns on a full thread come back again and again; give the pool enough room for the sessions in use.
- `hardTimeoutMs` terminates a thread whose task runs longer. Bringing a session back opens its book and replays its journal, each of which was answered within the deadline once, so a restore has the deadline of `1 +` the calls it replays. `maxHeapMb` is each thread's heap in Node (`resourceLimits.maxOldGenerationSizeMb`); a thread that fills it is terminated by Node and its task fails with `out_of_memory`. Node terminates a thread cleanly when many small objects fill its heap, but one allocation larger than what the heap has left brings the whole process down: a large array, or the tables of a large spill's cells while its result, hash or changes are made. So a pool with `maxHeapMb` bounds both by the heap where its `limits` give none: `cells` is `maxHeapMb · 2^20 / 64`, a value per 64 bytes of heap, and `spill` is the smaller of `cells` and `maxHeapMb · 2^20 / 1024`, a value per KB. For 384 MB that is 6,291,456 values in one array and 393,216 in one spill, so `SUM(SEQUENCE(500000))` and `TOCOL(A:A)` compute there. A host that raises `cells` keeps the bound on spills; one that gives `spill` sets it. The range caches of a thread's sessions and computations hold a quarter of `maxHeapMb` together, 96 MiB for 384 MB, and drop the ranges read least recently past it ([cache.md](cache.md#cache-budget)).
- The two divisors are measured on Node 24 and 26, each run in a process of its own. Without limits, an array (`SUM(SEQUENCE(n))`, `SEQUENCE(n)*2`, `MMULT`, `MUNIT`, `TOCOL` of whole columns, text made by `&`) brought the process down from 30 bytes of heap per value (35 million values in 1,024 MB) on heaps of 256 to 2,048 MB; short of that a thread ran out of memory cleanly. 64 bytes is twice that. A spill read by `compute`, `getResult()`, `getHash()` or `setCells` brought the process down from 286 bytes of heap per value (1.4 million values in 384 MB, 2.8 million in 768 MB), and 1,024 bytes is over three times that. With the defaults these formulas, at the limits and past them, computed or were refused with `limit` on heaps of 256, 384, 512, 768 and 1,024 MB on both Nodes (720 runs); at twice the defaults they computed, were refused or ran out of memory cleanly (660 runs). The process never went down.
- A thread that waits keeps no process alive: an open pool does not stop Node from exiting. A Node thread starts without the process's own flags (`execArgv`), which are meant for the host's entry.

Live sessions and heavy one-off computations (building versions of a book, say) are best kept in two pools, so that a long computation does not hold a thread the sessions need.

### When a thread is lost

The pool keeps, for each session, a snapshot of its book and a journal of the calls that changed it since the snapshot, with the clock the session had after each; the snapshot carries the session's revision, and the journal each call's edit options, so a session brought back has its revision and checks a `baseRevision` as it would have. The first snapshot is the book given to `createSession`; when the journal reaches `snapshotEvery` calls (64 by default), the pool takes a new one from the thread and drops the journal up to it.

When a thread is terminated or fails, its task fails, and a new thread starts in its place when a task needs one; a thread that cannot start fails one task at a time. The sessions it held come back on their next call: the new thread opens the snapshot and replays the journal, with the recorded clocks (a live session too), and then takes the call. The call that was running is not in the journal, so the session is as it was before that call, with the hash of a session that never lost its thread. Each changing call carries a sequence number; a thread answers the last one again rather than applying it twice, and refuses one out of order. A session whose restore fails is closed with `restore_failed`, and the thread forgets what it had of it; the other sessions of the thread go on. A session closed while its thread is lost is closed all the same, and so is one whose `close()` waits when the pool closes.

This has a cost on the calling side: it holds each session's book as text. On the synthetic book of 34,000 rows the text takes 29 MB (the book's objects took 68 MB), a new snapshot from the thread stops the event loop for 3 ms (it was 234 to 290 ms), and `getBook()` takes 450 ms read a chunk at a time, stopping the event loop for 20 ms at most (it was 390 ms in one stretch of 230 ms).

### Undo and redo

`undo()` and `redo()` of an asynchronous session take `timeoutMs` and `signal` as the other calls do, and one stopped by them changes nothing: the book, its hash and the history stay as they were. `history` is the session's history as of its last finished call, so a toolbar reads it without a call.

In a pool the history lives in the thread, beside the book. An undo and a redo change the book, so they go into the journal like the other changes, and a snapshot carries the history the session had then (as JSON text, as the book, when JSON holds it). A session restored from its snapshot and journal, after its thread was lost or when it comes back to a full thread (`maxSessionsPerThread`), has the history it would have had without leaving: the same steps to undo and redo, and an undo after the restore gives the book an undo without the loss gives. The history is not cut at a snapshot. For a structure operation, done, undone or redone, the thread sends where it took the cells (in numbers: lines, a range, cells pasted over, a sheet) and the sheets of both books, and the calling side makes `mapCell` from them.

A pool with `maxHeapMb` bounds the histories of a thread's sessions by its heap, as it bounds their range caches: a session that gives no `undo.maxMb` keeps `maxHeapMb / (8 × maxSessionsPerThread)` MB of steps, rounded down and one at least, so the histories of a full thread hold an eighth of its heap together: 8 MB a session for 384 MB and 6 sessions a thread. Without `maxSessionsPerThread` the pool's `maxSessions` over its threads counts the sessions of a thread, and without either one session. A session's own `undo` takes the place of this, and without `maxHeapMb` a session keeps 64 MB as elsewhere. Past the budget the oldest steps go: on 20,000 rows of 10 values, after four deletes of 1,000 rows (about 1.2 MB each) a delete of 4,000 rows (5 MB) leaves three steps, 7.6 MB together.

What a snapshot holds grows by what the history holds, which `maxMb` bounds. On the synthetic book of 34,000 rows an `insertRows` adds 0.14 MB of text to it and a `deleteRows` of 1,000 rows 0.53 MB; edits of cells add their old inputs, a few dozen bytes each.

### In a browser

The same pool runs on Web Workers with the same messages, started as `new Worker(new URL('./thread.js', import.meta.url), { type: 'module' })`, the form bundlers recognise. Where a bundler puts the thread's script elsewhere, `spawn` starts it. A hard deadline terminates the worker. There is no memory limit: `maxHeapMb` throws `InputError` there.

`xlcalc` imports nothing of the threads: a page that computes in its own thread does not load them.

### Locales

A pool's threads compute in the locales loaded where the pool runs (`registerLocale`): a thread starts with those loaded then, and a locale loaded later is sent to every thread, so a host loads ru-RU once, before or after it makes the pool. A locale goes as its data, which a structured clone carries; its messages are functions, which no message between threads carries, so a thread imports the package's module of the tag itself for them (`src/worker/locales.ts`), a chunk of its own to a bundler. A locale outside the package, or one whose module a thread cannot import, computes there as sent and speaks English there, as it does where the pool runs. A thread loads what was sent before it runs its next task. A book in a locale not loaded is refused by the thread with the `InputError` of the synchronous engine.

The pool is checked in headless Chromium, Firefox and WebKit through Playwright, and gives the hashes of Node in each. WebKit stands there for Safari, which Playwright does not drive.

## What it costs the synchronous path

The evaluation loops are generators that yield only when a calculator's slice is used up, and each check is one comparison when no calculator runs. On the full synthetic book (3,907 formulas), against the engine before them, five runs each taken in turns, medians: a cold computation 322 and 327 ms with closures, 341 and 336 ms with eval; an edit of a criterion cell 1.9 and 1.8 ms at p95 with closures, 1.8 and 1.9 ms with eval. The values hash is the same.

## How it is tested

Not against Excel but against the synchronous engine: the case books, edits, structure operations, undo and the loss of a thread give the synchronous engine's values and hashes, in the same thread and in threads, with closures and eval, in Node and in Chromium, Firefox and WebKit.
