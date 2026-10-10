# Asynchronous and threaded computation

A server that holds books for many users, or a page that must stay responsive, cannot let one large book hold its thread. The engine computes the same way in three places, with the same values and hashes:

- synchronously, as `compute` and `openSession` always did;
- in the same thread, in portions: `createInlineCalc` gives way to the event loop between formulas, and a call stops at its soft deadline or when its `AbortSignal` fires;
- in a pool of threads: `createCalcPool` from `xlcalc/worker` keeps each session's book in a thread (Node `worker_threads`, or Web Workers in a browser), terminates a thread that runs past a hard deadline, bounds a thread's heap, and brings a session back after its thread is lost.

Both calculators have one API: `compute(book, opts)`, `openSession(book, opts)` and `close()`, all promises. A session has the methods of `Session`, as promises.

```js
import { createInlineCalc } from 'xlcalc'
import { createCalcPool } from 'xlcalc/worker'

const calc = createInlineCalc()                                  // the same thread
const pool = createCalcPool({ threads: 4, hardTimeoutMs: 10_000, maxHeapMb: 512 })

const result = await calc.compute(book, { timeoutMs: 2_000 })
const session = await pool.openSession(book)
const changes = await session.setCells(edits, { signal: AbortSignal.timeout(5_000) })
```

## Calls

A session takes its calls in order: each starts when the one before has finished. Two `setCells` called without waiting give each promise its own changes, as two synchronous calls would. A call that fails changes nothing: the session is as it was before the call, and the next call works on that state.

`restructure` answers `{ ok, mapCell, changes }` without the book, which a thread would otherwise send back on every operation; `book()` fetches it. From a thread, `book()`, `result()` and `explain()` answer copies.

A session's `clock` is its clock as of its last finished call.

`hash()` answers the values hash `result()` would give, without building the result: from a thread, on the synthetic book of 34,000 rows, 0.34 s with the event loop free (`result()` takes 1.4 s and stops it for 260 ms while the result is read).

`cells(sheet, range)` answers the cells of one rectangle as `result()` gives them, with each cell's own `style`, so a window over a sheet needs no copy of the whole result. From a thread only the rectangle comes back: 20,000 cells of the synthetic book of 34,000 rows take 27 ms there and 16 ms in a synchronous session, where the whole `result()` from the thread takes 1.7 s. Given several rectangles, `cells(sheet, ranges)` reads them in one call, so the rows of a filtered window, far apart, are one message to the thread and back. Ten rectangles of 50 rows of the synthetic book of 34,000 rows take 11 ms from a thread in one call or in ten; with a height set for each of its 34,000 rows, 15 ms in one call and 20 ms in ten: a session reads a sheet's row and column settings once, and again after they change (it was 42 ms in one call and 82 ms in ten when each call read them).

What a window needs beside its cells comes without `book()` too. `layout()` answers the book without its cells, and `extent(sheet)` where a sheet's cells end, spills included; neither reads the cells. From a thread, on the synthetic book of 34,000 rows, `layout()` takes 0.1 ms and `extent()` of its largest sheet 0.1 ms, where `book()` takes 460 ms; with a height for each of the 34,000 rows of that sheet and 500 merged areas, `layout()` takes 7 ms. A host that found the extent itself read `book()` and a window over every sheet's formulas for their spills, about 700 ms there.

`patchCells` changes cells' styles, formats, meta, notes and links without their values. A host that formats a rectangle sends the patch alone, where with `setCells` it read the cells first to send their values back: 2,000 cells of the synthetic book of 34,000 rows take 9 ms from a thread as a patch and 15 ms read and written back. It goes into the session's journal like the other calls that change the session, so a session whose thread was lost comes back with its patches.

In the same thread `result()`, `hash()` and `cells()` give way to the event loop between portions, as a computation does. On the synthetic book of 34,000 rows `result()` takes 0.95 s there and the event loop waits 50 ms at most: listing the 506,594 cells of its largest sheet is one step.

A call that stopped while it recomputed (`setCells`, `setClock` or `setLayout` past its deadline, aborted, or over the limits) leaves the session to compute its book again, and the session's next call does that first. That computation does not count in the next call's `timeoutMs`, only its `signal` and a pool's hard deadline bound it, so a call after a stopped one can take longer than its `timeoutMs` by the time of a cold computation of the book.

## Deadlines and cancellation

`timeoutMs` of a call is a soft deadline: milliseconds of computing, counted from when the call starts computing, not while it waits for its turn. `signal` stops a call at any time: a call that waits is refused at once, a running one at its next check. For a deadline from the moment of the call, pass `AbortSignal.timeout(ms)`.

The engine checks the deadline and the signal before each formula and inside the long loops: a range read whole (whole columns), operators and functions over arrays, a spill and its cells, the indexes and the selection of `SUMIFS`, `COUNTIFS` and their kin, the matrix functions, iterative calculation, the result's cells and its hashes, and a structure operation's cells (each cell it moves or reads and each formula it rewrites). A slice can run past `sliceMs` by the time of the formula in flight when it ends. Some steps have no check inside: listing or copying the cells of a sheet (40 ms for the 506,594 cells of the largest sheet of the synthetic book of 34,000 rows), sorting them, and preparing a book's state before its formulas compute, which a structure operation and new names do for the whole book. In a pool, `hardTimeoutMs` covers these: a task that runs longer, counted from when it reaches its thread, terminates the thread.

A structure operation stops at its checks but does not give way to the event loop. In the same thread a signal that fires while it runs is seen when it ends, and its deadline stops it on the way: on the synthetic book of 34,000 rows `insertRows` (0.6 s without a deadline) stops after 43 ms with `timeoutMs: 20`. A deadline that passes later, while the new book's state is prepared, is seen up to 0.4 s late there.

A call that stops by its deadline or its signal throws `CalcError` with `soft_timeout` or `aborted` (the signal's reason is the error's `cause`).

## Limits

`limits` bound what a book may make: `formulas`, the formulas of the book, `cells`, the values of one range read whole or one array, and `spill`, the values one spill fills. `SUM`, `COUNT`, `MIN` and `MAX` of one range read it cell by cell without making it whole, so `cells` does not bound them. They are options of `compute` and `openSession` too, not only of the calculators. A calculator's `limits` hold for every call; a call's own `limits` take their place field by field, so a call that gives `formulas` alone keeps the calculator's `cells` and `spill`.

```js
compute(book, { limits: { formulas: 100_000, cells: 5_000_000, spill: 500_000 } })
```

A book over a limit is not computed: the call throws `CalcError` with code `limit`, `limit` (`formulas`, `cells` or `spill`), `max`, `size` and, for `cells` and `spill`, the formula's `cell`. A limit is no cell error such as `#NUM!`, because then the values and their hash would depend on the limits; a result is either exact or absent. Arrays are checked before they are made: without `cells`, `SUM(SEQUENCE(20000,20000))` asks for 400 million values and brings a Node process down.

A spill costs more than its array, because each of its cells gets entries in the session's tables: the sheet's cells, the changes of a call, the cells of a result. So it has a limit of its own, checked before its cells are written. `spill` counts the cells a spill stores. A large array whose cells past some row and column all hold one value stores only the part before them ([spill.md](spill.md)): `A:A*2` over 500 rows of data stores 500 cells of its 1,048,576, so such a formula is not refused. Without `spill`, `cells` bounds spills too.

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
| `limit` | the book is over `limits` |
| `queue_full` | `maxQueue` tasks already wait for a thread |
| `too_many_sessions` | the pool holds `maxSessions` sessions |
| `closed` | the calculator or the session is closed |

Wrong input still throws `InputError`, from a thread too. A book or an argument that a structured clone cannot carry to a thread (a function in a cell's `meta`, which the contract types as JSON) is refused by the pool with `InputError` naming the argument (`book`, `edits`); the synchronous and the same-thread calls pass `meta` through untouched, as before.

## In the same thread

`createInlineCalc({ sliceMs, timeoutMs, limits })` computes in slices of `sliceMs` milliseconds (10 by default) and gives way to the event loop between them: through `setImmediate` in Node and a `MessageChannel` message in a browser, so timers, input and other messages run while a book computes. `timeoutMs` and `limits` are the defaults of every call. Several sessions and computations of one calculator share the thread, slice by slice.

`close()` stops what runs and waits with `closed`.

## In threads

`createCalcPool(opts)` starts `threads` threads (2 by default). A session's book goes to its thread once, when the session opens; afterwards only edits go there and only changes come back.

A book crosses to and from a thread as JSON text, its cells in chunks of 16,384 written and read between turns of the event loop, so the calling side never stops for a whole book. `openSession` and `compute` take the book at once (its cells, not copies of them: what the caller changes in the book afterwards does not reach the pool, as with a synchronous session) and write it while it waits; a task keeps its place in the queue meanwhile. On the synthetic book of 34,000 rows (506,594 cells of data) opening a session stops the event loop for 62 ms, where copying the book and cloning it to the thread took 324 to 395 ms. A book that JSON cannot hold as it is (a non-finite number or `-0`, a function, a date object or an undefined value in `meta`) goes as objects, by structured clone, as before.

- A session is placed on the thread that holds the fewest sessions and stays there. One-off computations wait in one queue for any thread. A thread that becomes idle takes the next call of its own sessions first, then the next one-off computation.
- `maxQueue` bounds the tasks that wait for a thread (`queue_full`); a task an idle thread takes at once does not wait, so `maxQueue: 0` refuses only what would wait. `maxSessions` bounds the open sessions of the pool (`too_many_sessions`).
- `maxSessionsPerThread` bounds the sessions a thread holds at once. Past it, the session called least recently leaves the thread: the thread sends its book back as the session's new snapshot and forgets it, and the session comes back on its next call by opening that snapshot, with no journal to replay (see below). Sessions that take turns on a full thread come back again and again; give the pool enough room for the sessions in use.
- `hardTimeoutMs` terminates a thread whose task runs longer. Bringing a session back opens its book and replays its journal, each of which was answered within the deadline once, so a restore has the deadline of `1 +` the calls it replays. `maxHeapMb` is each thread's heap in Node (`resourceLimits.maxOldGenerationSizeMb`); a thread that fills it is terminated by Node and its task fails with `out_of_memory`. Node terminates a thread cleanly when many small objects fill its heap, but one allocation larger than what the heap has left brings the whole process down: a large array, or the tables of a large spill's cells while its result, hash or changes are made. So a pool with `maxHeapMb` bounds both by the heap where its `limits` give none: `cells` is `maxHeapMb · 2^20 / 64`, a value per 64 bytes of heap, and `spill` is the smaller of `cells` and `maxHeapMb · 2^20 / 1024`, a value per KB. For 384 MB that is 6,291,456 values in one array and 393,216 in one spill, so `SUM(SEQUENCE(500000))` and `TOCOL(A:A)` compute there. A host that raises `cells` keeps the bound on spills; one that gives `spill` sets it. The range caches of a thread's sessions and computations hold a quarter of `maxHeapMb` together, 96 MiB for 384 MB, and drop the ranges read least recently past it ([cache.md](cache.md#cache-budget)).
- The two divisors are measured on Node 24 and 26, each run in a process of its own. Without limits, an array (`SUM(SEQUENCE(n))`, `SEQUENCE(n)*2`, `MMULT`, `MUNIT`, `TOCOL` of whole columns, text made by `&`) brought the process down from 30 bytes of heap per value (35 million values in 1,024 MB) on heaps of 256 to 2,048 MB; short of that a thread ran out of memory cleanly. 64 bytes is twice that. A spill read by `compute`, `result()`, `hash()` or `setCells` brought the process down from 286 bytes of heap per value (1.4 million values in 384 MB, 2.8 million in 768 MB), and 1,024 bytes is over three times that. With the defaults these formulas, at the limits and past them, computed or were refused with `limit` on heaps of 256, 384, 512, 768 and 1,024 MB on both Nodes (720 runs); at twice the defaults they computed, were refused or ran out of memory cleanly (660 runs). The process never went down.
- A thread that waits keeps no process alive: an open pool does not stop Node from exiting. A Node thread starts without the process's own flags (`execArgv`), which are meant for the host's entry.

Live sessions and heavy one-off computations (building versions of a book, say) are best kept in two pools, so that a long computation does not hold a thread the sessions need.

### When a thread is lost

The pool keeps, for each session, a snapshot of its book and a journal of the calls that changed it since the snapshot, with the clock the session had after each. The first snapshot is the book given to `openSession`; when the journal reaches `snapshotEvery` calls (64 by default), the pool takes a new one from the thread and drops the journal up to it.

When a thread is terminated or fails, its task fails, and a new thread starts in its place when a task needs one; a thread that cannot start fails one task at a time. The sessions it held come back on their next call: the new thread opens the snapshot and replays the journal, with the recorded clocks (a live session too), and then takes the call. The call that was running is not in the journal, so the session is as it was before that call, with the hash of a session that never lost its thread. Each changing call carries a sequence number; a thread answers the last one again rather than applying it twice, and refuses one out of order. A session whose restore fails is closed with `restore_failed`, and the thread forgets what it had of it; the other sessions of the thread go on. A session closed while its thread is lost is closed all the same, and so is one whose `close()` waits when the pool closes.

This has a cost on the calling side: it holds each session's book as text. On the synthetic book of 34,000 rows the text takes 29 MB (the book's objects took 68 MB), a new snapshot from the thread stops the event loop for 3 ms (it was 234 to 290 ms), and `book()` takes 450 ms read a chunk at a time, stopping the event loop for 20 ms at most (it was 390 ms in one stretch of 230 ms).

### Undo and redo

`undo()` and `redo()` of an asynchronous session take `timeoutMs` and `signal` as the other calls do, and one stopped by them changes nothing: the book, its hash and the history stay as they were. `history` is the session's history as of its last finished call, so a toolbar reads it without a call.

In a pool the history lives in the thread, beside the book. An undo and a redo change the book, so they go into the journal like the other changes, and a snapshot carries the history the session had then (as JSON text, as the book, when JSON holds it). A session restored from its snapshot and journal, after its thread was lost or when it comes back to a full thread (`maxSessionsPerThread`), has the history it would have had without leaving: the same steps to undo and redo, and an undo after the restore gives the book an undo without the loss gives. The history is not cut at a snapshot. For a structure operation undone or redone, the thread sends the operation and the sheets of both books, and the calling side makes `mapCell` from them.

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

## How it was checked

- `test/async-inline.test.ts` and `test/async-pool.test.ts` compute the coverage measurement (in en-US and ru-RU), the semantics cases, the synthetic book, the spill book, a seeded book, iterative calculation and the 1904 date system in the same thread and in threads, with closures and eval, and compare values, issues and hashes with the synchronous engine; a session over 40 edits gives the changes and the hash of a synchronous session after each, in both execution modes and in both calculators; two sessions on one thread with their calls mixed without waiting each give their own changes and hash; a structure operation gives the synchronous session's changes and `mapCell` on eleven operations, filtered rows among them.
- In the same thread a timer runs while a book computes; deadlines, aborts (waiting and running) and limits give `CalcError`; a stopped `setCells`, `setLayout` or `restructure` leaves the book and the hash as they were; 300 formulas that each sort 10,000 values stop at a deadline of 100 ms within 200 ms, the event loop never waiting 50 ms (measured alone: 100 to 102 ms, waits of 10 to 21 ms with slices of 10 ms).
- `test/async-inline-long.test.ts`: in the same thread, on the synthetic book of 8,500 rows, `result()`, `hash()` and `cells()` of a large rectangle give the synchronous session's values while the event loop runs, never waiting a quarter of the call's time; `insertRows`, `deleteColumns`, `renameSheet`, `moveRange` and `renameTable` over 60,000 formulas stop at a deadline of 20 ms within a quarter of the operation's time, and the session keeps its hash and then gives the synchronous session's changes. The result's steps and the checks of structure operations cost the synchronous path nothing measurable: a cold computation of the synthetic book, five runs each taken in turns with the engine without them, medians, two rounds: 246 and 253 ms against 249 and 245 ms with closures, 296 and 297 against 298 and 302 ms with eval (3,400 rows); 1,833 and 1,656 against 1,864 and 1,629 ms with closures, 1,821 and 1,820 against 1,821 and 1,802 ms with eval (34,000 rows); the values hash is the same.
- `test/async-recovery.test.ts` terminates threads by a hard deadline, by `maxHeapMb` (twenty spilled arrays of a million values in a 64 MB heap) and by a crash, and checks that the session's hash after it comes back equals that of a synchronous session given the same answered calls, from a snapshot and a journal cut at it too; a restore that fails closes that session only, and ten of them on a 16 MB book leave a 96 MB thread alive; a session of ten calls of 150 ms comes back within a deadline of 800 ms after it left a full thread and after a crash; sessions take turns on a full thread; `maxQueue`, `maxSessions`, the order of a thread's tasks, closing.
- `test/session-window.test.ts`: in a synchronous session, in the same thread and in a thread, with closures and eval, each sheet's `extent` equals the reach of the cells, spill ranges, merged areas and tables of the session's own result and book, on the books of the tests and while edits grow and shrink it (a spill taller, blocked, shorter, a cell far away patched and cleared, merged areas, inserted rows, a deleted table); `layout()` equals the book without its cells after layout edits, a structure operation and new styles, and a thread's is a copy; `cells` of several rectangles equals the rectangles read one by one, each cell once and in address order. `test/session-patch.test.ts`: patches on an empty cell, a value, a formula, an old-kind array's area and a spill keep the values hash and the arrays, list the cells a new format shows otherwise, leave a live session's clock as it was, change nothing when refused, and give the synchronous session's changes and book in the same thread and in a thread, after a thread is lost too, from the journal and from a snapshot taken after them. The seeded sessions of `test/session-fuzz.test.ts` patch cells after each step and compare each sheet's extent with the full computation's cells.
- `test/session-undo-fuzz.test.ts`: seeded sessions make changes of every kind (edits, a table's header typed over, patches, names, styles, layout edits, every structure operation, sheets hidden and shown among them), undo and redo some of them and make new changes, and after each move the book equals the book its place in the history had, the hash equals a full computation's of that book, the history counts the steps each way, a host's screen moved by `mapCell` and given the changes holds the session's values, and a refused change leaves the book and the history as they were; in a synchronous session, in the same thread and in a thread lost now and then with snapshots every 3 calls, with closures and eval. A history that does not put back the style table, the names, the cells a structure operation emptied, a line's entries or a sheet's fields, or a snapshot without the history, makes it fail. `test/session-undo.test.ts` takes each kind of step back and forth in both modes, and `test/session-undo-async.test.ts` compares the same thread and a thread with a synchronous session, stops an undo by its deadline and its signal, loses a thread with undos in the journal and in a snapshot, and has two sessions take turns on a full thread; in a pool of 384 MB and 6 sessions a thread a session keeps 8 MB of steps by the estimate, as a synchronous session with `maxMb: 8`, and one with its own `maxMb` keeps its own.
- `test/async-pool-edges.test.ts`: a book or an edit with a function in `meta` is refused with `InputError` and the thread goes on; a session closed while the pool closes or while its thread is lost is closed; `maxQueue: 0` takes what an idle thread takes at once; `maxHeapMb` with `spawn` is refused.
- `test/limits.test.ts`: the limits in `compute`, in a session and in `createInlineCalc`, with `SEQUENCE`, `MUNIT`, `MMULT`, lifted operators, `TOCOL` of whole columns and spills: `spill` refuses a spill of more stored cells (an old-kind array formula too) but not an array that does not spill, and `A:A*2`, `TOCOL(A:A)` and `IF(A:A>250,A:A,0)` over 500 rows compute under a `spill` of 1,000. `test/async-pool.test.ts`: a pool of 64 MB refuses an array of 1,048,577 values and a spill of 80,000, a pool of 384 MB computes `SUM(SEQUENCE(500000))` and `TOCOL(A:A)`, opens a running total over 10,000 rows and opens and edits running `AVERAGE`, `COUNTIF` and `AVERAGE` with `COUNTA` over 10,000 rows, and a host that raises `cells` keeps the bound on spills.
- `test/locale-pool.test.ts`: a pool started before ru-RU is loaded refuses a ru-RU book with the synchronous engine's `InputError`; once ru-RU is loaded its threads compute such a book and give a session's explanation and a refusal in Russian, as a synchronous session does, on both threads, with closures and eval; a pool started later and a thread started again after a crash have it; a locale outside the package computes in a thread as sent, in English there as here; a locale without its messages survives a structured clone, a locale with them does not.
- `test/worker-bundle.test.ts` bundles `xlcalc` for the browser and finds no file of the threads in it; `test/dist.test.ts` runs the pool of the built package by its name from a plain Node process, which exits with the pool open.
- `test/worker-browser.test.ts` loads the built package in headless Chromium, Firefox and WebKit (through `playwright-core`): the pool on Web Workers gives Node's hashes for the synthetic book and a session's edits, a hard deadline terminates a worker and the session comes back, and the same-thread calculator lets timers run. Each engine is skipped where it is not found.
