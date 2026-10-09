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

`cells(sheet, range)` answers the cells of one rectangle as `result()` gives them, with each cell's own `style`, so a window over a sheet needs no copy of the whole result. From a thread only the rectangle comes back: 20,000 cells of the synthetic book of 34,000 rows take 27 ms there and 16 ms in a synchronous session, where the whole `result()` from the thread takes 1.7 s.

A call that stopped while it recomputed (`setCells`, `setClock` or `setLayout` past its deadline, aborted, or over the limits) leaves the session to compute its book again, and the session's next call does that first. That computation does not count in the next call's `timeoutMs`, only its `signal` and a pool's hard deadline bound it, so a call after a stopped one can take longer than its `timeoutMs` by the time of a cold computation of the book.

## Deadlines and cancellation

`timeoutMs` of a call is a soft deadline: milliseconds of computing, counted from when the call starts computing, not while it waits for its turn. `signal` stops a call at any time: a call that waits is refused at once, a running one at its next check. For a deadline from the moment of the call, pass `AbortSignal.timeout(ms)`.

The engine checks the deadline and the signal before each formula and inside the long loops: a range read whole (whole columns), operators and functions over arrays, a spill and its cells, the indexes and the selection of `SUMIFS`, `COUNTIFS` and their kin, the matrix functions, iterative calculation, the result's cells. A slice can run past `sliceMs` by the time of the formula in flight when it ends. One long step without a check, such as a structure operation rewriting every formula, finishes before the call can stop. In a pool, `hardTimeoutMs` covers that: a task that runs longer, counted from when it reaches its thread, terminates the thread.

A call that stops by its deadline or its signal throws `CalcError` with `soft_timeout` or `aborted` (the signal's reason is the error's `cause`).

## Limits

`limits` bound what a book may make: `formulas`, the formulas of the book, and `cells`, the values of one range read whole, one array or one spill. They are options of `compute` and `openSession` too, not only of the calculators.

```js
compute(book, { limits: { formulas: 100_000, cells: 5_000_000 } })
```

A book over a limit is not computed: the call throws `CalcError` with code `limit`, `limit` (`formulas` or `cells`), `max`, `size` and, for `cells`, the formula's `cell`. A limit is no cell error such as `#NUM!`, because then the values and their hash would depend on the limits; a result is either exact or absent. Arrays are checked before they are made: without `cells`, `SUM(SEQUENCE(20000,20000))` asks for 400 million values and brings a Node process down.

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
- `hardTimeoutMs` terminates a thread whose task runs longer. Bringing a session back opens its book and replays its journal, each of which was answered within the deadline once, so a restore has the deadline of `1 +` the calls it replays. `maxHeapMb` is each thread's heap in Node (`resourceLimits.maxOldGenerationSizeMb`); a thread that fills it is terminated by Node and its task fails with `out_of_memory`. Node terminates a thread cleanly when many small objects fill its heap, but one allocation larger than what the heap has left brings the whole process down: a large array, or the table of a large spill's cells while its result, hash or changes are made. So a pool with `maxHeapMb` and without `limits.cells` takes `cells` of `maxHeapMb · 2^20 / 1024`, a value per KB of heap (262,144 values for 256 MB). The divisor is measured on Node 24 and 26 with heaps of 256 to 1024 MB: a spill read by `compute` brought the process down from 300 bytes of heap per value (2.8 million values in 800 MB), other arrays from 30 bytes, so the default stays three times below the worst.
- A thread that waits keeps no process alive: an open pool does not stop Node from exiting. A Node thread starts without the process's own flags (`execArgv`), which are meant for the host's entry.

Live sessions and heavy one-off computations (building versions of a book, say) are best kept in two pools, so that a long computation does not hold a thread the sessions need.

### When a thread is lost

The pool keeps, for each session, a snapshot of its book and a journal of the calls that changed it since the snapshot, with the clock the session had after each. The first snapshot is the book given to `openSession`; when the journal reaches `snapshotEvery` calls (64 by default), the pool takes a new one from the thread and drops the journal up to it.

When a thread is terminated or fails, its task fails, and a new thread starts in its place when a task needs one; a thread that cannot start fails one task at a time. The sessions it held come back on their next call: the new thread opens the snapshot and replays the journal, with the recorded clocks (a live session too), and then takes the call. The call that was running is not in the journal, so the session is as it was before that call, with the hash of a session that never lost its thread. Each changing call carries a sequence number; a thread answers the last one again rather than applying it twice, and refuses one out of order. A session whose restore fails is closed with `restore_failed`, and the thread forgets what it had of it; the other sessions of the thread go on. A session closed while its thread is lost is closed all the same, and so is one whose `close()` waits when the pool closes.

This has a cost on the calling side: it holds each session's book as text. On the synthetic book of 34,000 rows the text takes 29 MB (the book's objects took 68 MB), a new snapshot from the thread stops the event loop for 3 ms (it was 234 to 290 ms), and `book()` takes 450 ms read a chunk at a time, stopping the event loop for 20 ms at most (it was 390 ms in one stretch of 230 ms).

### In a browser

The same pool runs on Web Workers with the same messages, started as `new Worker(new URL('./thread.js', import.meta.url), { type: 'module' })`, the form bundlers recognise. Where a bundler puts the thread's script elsewhere, `spawn` starts it. A hard deadline terminates the worker. There is no memory limit: `maxHeapMb` throws `InputError` there.

`xlcalc` imports nothing of the threads: a page that computes in its own thread does not load them.

The pool is checked in headless Chromium, Firefox and WebKit through Playwright, and gives the hashes of Node in each. WebKit stands there for Safari, which Playwright does not drive.

## What it costs the synchronous path

The evaluation loops are generators that yield only when a calculator's slice is used up, and each check is one comparison when no calculator runs. On the full synthetic book (3,907 formulas), against the engine before them, five runs each taken in turns, medians: a cold computation 322 and 327 ms with closures, 341 and 336 ms with eval; an edit of a criterion cell 1.9 and 1.8 ms at p95 with closures, 1.8 and 1.9 ms with eval. The values hash is the same.

## How it was checked

- `test/async-inline.test.ts` and `test/async-pool.test.ts` compute the coverage measurement (in en-US and ru-RU), the semantics cases, the synthetic book, the spill book, a seeded book, iterative calculation and the 1904 date system in the same thread and in threads, with closures and eval, and compare values, issues and hashes with the synchronous engine; a session over 40 edits gives the changes and the hash of a synchronous session after each, in both execution modes and in both calculators; two sessions on one thread with their calls mixed without waiting each give their own changes and hash; a structure operation gives the synchronous session's changes and `mapCell` on eleven operations, filtered rows among them.
- In the same thread a timer runs while a book computes; deadlines, aborts (waiting and running) and limits give `CalcError`; a stopped `setCells`, `setLayout` or `restructure` leaves the book and the hash as they were; 300 formulas that each sort 10,000 values stop at a deadline of 100 ms within 200 ms, the event loop never waiting 50 ms (measured alone: 100 to 102 ms, waits of 10 to 21 ms with slices of 10 ms).
- `test/async-recovery.test.ts` terminates threads by a hard deadline, by `maxHeapMb` (twenty spilled arrays of a million values in a 64 MB heap) and by a crash, and checks that the session's hash after it comes back equals that of a synchronous session given the same answered calls, from a snapshot and a journal cut at it too; a restore that fails closes that session only, and ten of them on a 16 MB book leave a 96 MB thread alive; a session of ten calls of 150 ms comes back within a deadline of 800 ms after it left a full thread and after a crash; sessions take turns on a full thread; `maxQueue`, `maxSessions`, the order of a thread's tasks, closing.
- `test/async-pool-edges.test.ts`: a book or an edit with a function in `meta` is refused with `InputError` and the thread goes on; a session closed while the pool closes or while its thread is lost is closed; `maxQueue: 0` takes what an idle thread takes at once; `maxHeapMb` with `spawn` is refused.
- `test/limits.test.ts`: the limits in `compute` and in a session, with `SEQUENCE`, `MUNIT`, `MMULT`, lifted operators, `TOCOL` of whole columns and a spill.
- `test/worker-bundle.test.ts` bundles `xlcalc` for the browser and finds no file of the threads in it; `test/dist.test.ts` runs the pool of the built package by its name from a plain Node process, which exits with the pool open.
- `test/worker-browser.test.ts` loads the built package in headless Chromium, Firefox and WebKit (through `playwright-core`): the pool on Web Workers gives Node's hashes for the synthetic book and a session's edits, a hard deadline terminates a worker and the session comes back, and the same-thread calculator lets timers run. Each engine is skipped where it is not found.
