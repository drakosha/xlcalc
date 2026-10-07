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

## Deadlines and cancellation

`timeoutMs` of a call is a soft deadline: milliseconds of computing, counted from when the call starts computing, not while it waits for its turn. `signal` stops a call at any time: a call that waits is refused at once, a running one at its next check. For a deadline from the moment of the call, pass `AbortSignal.timeout(ms)`.

The engine checks the deadline and the signal between formulas and inside the long loops: a range read whole (whole columns), operators and functions over arrays, a spill and its cells, the indexes and the selection of `SUMIFS`, `COUNTIFS` and their kin, the matrix functions, iterative calculation, the result's cells. One long step without a check, such as a structure operation rewriting every formula, finishes before the call can stop. In a pool, `hardTimeoutMs` covers that: a task that runs longer, counted from when it reaches its thread, terminates the thread.

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

Wrong input still throws `InputError`, from a thread too.

## In the same thread

`createInlineCalc({ sliceMs, timeoutMs, limits })` computes in slices of `sliceMs` milliseconds (10 by default) and gives way to the event loop between them: through `setImmediate` in Node and a `MessageChannel` message in a browser, so timers, input and other messages run while a book computes. `timeoutMs` and `limits` are the defaults of every call. Several sessions and computations of one calculator share the thread, slice by slice.

`close()` stops what runs and waits with `closed`.

## In threads

`createCalcPool(opts)` starts `threads` threads (2 by default). A session's book goes to its thread once, when the session opens; afterwards only edits go there and only changes come back.

- A session is placed on the thread that holds the fewest sessions and stays there. One-off computations wait in one queue for any thread. A thread that becomes idle takes the next call of its own sessions first, then the next one-off computation.
- `maxQueue` bounds the tasks that wait for a thread (`queue_full`), `maxSessions` the open sessions of the pool (`too_many_sessions`).
- `maxSessionsPerThread` bounds the sessions a thread holds at once. Past it, the session called least recently leaves the thread: the thread forgets its book, and the session comes back on its next call (see below). Sessions that take turns on a full thread come back again and again; give the pool enough room for the sessions in use.
- `hardTimeoutMs` terminates a thread whose task runs longer. `maxHeapMb` is each thread's heap in Node (`resourceLimits.maxOldGenerationSizeMb`); a thread that fills it is terminated by Node and its task fails with `out_of_memory`. Node terminates a thread cleanly when many arrays fill its heap, but one array larger than the heap brings the whole process down, so a pool with `maxHeapMb` and without `limits.cells` takes `cells` of `maxHeapMb · 2^20 / 32` (8 million values for 256 MB).
- A thread that waits keeps no process alive: an open pool does not stop Node from exiting. A Node thread starts without the process's own flags (`execArgv`), which are meant for the host's entry.

Live sessions and heavy one-off computations (building versions of a book, say) are best kept in two pools, so that a long computation does not hold a thread the sessions need.

### When a thread is lost

The pool keeps, for each session, a snapshot of its book and a journal of the calls that changed it since the snapshot, with the clock the session had after each. The first snapshot is the book given to `openSession`; when the journal reaches `snapshotEvery` calls (64 by default), the pool takes a new one from the thread and drops the journal up to it.

When a thread is terminated or fails, its task fails, and a new thread starts in its place when a task needs one; a thread that cannot start fails one task at a time. The sessions it held come back on their next call: the new thread opens the snapshot and replays the journal, with the recorded clocks (a live session too), and then takes the call. The call that was running is not in the journal, so the session is as it was before that call, with the hash of a session that never lost its thread. Each changing call carries a sequence number; a thread answers the last one again rather than applying it twice, and refuses one out of order. A session whose restore fails is closed with `restore_failed`; the other sessions of the thread go on.

This has a cost on the calling side: it holds a copy of each session's book. On the synthetic book of 3,400 rows (54,778 cells) a snapshot from a thread takes 40 ms and its copy about 9 MB; on one of 34,000 rows (510,682 cells), 541 ms and 61 MB.

### In a browser

The same pool runs on Web Workers with the same messages, started as `new Worker(new URL('./thread.js', import.meta.url), { type: 'module' })`, the form bundlers recognise. Where a bundler puts the thread's script elsewhere, `spawn` starts it. A hard deadline terminates the worker. There is no memory limit: `maxHeapMb` throws `InputError` there.

`xlcalc` imports nothing of the threads: a page that computes in its own thread does not load them.

## What it costs the synchronous path

The evaluation loops are generators that yield only when a calculator's slice is used up, and each check is one comparison when no calculator runs. On the full synthetic book (3,907 formulas), against the engine before them, five runs each taken in turns, medians: a cold computation 322 and 327 ms with closures, 341 and 336 ms with eval; an edit of a criterion cell 1.9 and 1.8 ms at p95 with closures, 1.8 and 1.9 ms with eval. The values hash is the same.

## How it was checked

- `test/async-inline.test.ts` and `test/async-pool.test.ts` compute the coverage measurement (in en-US and ru-RU), the semantics cases, the synthetic book, the spill book, a seeded book, iterative calculation and the 1904 date system in the same thread and in threads, with closures and eval, and compare values, issues and hashes with the synchronous engine; a session over 40 edits gives the changes and the hash of a synchronous session after each, in both execution modes and in both calculators; two sessions on one thread with their calls mixed without waiting each give their own changes and hash; a structure operation gives the synchronous session's changes and `mapCell` on eleven operations, filtered rows among them.
- In the same thread a timer runs while a book computes; deadlines, aborts (waiting and running) and limits give `CalcError`; a stopped `setCells`, `setLayout` or `restructure` leaves the book and the hash as they were.
- `test/async-recovery.test.ts` terminates threads by a hard deadline, by `maxHeapMb` (twenty spilled arrays of a million values in a 64 MB heap) and by a crash, and checks that the session's hash after it comes back equals that of a synchronous session given the same answered calls, from a snapshot and a journal cut at it too; a restore that fails closes that session only; sessions take turns on a full thread; `maxQueue`, `maxSessions`, the order of a thread's tasks, closing.
- `test/limits.test.ts`: the limits in `compute` and in a session, with `SEQUENCE`, `MUNIT`, `MMULT`, lifted operators, `TOCOL` of whole columns and a spill.
- `test/worker-bundle.test.ts` bundles `xlcalc` for the browser and finds no file of the threads in it; `test/dist.test.ts` runs the pool of the built package by its name from a plain Node process, which exits with the pool open.
- `test/worker-browser.test.ts` loads the built package in headless Chromium (through `playwright-core`): the pool on Web Workers gives Node's hashes for the synthetic book and a session's edits, a hard deadline terminates a worker and the session comes back, and the same-thread calculator lets timers run. It is skipped where no Chromium is found.
