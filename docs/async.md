# Asynchronous and threaded computation

Draft of the design, for review before the code. The final page will describe the finished API the way `docs/spill.md` does.

A host that serves several users, or a page that must stay responsive, cannot let one large book hold its thread. The engine gets two asynchronous ways to compute, with one API:

- in the same thread, in portions: the computation yields to the event loop between formulas, stops at a soft deadline and on `AbortSignal`;
- in a pool of threads (Node `worker_threads`, or Web Workers in a browser): the book lives in a thread, a hard deadline terminates the thread, `maxHeapMb` bounds its memory, and a session survives the loss of its thread.

The synchronous `compute` and `openSession` stay as they are.

## API

From `xlcalc`, `createInlineCalc(opts?)`; from the new subpath `xlcalc/worker`, `createCalcPool(opts?)`. Both return a `Calc`:

```ts
interface Calc {
  compute(book: BookInput, opts?: AsyncComputeOptions): Promise<ComputeResult>
  openSession(book: BookInput, opts?: AsyncSessionOptions): Promise<AsyncSession>
  close(): Promise<void>
}

interface AsyncSession {
  readonly clock: Clock                      // as of the last finished call
  setCells(edits: CellEdit[], opts?: CalcOptions): Promise<CellChange[]>
  setNames(names: NameInput[], opts?: CalcOptions): Promise<CellChange[]>
  setClock(clock: Clock, opts?: CalcOptions): Promise<CellChange[]>
  restructure(op: StructureOp, opts?: CalcOptions): Promise<AsyncRestructureResult>
  setStyles(styles: CellStyle[]): Promise<void>
  setLayout(edits: LayoutEdit[], opts?: CalcOptions): Promise<CellChange[]>
  book(): Promise<BookInput>
  explain(sheet: string, cell: string, opts?: ExplainOptions): Promise<Explanation>
  result(opts?: CalcOptions): Promise<ComputeResult>
  close(): Promise<void>
}
```

The methods are those of `Session` and answer the same values. Two differences: `restructure` answers `{ ok, mapCell, changes }` without `book` (a thread would send the whole book back on every operation; `book()` fetches it), and every call returns a promise.

Types added to `src/contract.ts` (one `contract:` commit):

```ts
/** What `signal` needs of an AbortSignal; the platform's AbortSignal fits it. */
interface AbortSignalLike { readonly aborted: boolean; readonly reason?: unknown; addEventListener(type: 'abort', f: () => void, o?: { once?: boolean }): void; removeEventListener(type: 'abort', f: () => void): void }

interface CalcOptions {
  signal?: AbortSignalLike
  /** Soft deadline: milliseconds of computing, counted from when the call starts computing (not while it waits). */
  timeoutMs?: number
}
interface CalcLimits {
  /** Formulas in the book. */
  formulas?: number
  /** Values of one range, array or spill that a formula makes, checked before they are allocated. */
  cells?: number
}
type AsyncComputeOptions = ComputeOptions & CalcOptions
type AsyncSessionOptions = SessionOptions & CalcOptions       // signal and timeoutMs bound the opening
type AsyncRestructureResult = { ok: true; mapCell(ref: CellRef): CellRef | null; changes: CellChange[] } | { ok: false; code: StructureRefusal; message: string }

interface InlineCalcOptions { sliceMs?: number /* 10 */; timeoutMs?: number; limits?: CalcLimits }
interface CalcPoolOptions {
  threads?: number            // 2 by default
  maxHeapMb?: number          // Node: resourceLimits.maxOldGenerationSizeMb of each thread
  hardTimeoutMs?: number      // a task running longer terminates its thread
  timeoutMs?: number          // the soft deadline of every call unless the call gives its own
  maxQueue?: number           // tasks waiting for a thread; one more is refused with queue_full
  limits?: CalcLimits
  snapshotEvery?: number      // journal entries before a session's snapshot is taken again (64)
  sliceMs?: number            // portions inside a thread, so a cancel message gets through (50)
  spawn?: () => WorkerLike    // how to start a thread, for bundlers that move the thread's script
}
type CalcErrorCode = 'soft_timeout' | 'aborted' | 'hard_timeout' | 'out_of_memory' | 'thread_failed' | 'restore_failed' | 'limit' | 'queue_full' | 'closed'
```

`ComputeOptions` and `SessionOptions` also get `limits?: CalcLimits`, so the synchronous API can bound a book too; without it nothing changes.

Errors are one class, `CalcError extends Error`, with `code: CalcErrorCode`, and for `limit` also `limit: 'formulas' | 'cells'`, `max`, `size` and `cell` (the formula's cell). It is exported from `xlcalc` and from `xlcalc/worker`. Wrong input still throws `InputError`, from a thread too (rebuilt on the calling side with its `path`).

## One computation, run in steps

The evaluation loops become generators: `evaluate`, `iterate`, `settle`, `recompute` and `evaluateAll` in `evaluate.ts`, the compile loop of `createState` and `applyEdits`, `applyClock`, `applyMerges`, `applyHidden` in `engine.ts`. A generator yields only between two formulas, and only when a budget is active and its slice is used up. The synchronous functions run the same generator to its end in one `next()` (`createState`, `compute`, `openSession` keep their signatures), so there is one implementation of each step.

`session.ts` keeps its logic in one core whose operations are generators; `openSession` runs each at once, the asynchronous session in portions.

## Checks

One module, `src/exec/budget.ts`, holds the active budget of the computation being run (or none) and two checks:

- `spend(n)`: counts work; every 1,024 units it reads the clock and the signal, and throws `soft_timeout` or `aborted`, or marks the slice used up so the next step boundary yields;
- `allot(n)`: throws `limit` when `n` values exceed `limits.cells`, before the array is allocated.

Without a budget each check is one comparison with `null`. The synchronous path without options runs with none, and the budgets of `test/budgets.test.ts` measure what the steps and checks cost it.

Where the checks go:

| place | check |
|---|---|
| `evaluate`, each formula (iterative calculation and spill rounds go through it) | `spend(1)`, then a yield when the slice is used up |
| the compile loop of `createState`, each formula | `spend(1)`, yield; `limits.formulas` before compiling |
| `applyEdits`, formulas added | `limits.formulas` |
| `rebuildGraph`, each formula | `spend(1)` |
| `flat` and `aligned` in `runtime.ts` (a range read whole: whole columns) | `allot(h·w)`, `spend(h·w)` |
| `liftN` (operators and calls over arrays), `Arr.full` | `allot` before, `spend` each row |
| spill placement and old-kind areas (`spill/place.ts`, `spill/fixed.ts`) | `allot(rows·cols)`, `spend` each row of a large array |
| `buildIndex` and `select` (SUMIFS, COUNTIFS and their kin) | `spend` every 65,536 values |
| functions whose size comes from their arguments: `SEQUENCE`, `RANDARRAY`, `MUNIT`, `EXPAND`, `MMULT`, `WRAPROWS`, `WRAPCOLS` | `allot` |
| matrix functions with cubic loops: `MMULT`, `MINVERSE`, `MDETERM` | `spend` each row |
| `buildResult` and the changes of `restructure` and `setNames` (every cell of a book) | `spend` each cell |

A formula that throws gets its cell written into the error. A single long step without a check (one function looping inside, a structure operation rewriting every formula) cannot be stopped softly: the thread pool's hard deadline covers it.

The limits matter beyond time. Today `SUM(SEQUENCE(20000,20000))` (4·10^8 values) aborts the whole Node process with a heap failure, and so would it in a thread: measured on Node 26, a thread with `maxOldGenerationSizeMb: 256` is terminated cleanly when it fills its heap with many arrays of up to 2·10^7 values, but one array larger than the limit (6·10^7 values) brings the process down. So a pool with `maxHeapMb` and no `limits.cells` takes `limits.cells` of `maxHeapMb · 2^20 / 32` (8 million values for 256 MB).

## In the same thread

`createInlineCalc` runs each call through a driver: it makes the call's budget active, calls `next()` on the call's generator, and between portions yields to the event loop (`setImmediate` in Node, `scheduler.yield` or a `MessageChannel` in a browser, `setTimeout` elsewhere). An abort that comes while the call waits for its turn rejects it at once; one that comes during a yield is thrown into the generator at the point where it stopped.

A session takes its calls in order: each waits in the session's queue until the one before has finished, so `setCells` called twice without waiting gives each promise its own changes, as two synchronous calls would.

A call that stops (soft deadline, abort, limit) changes nothing: the session is as it was before the call. `setNames` and `restructure` build the new state aside, so the old one stays. `setCells`, `setClock` and `setLayout` change the state in place; before them the session notes what they change in the book's input (the cells' old inputs, the sheets' old layout, the clock), puts it back when they stop, and computes the book again on its next call, a cold computation. A full computation of a session's book gives the session's hash (the session tests check this after 200 random edits), so the session after a stopped call has the hash of the session that never got it. The synchronous session does the same when a limit stops it.

## In threads

### Messages

Each thread runs `src/worker/thread.ts`. It holds sessions of the same kind as in the same thread, and runs every task through the same driver with `sliceMs` of 50 ms, so a cancel message is read between portions. The calling side keeps one task in flight per thread.

| to the thread | answer |
|---|---|
| `compute { task, book, opts }` | the result |
| `open { task, sid, book, opts }` | the session's clock |
| `call { task, sid, seq, op, args }` | the value and the session's clock after it |
| `restore { task, sid, snapshot, journal }` | done |
| `snapshot { task, sid }` | `{ book, clock, seq }` |
| `close { task, sid }` | done |
| `cancel { task }` | none; the cancelled task answers `aborted` |

An answer is `{ task, ok: true, value }` or `{ task, ok: false, error }`, the error as plain data (`name`, `message`, `code`, `path`, `limit` and so on) made into `CalcError` or `InputError` again on the calling side. Books, results and changes go by structured clone; nothing else is shared, and no `SharedArrayBuffer` is used. The book goes to a thread once, at `open`; afterwards only edits go there and only changes come back.

`mapCell` of `restructure` is a function and cannot cross threads. The thread answers, beside the changes, an outline of the book before the operation: sheet names, hidden rows of filtered sheets. The calling side applies the operation (`applyOp`) to a book of empty sheets with that outline, which gives the same `mapCell`; a test compares it with the synchronous session's on the structure tests' operations.

### Dispatch

A session is placed on the thread with the fewest sessions and stays there. One-off `compute` calls wait in a shared queue that any idle thread takes from; a thread takes the oldest of its sessions' next call and the shared queue's next task. `maxQueue` counts the tasks waiting, all threads together; restore and snapshot tasks are the pool's own and do not count. An abort removes a waiting task; a running one gets a `cancel` message.

### Deadlines and memory

A task that runs past `hardTimeoutMs` (counted from when it was sent to the thread) is preempted: the thread is terminated, the task rejects with `hard_timeout`, and a new thread takes the old one's place. A thread that reaches `maxHeapMb` is terminated by Node (`ERR_WORKER_OUT_OF_MEMORY`): the running task rejects with `out_of_memory`. A thread that exits or throws outside a task rejects its running task with `thread_failed`. Threads are `unref`ed while idle, so an open pool does not keep the process alive.

### Journal and recovery

The calling side keeps, for each session, a snapshot and a journal:

- the snapshot is the book, the options and the clock at sequence number `seq`: at first the book given to `openSession` (its sheets and cells objects copied, as the session copies them) with the clock the thread answered;
- the journal lists the calls that changed the session and were answered since the snapshot, each with its sequence number, its arguments and the clock the session had after it (a live session takes a new clock on each call; replay reuses the recorded one).

When the journal reaches `snapshotEvery` entries, the pool asks the thread for a snapshot (its `book()`, clock and `seq`) and drops the journal up to it.

When a thread is lost, every session it held is restored on the new thread: `restore` opens the snapshot and replays the journal, before any waiting call of that session. The call that was running is not in the journal, so the session is as it was before that call, as after a soft stop. A thread keeps each session's last `seq` and its last answer: a call with that `seq` again gets the same answer, an older one is refused, so a repeated message cannot apply an edit twice. If a restore fails (the thread is lost again during it, or it throws), that session is closed with `restore_failed` and its waiting calls reject with it; the other sessions are restored on another new thread.

### In a browser

The same pool runs on Web Workers with the same messages: the thread starts with `new Worker(new URL('./thread.js', import.meta.url), { type: 'module' })`, or with `spawn` when a bundler places the script elsewhere. A hard deadline terminates the worker. There is no memory limit: `maxHeapMb` throws `InputError` there. Node or browser is decided when the pool starts (`process.getBuiltinModule('node:worker_threads')`, then the global `Worker`), so the subpath has no `node:` import and a browser bundle of it takes nothing of Node.

The thread's code is not imported by `xlcalc`: a test bundles the root entry and checks that no file of `src/worker/` is in it, as `test/xlsx-bundle.test.ts` does for xlsx.

## Files

- `src/exec/budget.ts`: the budget, `spend`, `allot`, `CalcError`, the synchronous runner of a generator.
- `src/async/`: the driver and the yield, the call queue, the asynchronous session, `createInlineCalc`.
- `src/worker/`: messages, the thread, the pool, the journal, the platforms (Node and Web), the outline for `mapCell`, the subpath's entry.
- Points in `engine.ts`, `evaluate.ts`, `session.ts`, `graph.ts`, `runtime.ts`, `criteria.ts`, `arrays/`, `spill/`, `result.ts` and the functions named above.
- `package.json`: the subpath `./worker`.

## Tests

- Hashes equal in every mode: the synthetic book and the case books computed synchronously, in the same thread and in threads, with closures and eval; a session with a series of edits in the three modes; a session restored after its thread was terminated (by a hard deadline, by `maxHeapMb`, by a crash) against a synchronous session given the same answered calls.
- In the same thread: a timer fires while a large book computes; soft deadline and abort give `CalcError` and leave the session's hash unchanged; calls made without waiting keep their order and their own changes.
- Limits: `limits.formulas`, `limits.cells` in each mode; `SEQUENCE(20000,20000)` gives `limit`, not a crash.
- The pool: `queue_full`, abort while waiting and while running, a snapshot taken and the journal cut, a failed restore, `mapCell` against the synchronous session's.
- The root entry's bundle takes nothing of `src/worker/`.
- A browser: the pool on Web Workers in headless Chromium, through `playwright-core` with a Chromium found on the machine; skipped where none is.

## Open questions

- `limits` in the synchronous `ComputeOptions` and `SessionOptions`: proposed above; the alternative is to keep it for the asynchronous calls only.
- Preemption is read here as the hard deadline terminating a task's thread. If it should also mean a waiting one-off computation giving way to session calls, or the oldest waiting task dropped when the queue is full, it is a small change in dispatch.
- Sessions share threads: a terminated thread costs every session on it a cold computation and a replay. One thread per session is the other choice, at the cost of a Node isolate (about 10 MB) per session.
- The browser test needs `playwright-core` as a development dependency (no browser download; it drives a Chromium already on the machine).
