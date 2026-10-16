/**
 * The asynchronous session in the same thread (`docs/async.md`): the steps of `session.ts` run in portions, its calls
 * one after another. A thread of the pool holds its sessions the same way.
 */
import type { ChangesOptions, PreviewChangesOptions, ReadCellsOptions, FindCellsOptions, DependentsOptions, ReplaceOptions, ExplainOptions } from '../contract.js'
import type { CopyTextOptions } from '../contract.js'
import type { AsyncSession, AsyncSessionOptions, CalcLimits, CalcOptions, ChangingMethod, Clock, EditOptions, SessionHistory, SessionOptions } from '../contract.js'
import { CalcError, type Budget, type Steps } from '../exec/budget.js'
import { openSteps, type SessionHooks } from '../session.js'
import { subscribers, withSubscribe } from '../session/subscribers.js'
import type { HistoryData } from '../session/undo.js'
import { callBudget, drive } from './drive.js'
import { callLimits, checkCall, checkEditCall, splitCall } from './options.js'
import { Serial } from './queue.js'
import { FILL_OPTIONS, FIND_CALL_OPTIONS, READ_OPTIONS, REPLACE_OPTIONS, strictMethods, TEXT_OPTIONS } from '../strict-args.js'

/** What a calculator gives its sessions. */
export interface InlineDefaults {
  sliceMs: number
  timeoutMs?: number
  limits?: CalcLimits
  /** The budgets of the calls running now, which closing the calculator stops. */
  running?: Set<Budget>
}

export interface InlineSession extends AsyncSession {
  /** Stops the session at once: its running call and the waiting ones reject with `why`. */
  kill(why: CalcError): void
  /** What its history holds now, for a pool's snapshot. */
  historyData(): HistoryData
}

/** The calls of `steps` with a budget each, from the calculator's defaults and the call's options. */
export function budgeted(d: InlineDefaults, limits: CalcLimits | undefined) {
  return async <T>(steps: Steps<T>, o: CalcOptions): Promise<T> => {
    const b = callBudget(limits, o.signal, d.sliceMs)
    d.running?.add(b)
    try {
      return await drive(steps, b, o.timeoutMs ?? d.timeoutMs)
    } finally {
      d.running?.delete(b)
    }
  }
}

export async function openInlineSession(book: Parameters<typeof openSteps>[0], opts: AsyncSessionOptions | undefined, d: InlineDefaults, hooks?: SessionHooks): Promise<InlineSession> {
  const { signal, timeoutMs, rest } = splitCall<SessionOptions>(opts)
  const o = { ...rest, limits: callLimits(d.limits, rest?.limits) } as SessionOptions
  const go = budgeted(d, o.limits)
  const core = await go(openSteps(book, o, hooks), { signal, timeoutMs })
  const queue = new Serial()
  let clock: Clock = core.clock
  let history: SessionHistory = core.history
  let revision = core.revision
  const subs = subscribers(revision)
  // close() called: new calls are refused; kill(): the waiting ones too, and the running one stops
  let ended: CalcError | null = null
  let killed: CalcError | null = null
  let running: Budget | null = null
  const turn = <T>(task: () => Promise<T>, signal?: CalcOptions['signal']): Promise<T> => {
    if (ended) return Promise.reject(ended)
    return queue.run(async () => {
      if (killed) throw killed
      return task()
    }, signal)
  }
  /** What the session is after a call: its clock, history and revision. */
  const after = <T>(r: T): T => {
    clock = core.clock
    history = core.history
    revision = core.revision
    return r
  }
  /** A call's steps run in its turn; a changing call (`method`) tells its answer to the subscribers before its promise settles. */
  const steps = <T>(f: () => Steps<T>, c: CalcOptions, method?: ChangingMethod): Promise<T> =>
    turn(async () => {
      const b = callBudget(o.limits, c.signal, d.sliceMs)
      running = b
      d.running?.add(b)
      try {
        const r = after(await drive(f(), b, c.timeoutMs ?? d.timeoutMs))
        return method ? subs.after(method, r) : r
      } finally {
        running = null
        d.running?.delete(b)
      }
    }, c.signal)
  const call = <T>(f: () => Steps<T>, co?: CalcOptions): Promise<T> => {
    try {
      return steps(f, checkCall(co))
    } catch (e) {
      return Promise.reject(e)
    }
  }
  /** A call that takes edit options (`extra` beside `baseRevision`), which go to the session; `method`, one that changes the book. */
  const edit = <T>(f: (eo: EditOptions | undefined) => Steps<T>, co: unknown, extra?: readonly string[], method?: ChangingMethod): Promise<T> => {
    try {
      const [c, eo] = checkEditCall(co, extra)
      return steps(() => f(eo), c, method)
    } catch (e) {
      return Promise.reject(e)
    }
  }
  const plain = <T>(f: () => T): Promise<T> => turn(async () => after(f()))
  return withSubscribe(strictMethods<InlineSession>({
    get clock() {
      return { ...clock }
    },
    get history() {
      return { ...history }
    },
    get revision() {
      return revision
    },
    applyChanges: (ops, co) => edit((eo) => core.applyChanges(ops, eo as ChangesOptions), co, ['clock'], 'applyChanges'),
    previewChanges: (ops, co) => edit((eo) => core.previewChanges(ops, eo as PreviewChangesOptions), co, ['clock', 'maxChanges']),
    readCells: (refs, co) => edit((eo) => core.readCells(refs, eo as ReadCellsOptions), co, READ_OPTIONS),
    findCells: (refs, query, co) => edit((eo) => core.findCells(refs, query, eo as FindCellsOptions), co, FIND_CALL_OPTIONS),
    replaceInCells: (refs, query, replacement, co) => edit((eo) => core.replaceInCells(refs, query, replacement, eo as ReplaceOptions), co, REPLACE_OPTIONS, 'replaceInCells'),
    getDependents: (refs, co) => edit((eo) => core.getDependents(refs, eo as DependentsOptions), co, ['maxCells', 'cursor', 'count', 'maxReferences']),
    summarizeRanges: (refs, co) => edit((eo) => core.summarizeRanges(refs, eo), co, ['rows']),
    setCells: (edits, co) => edit((eo) => core.setCells(edits, eo), co, undefined, 'setCells'),
    replaceCells: (edits, co) => edit((eo) => core.replaceCells(edits, eo), co, undefined, 'replaceCells'),
    clearCells: (ranges, what, co) => edit((eo) => core.clearCells(ranges, what, eo), co, undefined, 'clearCells'),
    patchCells: (patches, co) => edit((eo) => core.patchCells(patches, eo), co, undefined, 'patchCells'),
    setCellText: (sheet, cell, text, co) => edit((eo) => core.setCellText(sheet, cell, text, eo), co, TEXT_OPTIONS, 'setCellText'),
    previewCellText: (sheet, cell, text, co) => edit((eo) => core.previewCellText(sheet, cell, text, eo), co, TEXT_OPTIONS),
    setCellTexts: (entries, co) => edit((eo) => core.setCellTexts(entries, eo), co, TEXT_OPTIONS, 'setCellTexts'),
    previewCellTexts: (entries, co) => edit((eo) => core.previewCellTexts(entries, eo), co, TEXT_OPTIONS),
    pasteText: (sheet, cell, text, co) => edit((eo) => core.pasteText(sheet, cell, text, eo), co, undefined, 'pasteText'),
    previewPasteText: (sheet, cell, text, co) => edit((eo) => core.previewPasteText(sheet, cell, text, eo), co),
    fillCellText: (sheet, range, text, co) => edit((eo) => core.fillCellText(sheet, range, text, eo), co, FILL_OPTIONS, 'fillCellText'),
    previewFillCellText: (sheet, range, text, co) => edit((eo) => core.previewFillCellText(sheet, range, text, eo), co, FILL_OPTIONS),
    setNames: (names, co) => edit((eo) => core.setNames(names, eo), co, undefined, 'setNames'),
    setClock: (next, co) => edit((eo) => core.setClock(next, eo), co, undefined, 'setClock'),
    editStructure: (op, co) => edit((eo) => core.editStructure(op, eo), co, undefined, 'editStructure'),
    // eslint-disable-next-line require-yield
    setStyles: (styles, co) => edit(function* (eo) {
      return core.setStyles(styles, eo)
    }, co, undefined, 'setStyles'),
    setLayout: (edits, co) => edit((eo) => core.setLayout(edits, eo), co, undefined, 'setLayout'),
    getBook: () => plain(() => core.getBook()),
    getLayout: () => plain(() => core.getLayout()),
    getUsedRange: (sheet) => call(() => core.getUsedRange(sheet)),
    getEnd: (sheet, cell, direction) => call(() => core.getEnd(sheet, cell, direction)),
    explain: (sheet, cell, eo) => {
      try {
        const { signal, timeoutMs, rest } = splitCall<ExplainOptions>(eo)
        return call(() => core.explain(sheet, cell, rest), { signal, timeoutMs })
      } catch (e) {
        return Promise.reject(e)
      }
    },
    getCells: (sheet, range, co) => call(() => core.getCells(sheet, range), co),
    copyText: (sheet, range, o) => {
      try {
        const { signal, timeoutMs, rest } = splitCall<CopyTextOptions>(o)
        return call(() => core.copyText(sheet, range, rest), { signal, timeoutMs })
      } catch (e) {
        return Promise.reject(e)
      }
    },
    getCell: (sheet, cell) => call(() => core.getCell(sheet, cell)),
    getViewport: (sheet, range, vo) => call(() => core.getViewport(sheet, range, vo)),
    getCellIssue: (sheet, cell) => call(() => core.getCellIssue(sheet, cell)),
    getValidation: (sheet, cell) => call(() => core.getValidation(sheet, cell)),
    validateEntry: (sheet, cell, input, co) => call(() => core.validateEntry(sheet, cell, input), co),
    getResult: (co) => call(() => core.getResult(), co),
    getSnapshot: (co) => call(() => core.getSnapshot(), co),
    getHash: (co) => call(() => core.getHash(), co),
    undo: (co) => edit((eo) => core.undo(eo), co, undefined, 'undo'),
    redo: (co) => edit((eo) => core.redo(eo), co, undefined, 'redo'),
    historyData: () => core.historyData(),
    subscribe: (listener) => subs.subscribe(listener),
    close() {
      subs.clear()
      if (ended) return Promise.resolve()
      // closing waits for the calls made before it, and closes the session whatever stops them
      const done = queue.run(async () => core.close())
      ended = new CalcError('closed', 'the session is closed')
      return done
    },
    kill(why: CalcError) {
      subs.clear()
      ended ??= why
      killed = why
      if (running) running.halted = why
      core.close()
    },
  }, true), subs)
}
