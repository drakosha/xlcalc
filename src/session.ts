/**
 * A session over a book: edits are applied atomically with one recomputation of dependents. A live
 * session (`live`) takes a fresh clock and a new seed on every recomputation; its current clock
 * (`clock`) is what a saved result records, and a full computation with it gives the same hash.
 *
 * The operations are steps (`openSteps`): `createSession` runs each to its end at once, the asynchronous session in
 * portions. A call that stops (a limit, a deadline, an abort) changes nothing: what it changed in place is put back,
 * and the state is built again from the book on the next call. Each call that changes the book answers an
 * `EditResult` (`session/answer.ts`) and adds one to the session's revision.
 */
import type { BookInput, BookLayout, BookSnapshot, CellEdit, CellInput, CellIssueDetail, CellPatch, CellReplace, CellStyle, CellTextEntry, CellTextOptions, CellTextPreview, CellTextResult, CellTextsPreview, CellTextsResult, CellValidation, CellView, ClearKind, Clock, ColumnWidth, ComputeResult, EditOptions, EditResult, EntryCheck, Explanation, ExplainOptions, LayoutEdit, NameInput, NumberFormat, RangeRef, Scalar, Session, SessionHistory, SessionOptions, SessionStructureResult, SheetExtent, SheetInput, StructureOp, UndoResult, Viewport, ViewportOptions } from './contract.js'
import { applyClock, applyEdits, applyHidden, applyMerges, buildState, setNow, type EngineState } from './engine.js'
import { limitsBudget, run, untimed, withBudget, type Steps } from './exec/budget.js'
import { checkExplainOptions, explainCell, explainText } from './explain.js'
import { currentClock, nowMs } from './env.js'
import { checkAddress, checkBook, checkCell, checkClock, checkEdits, checkLayoutEdits, checkNames, checkSessionOptions, checkSheet, checkString, checkStyles } from './input.js'
import { checkBase, editFields, inOrder, joined, layoutChanged, nothing, repaintOf, ruleRanges, sortedKeys, type FieldChange, type Outcome } from './session/answer.js'
import { applyLayout, bookLayout, checkStylesUsed, contentEdits, inputsOf } from './session/book.js'
import { checkFieldEdits, resolveEdits } from './session/cells.js'
import { clearPlan } from './session/clear.js'
import { checkArrayEdits } from './spill/fixed.js'
import { sheetExtent } from './session/extent.js'
import { checkRange, checkViewportOptions, viewport } from './session/viewport.js'
import { cellValue, editedCells, layoutCells, shownChanges, snapshot } from './session/looks.js'
import { movedChanges } from './session/moved.js'
import { checkPatches } from './session/patch.js'
import { patchPlan, type Plan } from './session/styles.js'
import { bookDiff, DiffPass, type BookDiff } from './session/rebook.js'
import { columnWidening, entryCheck, validationAlert } from './session/text.js'
import { checkTextEntries, textAlerts, textsFound, textWidths, typedTexts, type Typed } from './session/texts.js'
import { applyStep, bookStep, cellsStep, History, layoutStep, manyStep, stylesStep, type HistoryData, type SessionCore, type Step, type Taken } from './session/undo.js'
import { headerEdits } from './tables/headers.js'
import { checkTables } from './tables/check.js'
import { issueDetail } from './issues.js'
import { dropStats } from './cond/stats.js'
import { fitEdits } from './layout/fit.js'
import { hiddenRows } from './layout/hidden.js'
import { parseMerge } from './layout/merges.js'
import { formatAddress, parseAddress } from './model.js'
import { buildResult, cellsIn, clockSerial, spillFields, stateHash } from './result.js'
import { formatCodeIn } from './locale/rules.js'
import { cellLooks } from './shown.js'
import { applyOp, checkedMapCell } from './structure/index.js'
import { inverse, movesOf, type Geometry } from './structure/geometry.js'
import { stateValues } from './structure/values.js'
import { sheetKey } from './syntax.js'
import { cellValidation, hasRule, withInput, withInputs } from './validation/session.js'
import { cellTaken, editsTaken, opTaken } from './taken.js'
import { strictMethods } from './strict-args.js'

/** How a session is run by its host. */
export interface SessionHooks {
  /** The clock a live session takes at each recomputation (a thread replaying a journal gives the recorded ones). */
  clock?: () => Clock
  /** The history a session restored from a snapshot starts with. */
  history?: HistoryData
  /** The revision a session restored from a snapshot starts at. */
  revision?: number
}

/** A session's operations as steps; `getBook`, `getLayout`, `setStyles` and `close` compute nothing. */
export interface SessionSteps {
  readonly clock: Clock
  readonly revision: number
  setCells(edits: CellEdit[], opts?: EditOptions): Steps<EditResult>
  replaceCells(edits: CellReplace[], opts?: EditOptions): Steps<EditResult>
  clearCells(ranges: RangeRef[], what: ClearKind, opts?: EditOptions): Steps<EditResult>
  patchCells(patches: CellPatch[], opts?: EditOptions): Steps<EditResult>
  setCellText(sheet: string, cell: string, text: string, opts?: CellTextOptions): Steps<CellTextResult>
  previewCellText(sheet: string, cell: string, text: string, opts?: CellTextOptions): Steps<CellTextPreview>
  setCellTexts(entries: CellTextEntry[], opts?: CellTextOptions): Steps<CellTextsResult>
  previewCellTexts(entries: CellTextEntry[], opts?: CellTextOptions): Steps<CellTextsPreview>
  setNames(names: NameInput[], opts?: EditOptions): Steps<EditResult>
  setClock(clock: Clock, opts?: EditOptions): Steps<EditResult>
  editStructure(op: StructureOp, opts?: EditOptions): Steps<SessionStructureResult>
  setStyles(styles: CellStyle[], opts?: EditOptions): EditResult
  setLayout(edits: LayoutEdit[], opts?: EditOptions): Steps<EditResult>
  getBook(): BookInput
  getLayout(): BookLayout
  getUsedRange(sheet: string): Steps<SheetExtent>
  undo(opts?: EditOptions): Steps<UndoResult | null>
  redo(opts?: EditOptions): Steps<UndoResult | null>
  readonly history: SessionHistory
  /** What the history holds, for a snapshot. */
  historyData(): HistoryData
  explain(sheet: string, cell: string, opts?: ExplainOptions): Steps<Explanation>
  getCells(sheet: string, range: string | string[]): Steps<Record<string, CellView>>
  getCell(sheet: string, cell: string): Steps<CellView | undefined>
  getViewport(sheet: string, range: string, opts?: ViewportOptions): Steps<Viewport>
  getCellIssue(sheet: string, cell: string): Steps<CellIssueDetail | null>
  getValidation(sheet: string, cell: string): Steps<CellValidation | null>
  validateEntry(sheet: string, cell: string, input: CellInput | null): Steps<EntryCheck>
  getResult(): Steps<ComputeResult>
  getSnapshot(): Steps<BookSnapshot>
  getHash(): Steps<string>
  close(): void
}

/**
 * The name a number or a boolean typed into a header cell gives its column: as the cell's own number format shows it,
 * without the column's width (`1,50` by `0.00`, `31.01.2025` by the short date in ru-RU), or as `&` writes it where the
 * format shows it as General does (`0,000000001`, not `1E-09`): the set `tables/headers-ru-held`.
 */
function headerText(s: EngineState, v: Scalar, format: NumberFormat | undefined): string {
  const plain = String(s.ctx.rules.toText(v))
  if (format === undefined || typeof v !== 'number') return plain
  const shown = s.ctx.rules.display(v, format, Infinity).text
  return shown === s.ctx.rules.display(v, undefined, Infinity).text ? plain : shown
}

/** Opens a session in steps: the book checked, then computed. */
export function* openSteps(book: BookInput, opts?: SessionOptions, hooks: SessionHooks = {}): Steps<SessionSteps> {
  const o = checkSessionOptions(opts)
  book = checkBook(book)
  const live = o.live === true
  const fresh = hooks.clock ?? currentClock
  let clock: Clock = 'clock' in o && o.clock ? { ...o.clock } : fresh()
  const { externals, execution } = o
  const display = o.display !== false
  const build = (b: BookInput) => buildState(b, { now: clockSerial(clock), seed: clock.seed, externals, execution })
  let state: EngineState | null = yield* build(book)
  let revision = hooks.revision ?? 0
  // a call stopped while changing the state in place: the state is built again from its book before the next one
  let stale = false
  const st = () => {
    if (!state) throw new Error('session is closed')
    return state
  }
  function* ready(): Steps<EngineState> {
    const s = st()
    if (stale) {
      // what an earlier call left to do does not count in this call's soft deadline
      state = yield* untimed(build(s.input))
      stale = false
    }
    return state!
  }
  const sheetIndex = (s: EngineState) => (name: string) => s.sheetIdx.get(sheetKey(name))
  const styleCount = (s: EngineState) => s.input.styles?.length ?? 0
  /** The changes of cells `keys` as a host gets them, with the fields that changed of some and a new format's code. */
  const changes = (s: EngineState, keys: string[], fields: Map<string, FieldChange>) => {
    // the statistics of conditional formats over changed cells are taken again
    dropStats(s, keys)
    const look = display ? cellLooks(s) : null
    const code = formatCodeIn(s.input.locale)
    return keys.map((k) => {
      const [i, r, c] = k.split(':').map(Number)
      const value = cellValue(s, i, r, c)
      const f = fields.get(k)
      const format = f?.numberFormat
      return { sheet: s.sheetNames[i], cell: formatAddress(r, c), value, ...look?.(i, r, c, value), ...spillFields(s, i, r, c), ...f, ...(format !== undefined && format !== null ? { numberFormatCode: code(format) } : {}) }
    })
  }
  /** A call that changed the book answered: the revision one more, and what changed. */
  const answer = (s: EngineState, x: Outcome): EditResult => {
    const out: EditResult = { revision: ++revision, changes: changes(s, x.keys, x.fields) }
    if (x.layout.size) out.layout = [...x.layout]
    if (x.styles) out.styles = s.input.styles ?? []
    if (x.moved.length) out.moved = x.moved
    const repaint = repaintOf(s.input, x)
    if (repaint) out.repaint = repaint
    return out
  }
  const tick = (s: EngineState) => {
    clock = fresh()
    setNow(s, clockSerial(clock), clock.seed)
  }
  // a call that builds a new state aside: when it stops, the old state gets its clock back
  const untick = (s: EngineState, at: Clock) => {
    if (clock === at) return
    clock = at
    setNow(s, clockSerial(at), at.seed)
  }
  const history = new History(o.undo, hooks.history)
  // a step is made only when the history keeps steps (a book's diff costs a pass over its cells); a call that changed nothing makes none
  const keep = (step: () => Step | null) => {
    const x = history.on ? step() : null
    if (x) history.push(x)
  }
  /** Edits of cells (checked) applied; the inputs they replaced, which are put back if the call stops. */
  function* putCells(s: EngineState, edits: CellReplace[]): Steps<[Outcome, CellReplace[]]> {
    const before = display ? snapshot(s, editedCells(s, edits)) : null
    const was = inputsOf(s, edits)
    const at = clock
    try {
      // edits that keep a cell's value and formula change its input alone
      const loud = contentEdits(s, edits)
      if (loud.length && live) tick(s)
      const keys = loud.length ? yield* applyEdits(s, loud, live) : []
      const fields = editFields(s, was)
      return [{ ...nothing(), keys: inOrder(keys, [...(before ? shownChanges(s, before) : []), ...fields.keys()]), fields, valued: keys.length > 0 }, was]
    } catch (e) {
      for (const { sheet, cell, input } of was) {
        const cells = s.input.sheets[sheetIndex(s)(sheet)!].cells
        if (input === null) delete cells[cell]
        else cells[cell] = input
      }
      clock = at
      stale = true
      throw e
    }
  }
  /**
   * The state built again from `book`, the session's book changed, with its changes against the old one by the geometry
   * `g` (where the old cells went; the looks of both with `looks`) and, with `back`, the way back to the old book, taken
   * in the same pass over the cells; nothing changes if it stops.
   */
  function* rebuild(old: EngineState, book: BookInput, g: Geometry, looks: boolean, back: boolean): Steps<[EngineState, Outcome, BookDiff | null]> {
    const at = clock
    if (live) tick(old)
    try {
      const next = yield* build(book)
      const fields = new Map<string, FieldChange>()
      const pass = back ? new DiffPass(book, old.input) : undefined
      const keys = movedChanges(old, next, g, fields, display && looks ? [cellLooks(old), cellLooks(next)] : undefined, pass)
      const diff = pass ? yield* bookDiff(pass, inverse(g)) : null
      state = next
      return [next, { ...nothing(), keys: sortedKeys([...keys, ...fields.keys()]), fields, valued: true, layout: new Set(layoutChanged(old.input, next.input)) }, diff]
    } catch (e) {
      untick(old, at)
      throw e
    }
  }
  /**
   * Sheets of the book replaced by new ones of another layout: the arrays and SUBTOTALs it concerns computed again; the
   * changes and the sheets before, which are put back if the call stops.
   */
  function* relayout(s: EngineState, next: Map<number, SheetInput>, shown: Map<string, string> | null): Steps<[Outcome, SheetInput[]]> {
    const sheets = s.input.sheets.slice()
    for (const [i, x] of next) s.input.sheets[i] = x
    const ruled = [...next].filter(([i, x]) => x.conditionalFormats !== sheets[i].conditionalFormats)
    if (ruled.length) dropStats(s, [], true)
    try {
      // merged areas added or removed: the arrays whose range they meet spill or stop; hidden rows: SUBTOTAL and AGGREGATE count again
      const keys = new Set<string>()
      for (const [i, x] of next) {
        const [was, now] = [new Set(sheets[i].merges ?? []), new Set(x.merges ?? [])]
        const moved = [...was].filter((m) => !now.has(m)).concat([...now].filter((m) => !was.has(m)))
        if (moved.length) for (const k of yield* applyMerges(s, i, moved.map((m) => parseMerge(m)!))) keys.add(k)
        if (JSON.stringify(hiddenRows(x) ?? null) !== JSON.stringify(hiddenRows(sheets[i]) ?? null)) for (const k of yield* applyHidden(s, i)) keys.add(k)
      }
      const rules = ruled.flatMap(([i, x]) => [...ruleRanges(sheets[i]), ...ruleRanges(x)])
      return [{ ...nothing(), keys: inOrder([...keys], shown ? shownChanges(s, shown) : []), valued: keys.size > 0, layout: new Set([...next.keys()].map((i) => s.sheetNames[i])), rules }, sheets]
    } catch (e) {
      s.input.sheets.splice(0, sheets.length, ...sheets)
      stale = true
      throw e
    }
  }
  const core: SessionCore = { ready, putCells, rebuild, relayout, display, sheetIndex }
  /**
   * Cells' new inputs (checked) written, as `call`: a table's header cell typed over renames its column and the
   * references to it, which builds the book again; the state after and what changed, and the way back.
   */
  function* writeCells(s: EngineState, edits: CellReplace[], call: Step['call']): Steps<[EngineState, Outcome, () => Step | null]> {
    checkArrayEdits(s, edits)
    const renamed = headerEdits(s.input, edits, (v, format) => headerText(s, v, format))
    if (renamed) {
      const [next, out, diff] = yield* rebuild(s, renamed, [], true, history.on)
      return [next, out, () => diff && bookStep(call, diff, renamed)]
    }
    const [out, was] = yield* putCells(s, edits)
    return [s, out, () => (was.length ? cellsStep(call, was) : null)]
  }
  /**
   * A plan of looks applied after the cells it writes (`write`, which alone may recompute and stop): the style table and
   * the sheets' lines and rules, which compute nothing. The state after, what changed and the steps to take it back.
   */
  function* applyPlan(s: EngineState, plan: Plan, call: Step['call'], write: (s: EngineState, edits: CellReplace[]) => Steps<[EngineState, Outcome, () => Step | null]>): Steps<[EngineState, Outcome, (() => Step | null)[]]> {
    const [next, out, cellsBack] = yield* write(s, plan.edits)
    let o = out
    const steps = [cellsBack]
    if (plan.sheets.size) {
      const shown = display ? snapshot(next, layoutCells(next, plan.touched)) : null
      const [lo, before] = yield* relayout(next, plan.sheets, shown)
      o = joined(o, lo)
      steps.push(() => layoutStep(call, before, plan.sheets))
    }
    if (plan.styles) {
      const was = next.input.styles
      next.input.styles = plan.styles
      o = { ...o, styles: true }
      steps.push(() => stylesStep(was, call))
    }
    return [next, o, steps]
  }
  /** Look-only edits (a patch, cleared formats): they keep every cell's content, so nothing is renamed or computed. */
  function* lookCells(s: EngineState, edits: CellReplace[]): Steps<[EngineState, Outcome, () => Step | null]> {
    const [out, was] = yield* putCells(s, edits)
    return [s, out, () => (was.length ? cellsStep('patchCells', was) : null)]
  }
  /**
   * Typed cells written as `call`, in one step of the undo with the widths their columns take (`textWidths`), set unless
   * `widen` is false: the state after, what changed and the widths.
   */
  function* writeTexts(s: EngineState, cells: Typed[], call: Step['call'], widen: boolean): Steps<[EngineState, Outcome, ColumnWidth[]]> {
    const [next, out, back] = yield* writeCells(s, cells.map((x) => x.edit), call)
    let o = out
    const steps = [back]
    // a number that does not fit its column widens it, as Excel does, in the same step
    const widths = textWidths(next, cells)
    if (widths.length && widen) {
      const edits = widths.map((w) => ({ sheet: w.sheet, columns: w.column, set: { width: w.width, bestFit: true } }))
      const shown = display ? snapshot(next, layoutCells(next, edits)) : null
      const sheets = applyLayout({ ...next.input, sheets: next.input.sheets.slice() }, edits, (name) => sheetIndex(next)(name)!)
      const [lo, before] = yield* relayout(next, sheets, shown)
      o = joined(o, lo)
      steps.push(() => layoutStep(call, before, sheets))
    }
    if (cells.length) keep(() => manyStep(call, steps.map((f) => f())))
    return [next, o, widths]
  }
  /** What `read` gives with the typed cells written, which are taken back after it. */
  function* previewed<T>(s: EngineState, cells: Typed[], read: () => T): Steps<T> {
    try {
      return yield* withInputs(s, cells.map((x) => x.edit), read)
    } catch (e) {
      // the inputs put in for the preview may still be there: the state is built again on the next call
      stale = true
      throw e
    }
  }
  /** A step taken by undo or redo, answered as its call answered. */
  const taken = (t: Taken | null): UndoResult | null => {
    if (!t) return null
    const out: UndoResult = { call: t.call, ...answer(state!, t.outcome) }
    if (t.cells) out.cells = t.cells
    if (t.mapCell) out.mapCell = t.mapCell
    return out
  }
  const at = (s: EngineState, sheet: string, cell: string) => {
    const i = checkSheet(sheet, sheetIndex(s), 'sheet')
    const { row, col } = parseAddress(checkAddress(cell, 'cell'))
    return [i, row, col] as const
  }
  return {
    get clock() {
      return { ...clock }
    },
    get revision() {
      return revision
    },
    *setCells(edits: CellEdit[], eo?: EditOptions) {
      const s = yield* ready()
      checkBase(eo, revision)
      const [next, out, back] = yield* writeCells(s, resolveEdits(s, editsTaken(checkFieldEdits(edits, sheetIndex(s), styleCount(s)))), 'setCells')
      keep(back)
      return answer(next, out)
    },
    *replaceCells(edits: CellReplace[], eo?: EditOptions) {
      const s = yield* ready()
      checkBase(eo, revision)
      const [next, out, back] = yield* writeCells(s, checkEdits(edits, sheetIndex(s), styleCount(s)) as CellReplace[], 'replaceCells')
      keep(back)
      return answer(next, out)
    },
    *clearCells(ranges: RangeRef[], what: ClearKind, eo?: EditOptions) {
      const s = yield* ready()
      checkBase(eo, revision)
      const plan = clearPlan(s, ranges, what)
      const write = what === 'formats' ? lookCells : (x: EngineState, edits: CellReplace[]) => writeCells(x, edits, 'clearCells')
      const [next, out, steps] = yield* applyPlan(s, plan, 'clearCells', write)
      if (ranges.length) keep(() => manyStep('clearCells', steps.map((f) => f())))
      return answer(next, out)
    },
    *patchCells(patches: CellPatch[], eo?: EditOptions) {
      const s = yield* ready()
      checkBase(eo, revision)
      const plan = patchPlan(s, checkPatches(patches, sheetIndex(s), styleCount(s)))
      const [next, out, steps] = yield* applyPlan(s, plan, 'patchCells', lookCells)
      if (patches.length) keep(() => manyStep('patchCells', steps.map((f) => f())))
      return answer(next, out)
    },
    *setCellText(sheet: string, cell: string, text: string, eo?: CellTextOptions) {
      const s = yield* ready()
      const opts = checkBase<CellTextOptions>(eo, revision, ['widen'])
      const [i, r, c] = at(s, sheet, cell)
      const typed = typedTexts(s, [{ i, r, c, text: checkString(text, 'text') }], clock.now)
      if (!typed.ok) return typed.refusal
      const [next, out, widths] = yield* writeTexts(s, typed.cells, 'setCellText', opts.widen !== false)
      const alert = validationAlert(next, i, r, c)
      return { ok: true, ...answer(next, out), ...(widths.length ? { width: widths[0].width } : {}), ...(alert ? { validation: alert } : {}) }
    },
    *previewCellText(sheet: string, cell: string, text: string, eo?: CellTextOptions): Steps<CellTextPreview> {
      const s = yield* ready()
      checkBase<CellTextOptions>(eo, revision, ['widen'])
      const [i, r, c] = at(s, sheet, cell)
      const typed = typedTexts(s, [{ i, r, c, text: checkString(text, 'text') }], clock.now)
      if (!typed.ok) return typed.refusal
      return yield* previewed(s, typed.cells, () => {
        const width = columnWidening(s, i, r, c)
        const alert = validationAlert(s, i, r, c)
        return { ok: true as const, input: typed.cells[0].edit.input, value: cellValue(s, i, r, c), ...(width !== undefined ? { width } : {}), ...(alert ? { validation: alert } : {}) }
      })
    },
    *setCellTexts(entries: CellTextEntry[], eo?: CellTextOptions) {
      const s = yield* ready()
      const opts = checkBase<CellTextOptions>(eo, revision, ['widen'])
      const typed = typedTexts(s, checkTextEntries(entries, sheetIndex(s)), clock.now)
      if (!typed.ok) return { ...typed.refusal, ...typed.ref }
      const [next, out, widths] = yield* writeTexts(s, typed.cells, 'setCellTexts', opts.widen !== false)
      return { ok: true, ...answer(next, out), ...textsFound(widths, textAlerts(next, typed.cells)) }
    },
    *previewCellTexts(entries: CellTextEntry[], eo?: CellTextOptions): Steps<CellTextsPreview> {
      const s = yield* ready()
      checkBase<CellTextOptions>(eo, revision, ['widen'])
      const typed = typedTexts(s, checkTextEntries(entries, sheetIndex(s)), clock.now)
      if (!typed.ok) return { ...typed.refusal, ...typed.ref }
      return yield* previewed(s, typed.cells, () => {
        const cells = typed.cells.map(({ i, r, c, edit }) => ({ sheet: edit.sheet, cell: edit.cell, input: edit.input, value: cellValue(s, i, r, c) }))
        return { ok: true as const, cells, ...textsFound(textWidths(s, typed.cells), textAlerts(s, typed.cells)) }
      })
    },
    *setNames(names: NameInput[], eo?: EditOptions) {
      const old = yield* ready()
      checkBase(eo, revision)
      names = checkNames(names, 'names')
      // tables and names share one namespace
      checkTables({ ...old.input, names }, 'book')
      // names change the whole graph: the book is rebuilt, changes are collected over all cells (none moved)
      const book = { ...old.input, names }
      const [next, out, diff] = yield* rebuild(old, book, [], false, history.on)
      keep(() => diff && bookStep('setNames', diff, book))
      return answer(next, out)
    },
    *setClock(next: Clock, eo?: EditOptions) {
      if (live) throw new Error('a live session takes its clock from the system')
      const s = yield* ready()
      checkBase(eo, revision)
      const was = clock
      clock = { ...checkClock(next, 'clock') }
      try {
        const keys = yield* applyClock(s, clockSerial(clock), clock.seed)
        return answer(s, { ...nothing(), keys, valued: keys.length > 0 })
      } catch (e) {
        clock = was
        stale = true
        throw e
      }
    },
    *editStructure(op: StructureOp, eo?: EditOptions) {
      const old = yield* ready()
      checkBase(eo, revision)
      op = opTaken(op)
      const r = applyOp(old.input, op, stateValues(old))
      if (!r.ok) return r
      // a sheet hidden or shown changes no value: the book takes its sheets, nothing is built again
      if (op.op === 'hideSheet' || op.op === 'unhideSheet') {
        const next = new Map(r.book.sheets.flatMap((x, i): [number, SheetInput][] => (x === old.input.sheets[i] ? [] : [[i, x]])))
        const [out, before] = yield* relayout(old, next, null)
        keep(() => layoutStep('editStructure', before, next))
        return { ok: true, mapCell: checkedMapCell({ geometry: [], before: old.sheetNames, after: old.sheetNames }), ...answer(old, out) }
      }
      // the simple way: the state is built again from the new book
      const before = old.sheetNames
      const [next, out, diff] = yield* rebuild(old, r.book, r.geometry, true, history.on)
      keep(() => diff && bookStep('editStructure', diff, r.book))
      const moved = movesOf(r.geometry, before, next.sheetNames)
      return { ok: true, mapCell: checkedMapCell({ geometry: r.geometry, before, after: next.sheetNames }), ...answer(next, { ...out, moved }) }
    },
    setStyles(styles: CellStyle[], eo?: EditOptions) {
      const s = st()
      checkBase(eo, revision)
      checkStylesUsed(s.input, checkStyles(styles, 'styles').length)
      const was = s.input.styles
      s.input.styles = styles
      keep(() => stylesStep(was))
      return answer(s, { ...nothing(), styles: true })
    },
    *setLayout(edits: LayoutEdit[], eo?: EditOptions) {
      const s = yield* ready()
      checkBase(eo, revision)
      // AutoFit becomes the widths it sets
      const checked = fitEdits(s, checkLayoutEdits(edits, sheetIndex(s), styleCount(s)))
      const shown = display ? snapshot(s, layoutCells(s, checked)) : null
      // the new sheets, put into the book by `relayout`
      const next = applyLayout({ ...s.input, sheets: s.input.sheets.slice() }, checked, (name) => sheetIndex(s)(name)!)
      const [out, before] = yield* relayout(s, next, shown)
      if (edits.length) keep(() => layoutStep('setLayout', before, next))
      return answer(s, out)
    },
    getBook() {
      return st().input
    },
    getLayout() {
      return bookLayout(st().input)
    },
    *getUsedRange(sheet: string) {
      const s = yield* ready()
      return sheetExtent(s, checkSheet(sheet, sheetIndex(s), 'sheet'))
    },
    *explain(sheet: string, cell: string, eo?: ExplainOptions) {
      const s = yield* ready()
      const [i, row, col] = at(s, sheet, cell)
      const o = checkExplainOptions(eo)
      const e = explainCell(s, i, row, col, { depth: o.depth ?? 3, maxNodes: o.maxNodes ?? 200, expand: o.expand ?? false, errorDepth: o.errorDepth ?? 64 })
      return { ...e, text: explainText(e.root, s.input.locale) }
    },
    *getCells(sheet: string, range: string | string[]) {
      const s = yield* ready()
      const i = checkSheet(sheet, sheetIndex(s), 'sheet')
      const many = Array.isArray(range)
      return yield* cellsIn(s, i, (many ? range : [range]).map((text, k) => checkRange(text, many ? `range[${k}]` : 'range')), display)
    },
    *getCell(sheet: string, cell: string) {
      const s = yield* ready()
      const [i, r, c] = at(s, sheet, cell)
      return (yield* cellsIn(s, i, [{ r1: r, c1: c, r2: r, c2: c }], display))[cell]
    },
    *getViewport(sheet: string, range: string, o?: ViewportOptions) {
      const s = yield* ready()
      const i = checkSheet(sheet, sheetIndex(s), 'sheet')
      // the window of the revision it was read at, for a host to take later edits' changes over
      return { ...(yield* viewport(s, i, checkRange(range, 'range'), checkViewportOptions(o), display, clock.now)), revision }
    },
    *getCellIssue(sheet: string, cell: string) {
      const s = yield* ready()
      return issueDetail(s, ...at(s, sheet, cell))
    },
    *getValidation(sheet: string, cell: string) {
      const s = yield* ready()
      return cellValidation(s, ...at(s, sheet, cell))
    },
    *validateEntry(sheet: string, cell: string, input: CellInput | null) {
      const s = yield* ready()
      const [i, row, col] = at(s, sheet, cell)
      if (input !== null) input = cellTaken(checkCell(input, 'input', styleCount(s)))
      checkArrayEdits(s, [{ sheet, cell, input }])
      if (!hasRule(s, i, row, col)) return { ok: true }
      try {
        return yield* withInput(s, i, row, col, input, () => entryCheck(s, i, row, col))
      } catch (e) {
        // the input put in for the check may still be there: the state is built again on the next call
        stale = true
        throw e
      }
    },
    *getHash() {
      return yield* stateHash(yield* ready())
    },
    *getResult() {
      const t0 = nowMs()
      return yield* buildResult(yield* ready(), clock, t0, display)
    },
    *getSnapshot() {
      const t0 = nowMs()
      const s = yield* ready()
      return { book: s.input, result: yield* buildResult(s, clock, t0, display) }
    },
    *undo(eo?: EditOptions) {
      checkBase(eo, revision)
      return taken(yield* applyStep(core, history.undo.at(-1), (forth) => history.undone(forth)))
    },
    *redo(eo?: EditOptions) {
      checkBase(eo, revision)
      return taken(yield* applyStep(core, history.redo.at(-1), (back) => history.redone(back)))
    },
    get history() {
      return history.state
    },
    historyData: () => history.data(),
    close() {
      state = null
    },
  }
}

/**
 * Opens a session. Without options it keeps the clock of the moment it was opened (with a random seed)
 * until `setClock`; `{ live: true }` takes a fresh clock on every recomputation.
 */
export function createSession(book: BookInput, opts?: SessionOptions): Session {
  const budget = limitsBudget(checkSessionOptions(opts).limits)
  const now = <T>(steps: () => Steps<T>): T => withBudget(budget, () => run(steps()))
  const core = now(() => openSteps(book, opts))
  return strictMethods<Session>({
    get clock() {
      return core.clock
    },
    get revision() {
      return core.revision
    },
    setCells: (edits, o) => now(() => core.setCells(edits, o)),
    replaceCells: (edits, o) => now(() => core.replaceCells(edits, o)),
    clearCells: (ranges, what, o) => now(() => core.clearCells(ranges, what, o)),
    patchCells: (patches, o) => now(() => core.patchCells(patches, o)),
    setCellText: (sheet, cell, text, o) => now(() => core.setCellText(sheet, cell, text, o)),
    previewCellText: (sheet, cell, text, o) => now(() => core.previewCellText(sheet, cell, text, o)),
    setCellTexts: (entries, o) => now(() => core.setCellTexts(entries, o)),
    previewCellTexts: (entries, o) => now(() => core.previewCellTexts(entries, o)),
    setNames: (names, o) => now(() => core.setNames(names, o)),
    setClock: (clock, o) => now(() => core.setClock(clock, o)),
    editStructure: (op, o) => now(() => core.editStructure(op, o)),
    setStyles: (styles, o) => core.setStyles(styles, o),
    setLayout: (edits, o) => now(() => core.setLayout(edits, o)),
    getBook: () => core.getBook(),
    getLayout: () => core.getLayout(),
    getUsedRange: (sheet) => now(() => core.getUsedRange(sheet)),
    explain: (sheet, cell, o) => now(() => core.explain(sheet, cell, o)),
    getCells: (sheet, range) => now(() => core.getCells(sheet, range)),
    getCell: (sheet, cell) => now(() => core.getCell(sheet, cell)),
    getViewport: (sheet, range, o) => now(() => core.getViewport(sheet, range, o)),
    getCellIssue: (sheet, cell) => now(() => core.getCellIssue(sheet, cell)),
    getValidation: (sheet, cell) => now(() => core.getValidation(sheet, cell)),
    validateEntry: (sheet, cell, input) => now(() => core.validateEntry(sheet, cell, input)),
    getResult: () => now(() => core.getResult()),
    getSnapshot: () => now(() => core.getSnapshot()),
    getHash: () => now(() => core.getHash()),
    undo: (o) => now(() => core.undo(o)),
    redo: (o) => now(() => core.redo(o)),
    get history() {
      return core.history
    },
    close: () => core.close(),
  })
}
