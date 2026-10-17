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
import type { ChangesOptions, ChangesPreview, ChangesResult, PreviewChangesOptions, SessionOperation, ReadCellsOptions, ReadCellsResult, FindCellsOptions, DependentsOptions, DependentsResult, RangeSummary, SummaryOptions, ReplaceOptions, ReplaceResult } from './contract.js'
import { operationsPlan } from './session/operations.js'
import { checkCopyOptions, copyText } from './session/clipboard.js'
import type { CopyTextOptions } from './contract.js'
import { readQuery, findQuery, dependentsQuery, summaryQuery } from './session/queries.js'
import { tracker } from './evaluate.js'
import { toCellValue } from './engine.js'
import { outValue } from './result.js'
import type { BookInput, BookLayout, BookSnapshot, CellEdit, CellInput, CellIssueDetail, CellPatch, CellReplace, CellStyle, CellTextEntry, CellTextOptions, CellTextPreview, CellTextResult, CellTextsPreview, CellTextsResult, CellValidation, CellValidationRefusal, CellView, ClearKind, Clock, ColumnWidth, ComputeResult, EditOptions, EndDirection, EditResult, EntryCheck, Explanation, ExplainOptions, FillTextOptions, FillTextPreview, FillTextResult, LayoutEdit, NameInput, NumberFormat, PasteTextPreview, PasteTextResult, RangeRef, Scalar, Session, SessionHistory, SessionOptions, SessionStructureResult, SheetExtent, SheetInput, StructureOp, UndoResult, Viewport, ViewportOptions } from './contract.js'
import { applyClock, applyEdits, applyHidden, applyMerges, buildState, keyParts, setNow, sheetIndexOf, type EngineOptions, type EngineState } from './engine.js'
import { journaled, priorOf } from './held.js'
import { allotWrites, limitsBudget, run, untimed, withBudget, type Steps } from './exec/budget.js'
import { checkExplainOptions, explainCell, explainText } from './explain.js'
import { currentClock, nowMs } from './env.js'
import { checkAddress, checkBook, checkBoolean, checkCell, checkClock, checkEdits, checkLayoutEdits, checkNames, checkNumber, checkSessionOptions, checkSheet, checkString, checkStyles } from './input.js'
import { checkBase, editFields, inOrder, joined, layoutChanged, nothing, repaintOf, ruleRanges, type FieldChange, type Outcome } from './session/answer.js'
import { applyLayout, bookLayout, checkStylesUsed, contentEdits, inputsOf, type SetLayoutEdit } from './session/book.js'
import { typedHeights, wrapTyped } from './session/wrap.js'
import { checkFieldEdits, resolveEdits } from './session/cells.js'
import { clearPlan } from './session/clear.js'
import { subscribers } from './session/subscribers.js'
import { checkArrayEdits } from './spill/fixed.js'
import { sheetExtent } from './session/extent.js'
import { checkDirection, endOf } from './session/end.js'
import { textsChanged } from './session/search.js'
import { replacedTexts } from './session/replace.js'
import { snapshotBook } from './session/snapshot.js'
import { checkRange, checkViewportOptions, viewport } from './session/viewport.js'
import { cellValue, editedCells, layoutCells, shownChanges, snapshot } from './session/looks.js'
import { movedChanges } from './session/moved.js'
import { checkPatches } from './session/patch.js'
import { patchPlan, type Plan } from './session/styles.js'
import { bookDiff, DiffPass, type BookDiff } from './session/rebook.js'
import { columnWidening, entryCheck, validationAlert } from './session/text.js'
import { alertRefusal, checkActive, checkTextEntries, filledText, pastedText, previewCells, strictestRefusal, textAlerts, textsFound, textWidths, typedTexts, type Typed } from './session/texts.js'
import { applyStep, bookStep, cellsStep, History, layoutStep, manyStep, stylesStep, type HistoryData, type SessionCore, type Step, type Taken } from './session/undo.js'
import { headerEdits } from './tables/headers.js'
import { checkTables } from './tables/check.js'
import { issueDetail } from './issues.js'
import { dropStats } from './cond/stats.js'
import { fitEdits } from './layout/fit.js'
import { hiddenRows } from './layout/hidden.js'
import { parseMerge } from './layout/merges.js'
import { formatAddress, parseAddress, readAddress } from './model.js'
import { buildResult, cellAt, cellsIn, clockSerial, spillFields, stateHash } from './result.js'
import { formatCodeIn } from './locale/rules.js'
import { cellLooks } from './shown.js'
import { applyOp, checkedMapCell } from './structure/index.js'
import { mergedPlan } from './session/merged.js'
import { OUTLINE_FIELDS } from './structure/outline.js'
import { inverse, movesOf, type Geometry } from './structure/geometry.js'
import { stateValues } from './structure/values.js'
import { sheetKey } from './syntax.js'
import { cellValidation, hasRule, withInput, withInputs } from './validation/session.js'
import { cellTaken, editsTaken, opTaken } from './taken.js'
import { putInputs, withLooksOf } from './looks/sheet.js'
import { FILL_OPTIONS, REPLACE_OPTIONS, strictMethods, TEXT_OPTIONS } from './strict-args.js'

/** How a session is run by its host. */
export interface SessionHooks {
  /** The clock a live session takes at each recomputation (a thread replaying a journal gives the recorded ones). */
  clock?: () => Clock
  /** The history a session restored from a snapshot starts with. */
  history?: HistoryData
  /** The revision a session restored from a snapshot starts at. */
  revision?: number
  /** False only on a worker: its outgoing snapshot is copied by the transport before another call runs. */
  copySnapshot?: false
}

/** A session's operations as steps; `getBook`, `getLayout`, `setStyles` and `close` compute nothing. */
export interface SessionSteps {
  applyChanges(operations: SessionOperation[], opts?: ChangesOptions): Steps<ChangesResult>
  previewChanges(operations: SessionOperation[], opts?: PreviewChangesOptions): Steps<ChangesPreview>
  readCells(ranges: RangeRef[], opts?: ReadCellsOptions): Steps<ReadCellsResult>
  findCells(ranges: RangeRef[], query: string, opts?: FindCellsOptions): Steps<ReadCellsResult>
  replaceInCells(ranges: RangeRef[], query: string, replacement: string, opts?: ReplaceOptions): Steps<ReplaceResult>
  getDependents(ranges: RangeRef[], opts?: DependentsOptions): Steps<DependentsResult>
  summarizeRanges(ranges: RangeRef[], opts?: SummaryOptions): Steps<RangeSummary>
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
  pasteText(sheet: string, cell: string, text: string, opts?: EditOptions): Steps<PasteTextResult>
  previewPasteText(sheet: string, cell: string, text: string, opts?: EditOptions): Steps<PasteTextPreview>
  fillCellText(sheet: string, range: string, text: string, opts?: FillTextOptions): Steps<FillTextResult>
  previewFillCellText(sheet: string, range: string, text: string, opts?: FillTextOptions): Steps<FillTextPreview>
  setNames(names: NameInput[], opts?: EditOptions): Steps<EditResult>
  setClock(clock: Clock, opts?: EditOptions): Steps<EditResult>
  editStructure(op: StructureOp, opts?: EditOptions): Steps<SessionStructureResult>
  setStyles(styles: CellStyle[], opts?: EditOptions): EditResult
  setLayout(edits: LayoutEdit[], opts?: EditOptions): Steps<EditResult>
  getBook(): BookInput
  getLayout(): BookLayout
  getUsedRange(sheet: string): Steps<SheetExtent>
  getEnd(sheet: string, cell: string, direction: EndDirection): Steps<string>
  undo(opts?: EditOptions): Steps<UndoResult | null>
  redo(opts?: EditOptions): Steps<UndoResult | null>
  readonly history: SessionHistory
  /** What the history holds, for a snapshot. */
  historyData(): HistoryData
  explain(sheet: string, cell: string, opts?: ExplainOptions): Steps<Explanation>
  getCells(sheet: string, range: string | string[]): Steps<Record<string, CellView>>
  copyText(sheet: string, range: string, opts?: CopyTextOptions): Steps<string>
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

/** Edits of cells checked, one cell each, of a book in `locale`: over `limits.writes` they stop the call before anything is written. */
function written<T>(edits: T[], locale: string | undefined): T[] {
  allotWrites(edits.length, locale)
  return edits
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
  // the cells a host was given (`getBook`; `getSnapshot` gives a copy), which it may keep: a state built later copies
  // them, so they change no more, and takes the others as they are (an old state's, an operation's new ones, a deleted sheet's put back)
  const lent = new WeakSet<object>()
  const lend = (s: EngineState) => {
    for (const x of s.input.sheets) lent.add(x.cells)
    return s.input
  }
  const unlent = (cells: object) => !lent.has(cells)
  /** A state of book `b`: its cells copied (the host's book), or those `own` gives taken as they are; `more`, the options of its held values. */
  const build = (b: BookInput, own?: (cells: object) => boolean, more?: Pick<EngineOptions, 'cached' | 'prior' | 'rewritten'>) => buildState(b, { now: clockSerial(clock), seed: clock.seed, externals, execution, own, ...more })
  let state: EngineState | null = yield* build(book, undefined, { cached: o.cached })
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
      state = yield* untimed(build(s.input, unlent))
      stale = false
    }
    return state!
  }
  const sheetIndex = (s: EngineState) => (name: string) => sheetIndexOf(s, name)
  const styleCount = (s: EngineState) => s.input.styles?.length ?? 0
  /** The changes of cells `keys` as a host gets them, with the fields that changed of some and a new format's code. */
  const changes = (s: EngineState, keys: string[], fields: Map<string, FieldChange>) => {
    // the statistics of conditional formats over changed cells are taken again
    dropStats(s, keys)
    const look = display ? cellLooks(s) : null
    const code = formatCodeIn(s.input.locale)
    return keys.map((k) => {
      const [i, r, c] = keyParts(k)
      const value = cellValue(s, i, r, c)
      const f = fields.get(k)
      const format = f?.numberFormat
      return { sheet: s.sheetNames[i], cell: formatAddress(r, c), value, ...look?.(i, r, c, value), ...spillFields(s, i, r, c), ...f, ...(format !== undefined && format !== null ? { numberFormatCode: code(format) } : {}) }
    })
  }
  /** A call that changed the book answered: the revision one more, and what changed. */
  const answer = (s: EngineState, x: Outcome): EditResult => {
    textsChanged(s, x)
    const out: EditResult = { revision: ++revision, changes: changes(s, x.keys, x.fields) }
    if (x.layout.size) out.layout = [...x.layout]
    if (x.styles) out.styles = s.input.styles ?? []
    if (x.moved.length) out.moved = x.moved
    const repaint = repaintOf(s.input, x)
    if (repaint) out.repaint = repaint
    return out
  }
  let controlledClock: Clock | undefined
  const tick = (s: EngineState) => {
    clock = controlledClock ? { ...controlledClock } : fresh()
    setNow(s, clockSerial(clock), clock.seed)
  }
  // a call that builds a new state aside: when it stops, the old state gets its clock back
  const untick = (s: EngineState, at: Clock) => {
    if (clock === at) return
    clock = at
    setNow(s, clockSerial(at), at.seed)
  }
  const history = new History(o.undo, hooks.history)
  // the inputs the last `putCells` changed besides its edits (held values), which a preview puts back too
  let besides: CellReplace[] = []
  // a step is made only when the history keeps steps (a book's diff costs a pass over its cells); a call that changed nothing makes none
  const keep = (step: () => Step | null) => {
    const x = history.on ? step() : null
    if (x) history.push(x)
  }
  /**
   * Edits of cells (checked) applied, `restore` when they put inputs back (`applyEdits`); the inputs they replaced, and
   * those of the cells whose held value the engine took or dropped (`held.ts`), which are put back if the call stops.
   */
  function* putCells(s: EngineState, edits: CellReplace[], restore = false): Steps<[Outcome, CellReplace[]]> {
    const before = display ? snapshot(s, editedCells(s, edits)) : null
    const was = inputsOf(s, edits)
    const at = clock
    s.journal = new Map()
    try {
      // edits that keep a cell's value and formula change its input alone
      const loud = contentEdits(s, edits)
      if (loud.length && (live || controlledClock)) tick(s)
      let keys: string[]
      if (loud.length) keys = yield* applyEdits(s, loud, live || controlledClock !== undefined, restore)
      else if (controlledClock) { tick(s); keys = yield* applyClock(s, clockSerial(clock), clock.seed) }
      else keys = []
      const fields = editFields(s, was)
      const back = journaled(was, s.journal)
      besides = back.slice(was.length)
      return [{ ...nothing(), keys: inOrder(keys, [...(before ? shownChanges(s, before) : []), ...fields.keys()]), fields, valued: keys.length > 0 }, back]
    } catch (e) {
      putInputs(s.input.sheets, (name) => sheetIndex(s)(name)!, journaled(was, s.journal))
      clock = at
      stale = true
      throw e
    } finally {
      s.journal = undefined
    }
  }
  /**
   * The state built again from `book`, the session's book changed, with its changes against the old one by the geometry
   * `g` (where the old cells went; the looks of both with `looks`) and, with `back`, the way back to the old book, taken
   * in the same pass over the cells; nothing changes if it stops. `restore`: the book is one put back (an undo), whose
   * formulas a cycle stops hold what their inputs give (`held.ts`).
   */
  function* rebuild(old: EngineState, book: BookInput, g: Geometry, looks: boolean, back: boolean, restore = false): Steps<[EngineState, Outcome, BookDiff | null]> {
    const at = clock
    if (live || controlledClock) tick(old)
    try {
      // a formula a cycle stops now holds what it showed (`held.ts`): the sheets so changed are the new book's in the diff
      const rewritten = new Set<number>()
      const next = yield* build(book, unlent, { ...(restore ? {} : { prior: priorOf(old, g) }), rewritten })
      const fields = new Map<string, FieldChange>()
      const current = rewritten.size ? { ...book, sheets: book.sheets.map((x, i) => (rewritten.has(i) ? next.input.sheets[i] : x)) } : book
      const pass = back ? new DiffPass(current, old.input) : undefined
      const keys = movedChanges(old, next, g, fields, display && looks ? [cellLooks(old), cellLooks(next)] : undefined, pass)
      const diff = pass ? yield* bookDiff(pass, inverse(g)) : null
      state = next
      // a style a structure operation spelled out, or its undo took away
      const styles = JSON.stringify(old.input.styles ?? []) !== JSON.stringify(next.input.styles ?? [])
      return [next, { ...nothing(), keys, fields, valued: true, layout: new Set(layoutChanged(old.input, next.input)), styles }, diff]
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
    // a sheet made before cells were written keeps the cells' looks as they are now: a layout never changes them
    for (const [i, x] of next) next.set(i, (s.input.sheets[i] = withLooksOf(x, sheets[i])))
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
   * `widen` is false, the styles that wrap a text with a line break and the heights of the rows that wrap (`wrap.ts`);
   * `pasted`, a paste, which measures no column and fits no row: the state after, what changed and the widths.
   */
  function* writeTexts(s: EngineState, cells: Typed[], call: Step['call'], widen: boolean, pasted = false): Steps<[EngineState, Outcome, ColumnWidth[]]> {
    // a text with a line break wraps its cell, by a style the table may gain
    const wrapped = wrapTyped(s, cells)
    const [next, out, back] = yield* writeCells(s, cells.map((x) => x.edit), call)
    let o = out
    const steps = [back]
    if (wrapped) {
      const was = next.input.styles
      next.input.styles = wrapped
      o = { ...o, styles: true }
      steps.push(() => stylesStep(was, call))
    }
    // a number that does not fit its column widens it, and a row whose cell wraps fits it, as Excel does, in the same step
    const widths = pasted ? [] : textWidths(next, cells)
    const edits: SetLayoutEdit[] = [...(widen ? widths.map((w) => ({ sheet: w.sheet, columns: w.column, set: { width: w.width, bestFit: true } })) : []), ...(pasted ? [] : typedHeights(next, cells))]
    if (edits.length) {
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
    const p = { r: 0, c: 0 }
    readAddress(checkAddress(cell, 'cell'), p)
    return [i, p.r, p.c] as const
  }
  /** A text call's options checked: `baseRevision`, `widen`, `refuseInvalid`, `inPlace` and, entering a selection, `active`. */
  const textOptions = <T extends CellTextOptions>(eo: T | undefined, allowed: readonly string[] = TEXT_OPTIONS): T => {
    const o = checkBase<T>(eo, revision, allowed)
    if (o.refuseInvalid !== undefined) checkBoolean(o.refuseInvalid, 'opts.refuseInvalid')
    if (o.inPlace !== undefined) checkBoolean(o.inPlace, 'opts.inPlace')
    return o
  }
  const wayOf = (o: CellTextOptions) => (o.inPlace ? 'inPlace' : 'typed')
  /**
   * With `refuseInvalid`, the cell of `checked` whose value breaks the strictest rule of data validation once the typed
   * cells are written (`strictestRefusal`), refused before they are, as Excel's alert holds the entry until it is answered.
   */
  function* invalid(s: EngineState, o: CellTextOptions, cells: Typed[], checked = cells): Steps<CellValidationRefusal | null> {
    // a cell without a rule takes any value: nothing is put in to check it
    if (!o.refuseInvalid || !checked.some(({ i, r, c }) => hasRule(s, i, r, c))) return null
    return strictestRefusal(yield* previewed(s, cells, () => textAlerts(s, checked)))
  }
  /** A text typed into a cell (`setCellText` and its preview), its arguments checked: the state, the options, the cell and the cell typed. */
  function* typedCell(sheet: string, cell: string, text: string, eo: CellTextOptions | undefined) {
    const s = yield* ready()
    const opts = textOptions(eo)
    const [i, r, c] = at(s, sheet, cell)
    return { s, opts, i, r, c, typed: typedTexts(s, [{ i, r, c, text: checkString(text, 'text') }], clock.now, wayOf(opts)) }
  }
  /** Texts typed into cells (`setCellTexts` and its preview), checked: the state, the options and the cells typed. */
  function* typedCells(entries: CellTextEntry[], eo: CellTextOptions | undefined) {
    const s = yield* ready()
    const opts = textOptions(eo)
    return { s, opts, typed: typedTexts(s, written(checkTextEntries(entries, sheetIndex(s)), s.input.locale), clock.now, wayOf(opts)) }
  }
  /** A text pasted at a cell (`pasteText` and its preview), checked: the state and the cells pasted. */
  function* pastedAt(sheet: string, cell: string, text: string, eo: EditOptions | undefined) {
    const s = yield* ready()
    checkBase(eo, revision)
    return { s, pasted: pastedText(s, ...at(s, sheet, cell), checkString(text, 'text'), clock.now) }
  }
  /** A text entered into a range from its active cell (`fillCellText` and its preview), checked: the state, the options, the sheet, the active cell and the cells. */
  function* filledRange(sheet: string, range: string, text: string, eo: FillTextOptions | undefined) {
    const s = yield* ready()
    const opts = textOptions(eo, FILL_OPTIONS)
    const i = checkSheet(sheet, sheetIndex(s), 'sheet')
    const x = checkRange(range, 'range')
    const a = checkActive(opts.active, x)
    return { s, opts, i, a, filled: filledText(s, i, x, a, checkString(text, 'text'), clock.now, wayOf(opts)) }
  }
  /** Executes a staged batch; previews undo their sparse inputs or drop their aside-built state. */
  function* operations(given: SessionOperation[], eo: PreviewChangesOptions | undefined, preview: boolean): Steps<ChangesResult | ChangesPreview> {
    const old = yield* ready(),
      beforeRevision = revision,
      beforeClock = clock
    const opts = checkBase<PreviewChangesOptions>(eo, revision, preview ? ['clock', 'maxChanges'] : ['clock'])
    if (opts.clock) checkClock(opts.clock, 'opts.clock')
    const max = opts.maxChanges ?? 1000
    if (preview) checkNumber(max, 'opts.maxChanges', 1, 10000, true)
    const plan = operationsPlan(old, given, (v, format) => headerText(old, v, format))
    const beforeBook = { ...old.input, sheets: old.input.sheets.slice() }
    besides = []
    const captured = tracker()
    if (preview) old.capture = captured
    controlledClock = opts.clock ?? (preview && live ? fresh() : undefined)
    let completed = false,
      next = old,
      outcome: Outcome,
      back: (() => Step | null) | undefined
    try {
      if (plan.rebuild) {
        const [built, out, diff] = yield* rebuild(old, plan.book, [], true, !preview && history.on)
        next = built
        outcome = { ...out, styles: plan.styles }
        back = () => diff && bookStep('applyChanges', diff, plan.book)
      } else {
        const [built, out, cellsBack] = yield* writeCells(old, plan.edits, 'applyChanges')
        next = built
        outcome = out
        const steps: (() => Step | null)[] = [cellsBack]
        if (plan.layout.size) {
          const layouts = new Map([...plan.layout].map(([i, sh]): [number, SheetInput] => [i, withLooksOf({ ...sh, cells: next.input.sheets[i].cells }, next.input.sheets[i])]))
          const [lo, before] = yield* relayout(next, layouts, null)
          outcome = joined(outcome, lo)
          steps.push(() => layoutStep('applyChanges', before, layouts))
        }
        if (plan.styles) {
          const was = next.input.styles
          next.input.styles = plan.book.styles
          outcome = { ...outcome, styles: true }
          steps.push(() => stylesStep(was, 'applyChanges'))
        }
        back = () => manyStep('applyChanges', steps.map((f) => f()))
      }
      const used = { ...clock }
      if (preview) {
        // The old state's lazy fills must be read with its original clock.
        if (next !== old) setNow(old, clockSerial(beforeClock), beforeClock.seed)
        const list = changes(next, outcome.keys.slice(0, max), outcome.fields)
        const impacts = list.map((c) => {
          const i = old.sheetIdx.get(sheetKey(c.sheet))!,
            { row, col } = parseAddress(c.cell),
            k = `${i}:${row}:${col}`
          const before = captured.cells.has(k) ? outValue(toCellValue(captured.cells.get(k)!)) : cellValue(old, i, row, col)
          const issue = issueDetail(next, i, row, col)
          return { sheet: c.sheet, cell: c.cell, before, after: c.value, ...(issue ? { issue } : {}) }
        })
        return { revision: beforeRevision, clock: used, changes: list, impacts, totalChanges: outcome.keys.length, truncated: outcome.keys.length > max, ...(outcome.layout.size ? { layout: [...outcome.layout] } : {}), ...(outcome.styles ? { styles: next.input.styles ?? [] } : {}), ...(repaintOf(next.input, outcome) ? { repaint: repaintOf(next.input, outcome) } : {}) }
      }
      const result = { ...answer(next, outcome), clock: used }
      keep(back)
      completed = true
      return result
    } finally {
      old.capture = undefined
      controlledClock = undefined
      if (preview || !completed) {
        revision = beforeRevision
        clock = beforeClock
        state = old
        const back = [...plan.back, ...besides]
        putInputs(beforeBook.sheets, (name) => old.sheetIdx.get(sheetKey(name))!, back)
        old.input = beforeBook
        setNow(old, clockSerial(beforeClock), beforeClock.seed)
        if (next === old) {
          stale = true
          if (preview) {
            yield* applyEdits(old, back, true, true)
            stale = false
          }
        }
      }
      controlledClock = undefined
    }
  }
  return {
    *applyChanges(ops: SessionOperation[], eo?: ChangesOptions) {
      return (yield* operations(ops, eo, false)) as ChangesResult
    },
    *previewChanges(ops: SessionOperation[], eo?: PreviewChangesOptions) {
      return (yield* operations(ops, eo, true)) as ChangesPreview
    },
    *readCells(refs: RangeRef[], eo?: ReadCellsOptions) {
      return yield* readQuery(yield* ready(), revision, clock, refs, eo)
    },
    *findCells(refs: RangeRef[], query: string, eo?: FindCellsOptions) {
      return yield* findQuery(yield* ready(), revision, clock, refs, query, eo)
    },
    *replaceInCells(refs: RangeRef[], query: string, replacement: string, eo?: ReplaceOptions): Steps<ReplaceResult> {
      const s = yield* ready()
      const { cells, refused } = yield* replacedTexts(s, refs, query, replacement, checkBase<ReplaceOptions>(eo, revision, REPLACE_OPTIONS), clock.now)
      const not = refused.length ? { refused } : {}
      // nothing found: the book, its revision and its history as they were
      if (!cells.length) return { revision, changes: [], replaced: 0, ...not }
      const [next, out] = yield* writeTexts(s, cells, 'replaceInCells', false)
      return { ...answer(next, out), replaced: cells.length, ...not }
    },
    *getDependents(refs: RangeRef[], eo?: DependentsOptions) {
      return yield* dependentsQuery(yield* ready(), revision, clock, refs, eo)
    },
    *summarizeRanges(refs: RangeRef[], eo?: SummaryOptions) {
      return yield* summaryQuery(yield* ready(), revision, clock, refs, eo)
    },
    get clock() {
      return { ...clock }
    },
    get revision() {
      return revision
    },
    *setCells(edits: CellEdit[], eo?: EditOptions) {
      const s = yield* ready()
      checkBase(eo, revision)
      const [next, out, back] = yield* writeCells(s, resolveEdits(s, editsTaken(written(checkFieldEdits(edits, sheetIndex(s), styleCount(s)), s.input.locale))), 'setCells')
      keep(back)
      return answer(next, out)
    },
    *replaceCells(edits: CellReplace[], eo?: EditOptions) {
      const s = yield* ready()
      checkBase(eo, revision)
      const [next, out, back] = yield* writeCells(s, written(checkEdits(edits, sheetIndex(s), styleCount(s)) as CellReplace[], s.input.locale), 'replaceCells')
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
      const { s, opts, i, r, c, typed } = yield* typedCell(sheet, cell, text, eo)
      if (!typed.ok) return typed.refusal
      const refused = yield* invalid(s, opts, typed.cells)
      if (refused) return refused
      const [next, out, widths] = yield* writeTexts(s, typed.cells, 'setCellText', opts.widen !== false)
      const alert = validationAlert(next, i, r, c)
      return { ok: true, ...answer(next, out), ...(widths.length ? { width: widths[0].width } : {}), ...(alert ? { validation: alert } : {}) }
    },
    *previewCellText(sheet: string, cell: string, text: string, eo?: CellTextOptions): Steps<CellTextPreview> {
      const { s, opts, i, r, c, typed } = yield* typedCell(sheet, cell, text, eo)
      if (!typed.ok) return typed.refusal
      return yield* previewed(s, typed.cells, () => {
        const alert = validationAlert(s, i, r, c)
        if (alert && opts.refuseInvalid) return alertRefusal({ sheet: s.sheetNames[i], cell: formatAddress(r, c), ...alert })
        const width = columnWidening(s, i, r, c)
        return { ok: true as const, input: typed.cells[0].edit.input, value: cellValue(s, i, r, c), ...(width !== undefined ? { width } : {}), ...(alert ? { validation: alert } : {}) }
      })
    },
    *setCellTexts(entries: CellTextEntry[], eo?: CellTextOptions) {
      const { s, opts, typed } = yield* typedCells(entries, eo)
      if (!typed.ok) return { ...typed.refusal, ...typed.ref }
      const refused = yield* invalid(s, opts, typed.cells)
      if (refused) return refused
      const [next, out, widths] = yield* writeTexts(s, typed.cells, 'setCellTexts', opts.widen !== false)
      return { ok: true, ...answer(next, out), ...textsFound(widths, textAlerts(next, typed.cells)) }
    },
    *previewCellTexts(entries: CellTextEntry[], eo?: CellTextOptions): Steps<CellTextsPreview> {
      const { s, opts, typed } = yield* typedCells(entries, eo)
      if (!typed.ok) return { ...typed.refusal, ...typed.ref }
      return yield* previewed(s, typed.cells, () => {
        const alerts = textAlerts(s, typed.cells)
        const refused = opts.refuseInvalid ? strictestRefusal(alerts) : null
        return refused ?? { ok: true as const, cells: previewCells(s, typed.cells), ...textsFound(textWidths(s, typed.cells), alerts) }
      })
    },
    *pasteText(sheet: string, cell: string, text: string, eo?: EditOptions) {
      const { s, pasted } = yield* pastedAt(sheet, cell, text, eo)
      if (!pasted.ok) return pasted.refusal
      // a paste wraps a text with a line break but fits no row, widens no column and checks no rule of data validation
      const [next, out] = yield* writeTexts(s, pasted.cells, 'pasteText', false, true)
      return { ok: true, ...answer(next, out), range: pasted.range }
    },
    *previewPasteText(sheet: string, cell: string, text: string, eo?: EditOptions): Steps<PasteTextPreview> {
      const { s, pasted } = yield* pastedAt(sheet, cell, text, eo)
      if (!pasted.ok) return pasted.refusal
      return yield* previewed(s, pasted.cells, () => ({ ok: true as const, range: pasted.range, cells: previewCells(s, pasted.cells) }))
    },
    *fillCellText(sheet: string, range: string, text: string, eo?: FillTextOptions) {
      const { s, opts, i, a, filled } = yield* filledRange(sheet, range, text, eo)
      if (!filled.ok) return { ...filled.refusal, ...filled.ref }
      // the active cell's rule alone is checked, as Excel checks it
      const refused = yield* invalid(s, opts, filled.cells, filled.cells.filter((x) => x.r === a.r && x.c === a.c))
      if (refused) return refused
      const [next, out, widths] = yield* writeTexts(s, filled.cells, 'fillCellText', opts.widen !== false)
      const alert = validationAlert(next, i, a.r, a.c)
      return { ok: true, ...answer(next, out), ...textsFound(widths, []), ...(alert ? { validation: alert } : {}) }
    },
    *previewFillCellText(sheet: string, range: string, text: string, eo?: FillTextOptions): Steps<FillTextPreview> {
      const { s, opts, i, a, filled } = yield* filledRange(sheet, range, text, eo)
      if (!filled.ok) return { ...filled.refusal, ...filled.ref }
      return yield* previewed(s, filled.cells, () => {
        const alert = validationAlert(s, i, a.r, a.c)
        if (alert && opts.refuseInvalid) return alertRefusal({ sheet: s.sheetNames[i], cell: formatAddress(a.r, a.c), ...alert })
        return { ok: true as const, cells: previewCells(s, filled.cells), ...textsFound(textWidths(s, filled.cells), []), ...(alert ? { validation: alert } : {}) }
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
      // a merge that moves no cell is written as a patch is (`session/merged.ts`)
      const plan = op.op === 'mergeCells' || op.op === 'unmergeCells' ? mergedPlan(old, op) : null
      if (plan && 'ok' in plan) return plan
      if (plan) {
        const [next, out, steps] = yield* applyPlan(old, plan, 'editStructure', (x, edits) => writeCells(x, edits, 'editStructure'))
        keep(() => manyStep('editStructure', steps.map((f) => f())))
        return { ok: true, mapCell: checkedMapCell({ geometry: [], before: old.sheetNames, after: old.sheetNames }), ...answer(next, out) }
      }
      const r = applyOp(old.input, op, stateValues(old))
      if (!r.ok) return r
      // a sheet hidden or shown, and an outline command, move no cell: the book takes its sheets, nothing is built again
      if (((op.op === 'hideSheet' || op.op === 'unhideSheet') && r.book.activeSheet === old.input.activeSheet) || op.op in OUTLINE_FIELDS) {
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
      return lend(st())
    },
    getLayout() {
      return bookLayout(st().input)
    },
    *getUsedRange(sheet: string) {
      const s = yield* ready()
      return sheetExtent(s, checkSheet(sheet, sheetIndex(s), 'sheet'))
    },
    *getEnd(sheet: string, cell: string, direction: EndDirection) {
      const s = yield* ready()
      const [i, r, c] = at(s, sheet, cell)
      return formatAddress(...endOf(s, i, r, c, checkDirection(direction)))
    },
    *explain(sheet: string, cell: string, eo?: ExplainOptions) {
      const s = yield* ready()
      const [i, row, col] = at(s, sheet, cell)
      const o = checkExplainOptions(eo)
      const e = yield* explainCell(s, i, row, col, { depth: o.depth ?? 3, maxNodes: o.maxNodes ?? 200, expand: o.expand ?? false, errorDepth: o.errorDepth ?? 64 })
      return { ...e, text: yield* explainText(e.root, s.input.locale, s.input.externalBooks) }
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
      return yield* cellAt(s, i, r, c, display)
    },
    *copyText(sheet: string, range: string, o?: CopyTextOptions) {
      const s = yield* ready()
      const i = checkSheet(sheet, sheetIndex(s), 'sheet')
      return yield* copyText(s, i, checkRange(range, 'range'), checkCopyOptions(o))
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
      return cellValidation(s, ...at(s, sheet, cell), clock.now)
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
      const book = hooks.copySnapshot === false ? s.input : yield* snapshotBook(s.input)
      return { book, result: yield* buildResult(s, clock, t0, display, book) }
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
  const subs = subscribers(core.revision)
  return strictMethods<Session>({
    get clock() {
      return core.clock
    },
    get revision() {
      return core.revision
    },
    applyChanges: (ops, o) => subs.after('applyChanges', now(() => core.applyChanges(ops, o))),
    previewChanges: (ops, o) => now(() => core.previewChanges(ops, o)),
    readCells: (refs, o) => now(() => core.readCells(refs, o)),
    findCells: (refs, query, o) => now(() => core.findCells(refs, query, o)),
    replaceInCells: (refs, query, replacement, o) => subs.after('replaceInCells', now(() => core.replaceInCells(refs, query, replacement, o))),
    getDependents: (refs, o) => now(() => core.getDependents(refs, o)),
    summarizeRanges: (refs, o) => now(() => core.summarizeRanges(refs, o)),
    setCells: (edits, o) => subs.after('setCells', now(() => core.setCells(edits, o))),
    replaceCells: (edits, o) => subs.after('replaceCells', now(() => core.replaceCells(edits, o))),
    clearCells: (ranges, what, o) => subs.after('clearCells', now(() => core.clearCells(ranges, what, o))),
    patchCells: (patches, o) => subs.after('patchCells', now(() => core.patchCells(patches, o))),
    setCellText: (sheet, cell, text, o) => subs.after('setCellText', now(() => core.setCellText(sheet, cell, text, o))),
    previewCellText: (sheet, cell, text, o) => now(() => core.previewCellText(sheet, cell, text, o)),
    setCellTexts: (entries, o) => subs.after('setCellTexts', now(() => core.setCellTexts(entries, o))),
    previewCellTexts: (entries, o) => now(() => core.previewCellTexts(entries, o)),
    pasteText: (sheet, cell, text, o) => subs.after('pasteText', now(() => core.pasteText(sheet, cell, text, o))),
    previewPasteText: (sheet, cell, text, o) => now(() => core.previewPasteText(sheet, cell, text, o)),
    fillCellText: (sheet, range, text, o) => subs.after('fillCellText', now(() => core.fillCellText(sheet, range, text, o))),
    previewFillCellText: (sheet, range, text, o) => now(() => core.previewFillCellText(sheet, range, text, o)),
    setNames: (names, o) => subs.after('setNames', now(() => core.setNames(names, o))),
    setClock: (clock, o) => subs.after('setClock', now(() => core.setClock(clock, o))),
    editStructure: (op, o) => subs.after('editStructure', now(() => core.editStructure(op, o))),
    setStyles: (styles, o) => subs.after('setStyles', core.setStyles(styles, o)),
    setLayout: (edits, o) => subs.after('setLayout', now(() => core.setLayout(edits, o))),
    getBook: () => core.getBook(),
    getLayout: () => core.getLayout(),
    getUsedRange: (sheet) => now(() => core.getUsedRange(sheet)),
    getEnd: (sheet, cell, direction) => now(() => core.getEnd(sheet, cell, direction)),
    explain: (sheet, cell, o) => now(() => core.explain(sheet, cell, o)),
    getCells: (sheet, range) => now(() => core.getCells(sheet, range)),
    copyText: (sheet, range, o) => now(() => core.copyText(sheet, range, o)),
    getCell: (sheet, cell) => now(() => core.getCell(sheet, cell)),
    getViewport: (sheet, range, o) => now(() => core.getViewport(sheet, range, o)),
    getCellIssue: (sheet, cell) => now(() => core.getCellIssue(sheet, cell)),
    getValidation: (sheet, cell) => now(() => core.getValidation(sheet, cell)),
    validateEntry: (sheet, cell, input) => now(() => core.validateEntry(sheet, cell, input)),
    getResult: () => now(() => core.getResult()),
    getSnapshot: () => now(() => core.getSnapshot()),
    getHash: () => now(() => core.getHash()),
    undo: (o) => subs.after('undo', now(() => core.undo(o))),
    redo: (o) => subs.after('redo', now(() => core.redo(o))),
    get history() {
      return core.history
    },
    subscribe: (listener) => subs.subscribe(listener),
    close() {
      subs.clear()
      core.close()
    },
  })
}
