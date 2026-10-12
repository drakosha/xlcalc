/**
 * A session's history (`Session.undo`): each step holds the way back of one call, what that call replaced and no more:
 * the inputs of the cells it changed, the fields and the lines of the sheets whose layout it changed, the style table,
 * or the diff of a book built again (`rebook.ts`). Undoing a step makes the way forth from what it replaces, so the two
 * stacks hold one kind of step. Steps are kept by `depth` and by an estimate of their memory.
 */
import type { BookInput, CellReplace, CellRef, CellStyle, LayoutEdit, SessionHistory, SheetInput, UndoCall, UndoOptions } from '../contract.js'
import type { EngineState } from '../engine.js'
import type { Steps } from '../exec/budget.js'
import { movesOf, type Geometry } from '../structure/geometry.js'
import { checkedMapCell, namesOf } from '../structure/index.js'
import { sheetKey } from '../syntax.js'
import { joined, nothing, type Outcome } from './answer.js'
import { layoutCells, snapshot } from './looks.js'
import { applyDiff, type BookDiff } from './rebook.js'

/** The fields of a sheet but its cells, and how its lines (`rows`, `columns`) changed entry by entry. */
export interface SheetBack {
  /** The sheet, by index. */
  i: number
  /** Fields to set to these values, and fields to remove. */
  set: Record<string, unknown>
  unset: string[]
  /** Lines whose entries change: keys to remove, then entries to put. */
  lines: { field: 'rows' | 'columns'; drop: string[]; put: [string, object][] }[]
}

export type Step =
  | { kind: 'cells'; call: UndoCall; edits: CellReplace[]; size: number }
  | { kind: 'layout'; call: UndoCall; sheets: SheetBack[]; size: number }
  | { kind: 'styles'; call: UndoCall; styles: CellStyle[] | undefined; size: number }
  | { kind: 'book'; call: UndoCall; diff: BookDiff; size: number }
  /**
   * Several of one call, taken in their order either way: they change different parts of the book, and the first, the
   * cells' inputs, is the only one that recomputes, so a step stopped there has changed nothing.
   */
  | { kind: 'many'; call: UndoCall; steps: Step[]; size: number }

/** What a session's history holds, to restore it elsewhere (a pool's snapshot). */
export interface HistoryData {
  undo: Step[]
  redo: Step[]
}

const DEPTH = 100
const MAX_MB = 64

/**
 * A rough size in bytes of what a step keeps, as V8 holds it: an object a header and a slot of 24 bytes a field (its key
 * too in a large one, kept as a dictionary), an array 8 bytes an item, a string 24 bytes and a byte a character, a
 * number a heap number of 16. Measured on the cells of a sheet of 506,594 (126 bytes a cell, 118 in V8).
 */
export function weigh(v: unknown): number {
  if (typeof v === 'string') return 24 + v.length
  if (typeof v === 'number') return 16
  if (v === null || typeof v !== 'object') return 8
  let n = 32
  if (Array.isArray(v)) for (const x of v) n += 8 + weigh(x)
  else {
    let keys = 0
    for (const k in v) {
      n += 24 + weigh((v as Record<string, unknown>)[k])
      if (++keys > 64) n += 24 + k.length
    }
  }
  return n
}

export class History {
  undo: Step[] = []
  redo: Step[] = []
  private bytes = 0
  readonly depth: number
  private readonly max: number

  constructor(o: UndoOptions | undefined, data?: HistoryData) {
    this.depth = o?.depth ?? DEPTH
    this.max = (o?.maxMb ?? MAX_MB) * 2 ** 20
    if (data) {
      this.undo = data.undo
      this.redo = data.redo
      this.bytes = [...this.undo, ...this.redo].reduce((n, s) => n + s.size, 0)
    }
  }

  /** Whether steps are kept at all. */
  get on(): boolean {
    return this.depth > 0 && this.max > 0
  }

  get state(): SessionHistory {
    return { undo: this.undo.length, redo: this.redo.length }
  }

  /** A new change: what could be done again is forgotten. */
  push(step: Step | null) {
    for (const s of this.redo) this.bytes -= s.size
    this.redo = []
    this.keep(this.undo, step)
  }

  /** The step undone (the last of `undo`) and its way forth; none when it is too large to keep. */
  undone(forth: Step | null) {
    this.bytes -= this.undo.pop()!.size
    this.keep(this.redo, forth)
  }

  /** The step done again (the last of `redo`) and its way back. */
  redone(back: Step | null) {
    this.bytes -= this.redo.pop()!.size
    this.keep(this.undo, back)
  }

  /**
   * A step on top of `stack`: past `depth` or the budget the oldest steps go, those to undo first; a step larger than the
   * budget alone is not kept, and `stack` is cleared, as the steps under it cannot be reached without it.
   */
  private keep(stack: Step[], step: Step | null) {
    if (!step || step.size > this.max || !this.depth) {
      for (const s of stack) this.bytes -= s.size
      stack.length = 0
      return
    }
    stack.push(step)
    this.bytes += step.size
    while (this.undo.length > this.depth) this.bytes -= this.undo.shift()!.size
    // the oldest to undo first, then the last to redo; never the step just kept
    while (this.bytes > this.max) {
      const from = this.undo.length > (stack === this.undo ? 1 : 0) ? this.undo : this.redo
      this.bytes -= from.shift()!.size
    }
  }

  data(): HistoryData {
    return { undo: this.undo, redo: this.redo }
  }
}

/** The way from sheet `from` to sheet `to` (the same sheet before and after a change of its layout), but its cells and their looks. */
function sheetBack(i: number, from: SheetInput, to: SheetInput): SheetBack | null {
  const set: Record<string, unknown> = {}
  const unset: string[] = []
  const lines: SheetBack['lines'] = []
  for (const k of new Set([...Object.keys(from), ...Object.keys(to)])) {
    const [a, b] = [(from as unknown as Record<string, unknown>)[k], (to as unknown as Record<string, unknown>)[k]]
    if (k === 'cells' || k === 'looks' || a === b) continue
    if ((k === 'rows' || k === 'columns') && a && b) {
      // a line's entries change: the entries that differ, not the whole map
      const [x, y] = [a as Record<string, object>, b as Record<string, object>]
      const drop = Object.keys(x).filter((key) => x[key] !== y[key])
      const put = Object.keys(y).filter((key) => x[key] !== y[key]).map((key): [string, object] => [key, y[key]])
      lines.push({ field: k, drop, put })
    } else if (b === undefined) unset.push(k)
    else set[k] = b
  }
  return Object.keys(set).length || unset.length || lines.length ? { i, set, unset, lines } : null
}

/** Sheet `sheet` with the way applied: a new object with the same cells. */
export function sheetWith(sheet: SheetInput, back: SheetBack): SheetInput {
  const out = { ...sheet } as unknown as Record<string, unknown>
  for (const k of back.unset) delete out[k]
  Object.assign(out, back.set)
  for (const { field, drop, put } of back.lines) {
    const map = { ...(out[field] as Record<string, object>) }
    for (const k of drop) delete map[k]
    for (const [k, v] of put) map[k] = v
    out[field] = map
  }
  return out as unknown as SheetInput
}

/** The fields of `a` that `b` does not share, as `weigh` counts them. */
const own = (a: object, b: object | undefined): number => {
  let n = 32
  for (const [k, v] of Object.entries(a)) if (v !== (b as Record<string, unknown> | undefined)?.[k]) n += 24 + weigh(v)
  return n
}

/** The size of a book's diff from `current`: the fields and cells it keeps that `current` does not share. */
function diffSize(d: BookDiff, current: BookInput): number {
  const sheets = new Map(current.sheets.map((s) => [sheetKey(s.name), s]))
  return own(d.head, current) + weigh(d.geometry) + d.sheets.reduce((n, x) => n + ('rest' in x ? own(x.rest, sheets.get(sheetKey(x.rest.name))) + weigh(x.put) : 32), 0)
}

/** The steps of each kind with their sizes: cells' inputs to put back, as edits. */
export const cellsStep = (call: UndoCall, edits: CellReplace[]): Step => ({ kind: 'cells', call, edits, size: edits.reduce((n, e) => n + 48 + weigh(e.input), 0) })

/** The way back from sheets `after` (by index) to the sheets `before` of the book. */
export function layoutStep(call: UndoCall, before: SheetInput[], after: Map<number, SheetInput>): Step {
  const sheets = [...after].map(([i, x]) => sheetBack(i, x, before[i])).filter((x) => x !== null)
  return { kind: 'layout', call, sheets, size: sheets.reduce((n, x) => n + weigh(x.set) + weigh(x.lines), 0) }
}

export const stylesStep = (styles: CellStyle[] | undefined, call: UndoCall = 'setStyles'): Step => ({ kind: 'styles', call, styles, size: weigh(styles) })

/** The way back from `book`, the book a rebuild was given, by its diff. */
export const bookStep = (call: UndoCall, diff: BookDiff, book: BookInput): Step => ({ kind: 'book', call, diff, size: diffSize(diff, book) })

/** The steps of one call as one (`Step` many), a lone step as it is; none as an empty step, for a call given something to do that changed nothing. */
export function manyStep(call: UndoCall, steps: (Step | null)[]): Step {
  const xs = steps.filter((x) => x !== null)
  return xs.length === 1 ? xs[0] : { kind: 'many', call, steps: xs, size: xs.reduce((n, x) => n + x.size, 0) }
}

/** What a step needs of its session (`session.ts`): the ways it changes the book, which forward calls take too. */
export interface SessionCore {
  ready(): Steps<EngineState>
  putCells(s: EngineState, edits: CellReplace[], restore?: boolean): Steps<[Outcome, CellReplace[]]>
  rebuild(old: EngineState, book: BookInput, g: Geometry, looks: boolean, back: boolean, restore?: boolean): Steps<[EngineState, Outcome, BookDiff | null]>
  relayout(s: EngineState, next: Map<number, SheetInput>, shown: Map<string, string> | null): Steps<[Outcome, SheetInput[]]>
  display: boolean
  sheetIndex(s: EngineState): (name: string) => number | undefined
}

/** A step taken: the call it was of, what it changed, and the cells or the mapCell a host reads by. */
export interface Taken {
  call: UndoCall
  outcome: Outcome
  cells?: CellRef[]
  mapCell?(ref: CellRef): CellRef | null
}

/** Layout edits that stand for a way of a sheet's layout where the look of cells can change: its lines and merged areas. */
function touched(s: EngineState, b: SheetBack): LayoutEdit[] {
  const sheet = s.sheetNames[b.i]
  const keys = { rows: [] as string[], columns: [] as string[] }
  for (const { field, drop, put } of b.lines) keys[field].push(...drop, ...put.map(([k]) => k))
  for (const field of ['rows', 'columns'] as const)
    if (field in b.set || b.unset.includes(field)) keys[field].push(...Object.keys(s.input.sheets[b.i][field] ?? {}), ...Object.keys((b.set[field] as object | undefined) ?? {}))
  const out: LayoutEdit[] = [...keys.rows.map((rows) => ({ sheet, rows, set: {} })), ...keys.columns.map((columns) => ({ sheet, columns, set: {} }))]
  if ('merges' in b.set || b.unset.includes('merges')) out.push({ sheet, merges: (b.set.merges as string[] | undefined) ?? [] })
  return out
}

/** One step taken on the session's book (a step of several, each of them from the last): what it changed and the way to undo it. */
function* take(core: SessionCore, step: Step): Steps<[Taken, Step]> {
  const s = yield* core.ready()
  switch (step.kind) {
    case 'cells': {
      const [outcome, was] = yield* core.putCells(s, step.edits, true)
      const cells = step.edits.map((e) => ({ sheet: s.sheetNames[core.sheetIndex(s)(e.sheet)!], cell: e.cell }))
      return [{ call: step.call, outcome, cells }, cellsStep(step.call, was)]
    }
    case 'layout': {
      const next = new Map(step.sheets.map((b) => [b.i, sheetWith(s.input.sheets[b.i], b)]))
      const shown = core.display ? snapshot(s, layoutCells(s, step.sheets.flatMap((b) => touched(s, b)))) : null
      const [outcome, before] = yield* core.relayout(s, next, shown)
      const t: Taken = { call: step.call, outcome }
      // a sheet hidden or shown again: its cells stay where they are
      if (step.call === 'editStructure') t.mapCell = checkedMapCell({ geometry: [], before: s.sheetNames, after: s.sheetNames })
      return [t, layoutStep(step.call, before, next)]
    }
    case 'styles': {
      const was = s.input.styles
      if (step.styles) s.input.styles = step.styles
      else delete s.input.styles
      return [{ call: step.call, outcome: { ...nothing(), styles: true } }, stylesStep(was, step.call)]
    }
    case 'book': {
      const { geometry: g } = step.diff
      const [book, before] = [yield* applyDiff(s.input, step.diff), s.sheetNames]
      const after = namesOf(book)
      const [, outcome, diff] = yield* core.rebuild(s, book, g, step.call !== 'setNames', true, true)
      const t: Taken = { call: step.call, outcome: { ...outcome, moved: movesOf(g, before, after) } }
      if (step.call === 'editStructure') t.mapCell = checkedMapCell({ geometry: g, before, after })
      return [t, bookStep(step.call, diff!, book)]
    }
    case 'many': {
      let outcome = nothing()
      let cells: CellRef[] | undefined
      const others: Step[] = []
      for (const x of step.steps) {
        const [one, other] = yield* take(core, x)
        outcome = joined(outcome, one.outcome)
        if (one.cells) cells = [...(cells ?? []), ...one.cells]
        others.push(other)
      }
      return [{ call: step.call, outcome, ...(cells ? { cells } : {}) }, manyStep(step.call, others)]
    }
  }
}

/**
 * Takes a step of the history (undo or redo, the same way): the book changed back by it and recomputed; `done` gets the
 * way to the book before it, the step of the other stack. Nothing changes if it stops.
 */
export function* applyStep(core: SessionCore, step: Step | undefined, done: (other: Step | null) => void): Steps<Taken | null> {
  if (!step) return null
  const [t, other] = yield* take(core, step)
  done(other)
  return t
}
