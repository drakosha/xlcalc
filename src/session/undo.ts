/**
 * A session's history (`Session.undo`): each step holds the way back of one call, what that call replaced and no more:
 * the inputs of the cells it changed, the fields and the lines of the sheets whose layout it changed, the style table,
 * or the diff of a book built again (`rebook.ts`). Undoing a step makes the way forth from what it replaces, so the two
 * stacks hold one kind of step. Steps are kept by `depth` and by an estimate of their memory.
 */
import type { BookInput, CellChange, CellEdit, CellRef, CellStyle, LayoutEdit, SessionHistory, SheetInput, UndoCall, UndoOptions, UndoResult } from '../contract.js'
import type { EngineState } from '../engine.js'
import type { Steps } from '../exec/budget.js'
import { checkedMapCell } from '../structure/index.js'
import { sheetKey } from '../syntax.js'
import { layoutCells, snapshot } from './looks.js'
import { applyDiff, bookDiff, GEOMETRY, mapCellOf, type BookDiff, type Geometry } from './rebook.js'

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
  | { kind: 'cells'; call: 'setCells' | 'patchCells'; edits: CellEdit[]; size: number }
  | { kind: 'layout'; call: 'setLayout' | 'restructure'; sheets: SheetBack[]; size: number }
  | { kind: 'styles'; call: 'setStyles'; styles: CellStyle[] | undefined; size: number }
  | { kind: 'book'; call: Exclude<UndoCall, 'patchCells' | 'setStyles' | 'setLayout'>; diff: BookDiff; size: number }

/** What a session's history holds, to restore it elsewhere (a pool's snapshot). */
export interface HistoryData {
  undo: Step[]
  redo: Step[]
}

const DEPTH = 100
const MAX_MB = 64

/**
 * A rough size in bytes of what a step keeps: V8's objects (a header and a slot per field), strings at two bytes a
 * character; shared objects are counted each time, so it errs high.
 */
export function weigh(v: unknown): number {
  if (v === null || v === undefined) return 8
  switch (typeof v) {
    case 'string':
      return 16 + 2 * v.length
    case 'object': {
      let n = 32
      if (Array.isArray(v)) for (const x of v) n += 8 + weigh(x)
      else for (const k in v) n += 24 + 2 * k.length + weigh((v as Record<string, unknown>)[k])
      return n
    }
    default:
      return 16
  }
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

/** The way from sheet `from` to sheet `to` (the same sheet before and after a change of its layout), but its cells. */
export function sheetBack(i: number, from: SheetInput, to: SheetInput): SheetBack | null {
  const set: Record<string, unknown> = {}
  const unset: string[] = []
  const lines: SheetBack['lines'] = []
  for (const k of new Set([...Object.keys(from), ...Object.keys(to)])) {
    const [a, b] = [(from as unknown as Record<string, unknown>)[k], (to as unknown as Record<string, unknown>)[k]]
    if (k === 'cells' || a === b) continue
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

/** The size of a layout step. */
export const layoutSize = (sheets: SheetBack[]): number => sheets.reduce((n, s) => n + weigh(s.set) + weigh(s.lines), 0)

/** The fields of `a` that `b` does not share, as `weigh` counts them. */
const own = (a: object, b: object | undefined): number => {
  let n = 32
  for (const [k, v] of Object.entries(a)) if (v !== (b as Record<string, unknown> | undefined)?.[k]) n += 24 + weigh(v)
  return n
}

/** The size of a book's diff from `current`: the fields and cells it keeps that `current` does not share. */
export function diffSize(d: BookDiff, current: BookInput): number {
  const sheets = new Map(current.sheets.map((s) => [sheetKey(s.name), s]))
  return own(d.head, current) + weigh(d.op) + d.sheets.reduce((n, x) => n + ('rest' in x ? own(x.rest, sheets.get(sheetKey(x.rest.name))) + weigh(x.put) : 32), 0)
}

/** The size of a step of cells: the inputs it keeps. */
export const cellsSize = (edits: CellEdit[]): number => edits.reduce((n, e) => n + 48 + weigh(e.input), 0)

/** What a step needs of its session (`session.ts`): the ways it changes the book, which forward calls take too. */
export interface SessionCore {
  ready(): Steps<EngineState>
  putCells(s: EngineState, edits: CellEdit[]): Steps<[CellChange[], CellEdit[]]>
  rebuild(old: EngineState, book: BookInput, mapCell: (ref: CellRef) => CellRef | null, looks: boolean, back: () => Steps<BookDiff | null>): Steps<[EngineState, CellChange[], BookDiff | null]>
  relayout(s: EngineState, next: Map<number, SheetInput>, shown: Map<string, string> | null): Steps<CellChange[]>
  display: boolean
  sheetIndex(s: EngineState): (name: string) => number | undefined
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

/** A mapCell with its geometry: a pool's thread sends the geometry, the calling side makes the same mapCell. */
function registered(mapCell: (ref: CellRef) => CellRef | null, g: Geometry) {
  GEOMETRY.set(mapCell, g)
  return mapCell
}

/**
 * Takes a step of the history (undo or redo, the same way): the book changed back by it and recomputed; `done` gets the
 * way to the book before it, the step of the other stack (null when it is too large to make). Nothing changes if it stops.
 */
export function* applyStep(core: SessionCore, step: Step | undefined, done: (other: Step | null) => void): Steps<UndoResult | null> {
  if (!step) return null
  const s = yield* core.ready()
  let result: UndoResult
  let other: Step
  switch (step.kind) {
    case 'cells': {
      const [changes, was] = yield* core.putCells(s, step.edits)
      result = { call: step.call, changes, cells: step.edits.map((e) => ({ sheet: s.sheetNames[core.sheetIndex(s)(e.sheet)!], cell: e.cell })) }
      other = { kind: 'cells', call: step.call, edits: was, size: cellsSize(was) }
      break
    }
    case 'layout': {
      const before = s.input.sheets.slice()
      const next = new Map(step.sheets.map((b) => [b.i, sheetWith(before[b.i], b)]))
      const shown = core.display ? snapshot(s, layoutCells(s, step.sheets.flatMap((b) => touched(s, b)))) : null
      result = { call: step.call, changes: yield* core.relayout(s, next, shown) }
      // a sheet hidden or shown again: its cells stay where they are
      if (step.call === 'restructure') result.mapCell = registered(checkedMapCell(s.input, (ref) => ref), { op: null, way: 'back', before: s.sheetNames, after: s.sheetNames })
      const backs = step.sheets.map((b) => sheetBack(b.i, next.get(b.i)!, before[b.i])).filter((x) => x !== null)
      other = { kind: 'layout', call: step.call, sheets: backs, size: layoutSize(backs) }
      break
    }
    case 'styles': {
      const was = s.input.styles
      if (step.styles) s.input.styles = step.styles
      else delete s.input.styles
      result = { call: 'setStyles', changes: [] }
      other = { kind: 'styles', call: 'setStyles', styles: was, size: weigh(was) }
      break
    }
    case 'book': {
      const d = step.diff
      const book = yield* applyDiff(s.input, d)
      const names = book.sheets.map((x) => x.name)
      const map = mapCellOf(d.op, d.way, names)
      const [, changes, diff] = yield* core.rebuild(s, book, map, step.call !== 'setNames', () => bookDiff(book, s.input, d.op, d.way === 'back' ? 'forth' : 'back'))
      result = { call: step.call, changes }
      if (step.call === 'restructure') result.mapCell = registered(checkedMapCell(s.input, map), { op: d.op, way: d.way, before: s.sheetNames, after: names })
      other = { kind: 'book', call: step.call, diff: diff!, size: diffSize(diff!, book) }
      break
    }
  }
  done(other)
  return result
}
