/**
 * What a session's edit answers (`EditResult`): what the call changed, gathered while it runs (`Outcome`), and the
 * session's revision after it. Beside the cells whose value or shown text changed, a cell is listed when its formula,
 * its own style or its own number format changed: those of the edited cells, and after the book was built again (a
 * structure operation, a table's column renamed) those that differ from the cell that was at their old place.
 */
import type { AddressMove, BookInput, CellChange, CellInput, CellRef, EditOptions, RangeRef, SheetInput } from '../contract.js'
import { key, type EngineState } from '../engine.js'
import { CalcError, paced, type Steps } from '../exec/budget.js'
import { InputError, object, onlyKeys } from '../input.js'
import { parseAddress } from '../model.js'
import { sheetKey } from '../syntax.js'

/** What changed of a cell beside its value and shown text. */
export type FieldChange = Pick<CellChange, 'formula' | 'style' | 'numberFormat'>

/** What a call changed so far: the cells to list (`s:r:c`, in order) and the rest a host reads again. */
export interface Outcome {
  keys: string[]
  fields: Map<string, FieldChange>
  /** Some cell's value changed: the cells of conditional formats may look otherwise. */
  valued: boolean
  /** Sheets whose layout changed, by name. */
  layout: Set<string>
  styles: boolean
  moved: AddressMove[]
  /** The cells of rules of conditional formats that changed, as they were. */
  rules: RangeRef[]
}

export const nothing = (): Outcome => ({ keys: [], fields: new Map(), valued: false, layout: new Set(), styles: false, moved: [], rules: [] })

/** Keys as the engine orders them, then the others in address order. */
export function inOrder(keys: string[], more: Iterable<string>): string[] {
  const seen = new Set(keys)
  const extra: number[][] = []
  for (const k of more)
    if (!seen.has(k)) {
      seen.add(k)
      extra.push(k.split(':').map(Number))
    }
  return [...keys, ...extra.sort((a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2]).map((k) => k.join(':'))]
}

/** Two outcomes of one call, the second after the first. */
export function joined(a: Outcome, b: Outcome): Outcome {
  const fields = new Map(a.fields)
  for (const [k, f] of b.fields) fields.set(k, { ...fields.get(k), ...f })
  return {
    keys: inOrder(a.keys, b.keys),
    fields,
    valued: a.valued || b.valued,
    layout: new Set([...a.layout, ...b.layout]),
    styles: a.styles || b.styles,
    moved: [...a.moved, ...b.moved],
    rules: [...a.rules, ...b.rules],
  }
}

/** A cell's formula as a result gives it: canonical when it parses, else as written. */
export function formulaText(s: EngineState, i: number, r: number, c: number, input: CellInput): string | undefined {
  const id = s.at.get(key(i, r, c))
  return id === undefined ? input.formula : s.formulas[id]!.text
}

/** What differs between a cell's input before and now: its formula (as now), own style and own number format. */
function fieldsOf(s: EngineState, i: number, r: number, c: number, was: CellInput | undefined, now: CellInput | undefined): FieldChange | null {
  if (was === now) return null
  const out: FieldChange = {}
  if (was?.formula !== now?.formula) out.formula = now?.formula === undefined ? null : formulaText(s, i, r, c, now)
  if (was?.style !== now?.style) out.style = now?.style ?? null
  if (was?.format !== now?.format) out.numberFormat = now?.format ?? null
  return 'formula' in out || 'style' in out || 'numberFormat' in out ? out : null
}

/** The fields that changed of edited cells, from their inputs before the edits (`was`) to the book's now. */
export function editFields(s: EngineState, was: { sheet: string; cell: string; input: CellInput | null }[]): Map<string, FieldChange> {
  const out = new Map<string, FieldChange>()
  for (const { sheet, cell, input } of was) {
    const i = s.sheetIdx.get(sheetKey(sheet))!
    const { row, col } = parseAddress(cell)
    const f = fieldsOf(s, i, row, col, input ?? undefined, s.input.sheets[i].cells[cell])
    if (f) out.set(`${i}:${row}:${col}`, f)
  }
  return out
}

/**
 * The cells of a book built again (`next`) whose formula (as each state shows it), own style or own number format
 * differ from those of the cell `mapCell` brought there from the old book (none: the cell stays where it is), by key of
 * `next`; a cell nothing came to differs from an empty one.
 */
export function* movedFields(old: EngineState, next: EngineState, mapCell: ((ref: CellRef) => CellRef | null) | null): Steps<Map<string, FieldChange>> {
  const out = new Map<string, FieldChange>()
  // the cells of each sheet of `next` some old cell came to, by address
  const reached = next.input.sheets.map(() => new Set<string>())
  const textOf = (s: EngineState, i: number, addr: string, input: CellInput | undefined) => {
    if (input?.formula === undefined) return undefined
    const { row, col } = parseAddress(addr)
    return formulaText(s, i, row, col, input)
  }
  const compare = (j: number, addr: string, was: CellInput | undefined, wasText: string | undefined, now: CellInput | undefined) => {
    const nowText = textOf(next, j, addr, now)
    const f: FieldChange = {}
    if (wasText !== nowText) f.formula = nowText ?? null
    if (was?.style !== now?.style) f.style = now?.style ?? null
    if (was?.format !== now?.format) f.numberFormat = now?.format ?? null
    if (!('formula' in f || 'style' in f || 'numberFormat' in f)) return
    const { row, col } = parseAddress(addr)
    out.set(`${j}:${row}:${col}`, f)
  }
  for (let s = 0; s < old.input.sheets.length; s++) {
    const cells = old.input.sheets[s].cells
    for (const addr in cells) {
      if (paced(1)) yield
      const to = mapCell ? mapCell({ sheet: old.sheetNames[s], cell: addr }) : { sheet: old.sheetNames[s], cell: addr }
      if (!to) continue
      const j = next.sheetIdx.get(sheetKey(to.sheet))
      if (j === undefined) continue
      reached[j].add(to.cell)
      const [was, now] = [cells[addr], next.input.sheets[j].cells[to.cell]]
      // the same input without a formula: nothing of it can differ (a formula's text may, with new names)
      if (was === now && was.formula === undefined) continue
      compare(j, to.cell, was, textOf(old, s, addr, was), now)
    }
  }
  for (let j = 0; j < next.input.sheets.length; j++) {
    const cells = next.input.sheets[j].cells
    for (const addr in cells) {
      if (paced(1)) yield
      if (!reached[j].has(addr)) compare(j, addr, undefined, undefined, cells[addr])
    }
  }
  return out
}

/** Whether two sheets differ in anything but their cells. */
function layoutDiffers(a: SheetInput, b: SheetInput): boolean {
  for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) {
    if (k === 'cells') continue
    const [x, y] = [(a as unknown as Record<string, unknown>)[k], (b as unknown as Record<string, unknown>)[k]]
    if (x !== y && JSON.stringify(x) !== JSON.stringify(y)) return true
  }
  return false
}

/** The sheets of `next` whose layout differs from that of the sheet of `old` with their name (a new or renamed sheet too). */
export function layoutChanged(old: BookInput, next: BookInput): string[] {
  const was = new Map(old.sheets.map((s) => [sheetKey(s.name), s]))
  return next.sheets.filter((s) => {
    const o = was.get(sheetKey(s.name))
    return !o || layoutDiffers(o, s)
  }).map((s) => s.name)
}

/** The rectangles of a sheet's conditional formats. */
export const ruleRanges = (sheet: SheetInput): RangeRef[] => (sheet.conditionalFormats ?? []).flatMap((f) => f.ranges.map((range) => ({ sheet: sheet.name, range })))

/** What a host paints again: the cells of the rules that changed, and of every rule when a value changed; none without rules. */
export function repaintOf(book: BookInput, o: Outcome): RangeRef[] | undefined {
  if (!o.valued && !o.rules.length) return undefined
  const out = new Map<string, RangeRef>()
  for (const x of [...o.rules, ...(o.valued ? book.sheets.flatMap(ruleRanges) : [])]) out.set(`${sheetKey(x.sheet)}!${x.range}`, x)
  return out.size ? [...out.values()] : undefined
}

/** The keys of cells in address order, once each. */
export const sortedKeys = (keys: Iterable<string>): string[] =>
  [...new Set(keys)]
    .map((k) => k.split(':').map(Number))
    .sort((a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2])
    .map((k) => k.join(':'))

/** An edit's options checked (`allowed` beside `baseRevision`), and the session's revision against its `baseRevision`. */
export function checkBase<T extends EditOptions>(v: unknown, revision: number, allowed: readonly string[] = [], path = 'opts'): T {
  if (v === undefined) return {} as T
  const o = object(v, path, 'options such as { baseRevision: 3 }')
  onlyKeys(o, path, ['baseRevision', ...allowed])
  const b = o.baseRevision
  if (b !== undefined && !(typeof b === 'number' && Number.isInteger(b) && b >= 0)) throw new InputError(`${path}.baseRevision`, `expected a revision, a whole number of at least 0, got ${JSON.stringify(b)}`)
  if (b !== undefined && b !== revision) throw new CalcError('stale_revision', `the edit is made over revision ${b}, the session is at revision ${revision}: it changed nothing`)
  return o as T
}
