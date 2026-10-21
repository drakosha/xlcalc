/**
 * What a session's edit answers (`EditResult`): what the call changed, gathered while it runs (`Outcome`), and the
 * session's revision after it. Beside the cells whose value or shown text changed, a cell is listed when its formula,
 * its own style, number format, note or hyperlink changed: those of the edited cells, and after the book was built again (a
 * structure operation, a table's column renamed) those that differ from the cell that was at their old place.
 */
import type { AddressMove, BookInput, CellChange, CellInput, CellRef, EditOptions, RangeRef, SheetInput } from '../contract.js'
import { sheetIndexOf, type EngineState } from '../engine.js'
import { CalcError } from '../exec/budget.js'
import { InputError } from '../input-error.js'
import { object, onlyOptions } from '../input/basic.js'
import { inputIn } from '../looks/sheet.js'
import { parseAddress, placeOf } from '../model.js'
import { sheetKey } from '../syntax.js'

/** What changed of a cell beside its value and shown text. */
export type FieldChange = Pick<CellChange, 'formula' | 'style' | 'numberFormat' | 'note' | 'link'>

/** What a call changed so far: the cells to list (their places, `placeOf`, in order) and the rest a host reads again. */
export interface Outcome {
  keys: number[]
  fields: Map<number, FieldChange>
  /** Some cell's value changed: the cells of conditional formats may look otherwise. */
  valued: boolean
  /** Sheets whose layout changed, by name. */
  layout: Set<string>
  styles: boolean
  moved: AddressMove[]
  /** The cells of rules of conditional formats that changed, as they were. */
  rules: RangeRef[]
  /** Rectangles whose look changed, too large to list by cell (`LISTED_LOOKS`). */
  looks: RangeRef[]
}

export const nothing = (): Outcome => ({ keys: [], fields: new Map(), valued: false, layout: new Set(), styles: false, moved: [], rules: [], looks: [] })

/** Places as the engine orders them, then the others in address order (a place's own order). */
export function inOrder(keys: number[], more: Iterable<number>): number[] {
  let seen: Set<number> | undefined
  const extra: number[] = []
  for (const k of more) {
    seen ??= new Set(keys)
    if (!seen.has(k)) {
      seen.add(k)
      extra.push(k)
    }
  }
  return extra.length ? [...keys, ...extra.sort((a, b) => a - b)] : [...keys]
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
    looks: [...a.looks, ...b.looks],
  }
}

/** A cell's formula as a result gives it: canonical when it parses, else as written. */
export function formulaText(s: EngineState, i: number, r: number, c: number, input: CellInput): string | undefined {
  const id = s.at.get(placeOf(i, r, c))
  return id === undefined ? input.formula : s.formulas[id]!.text
}

/**
 * What differs of a cell beside its value between the input `was` and the input `now` at its place: its formula as each
 * state shows it (`wasText`, `nowText`), its own style, number format, note and hyperlink; null when nothing does.
 */
export function fieldsBetween(was: CellInput | undefined, wasText: string | undefined, now: CellInput | undefined, nowText: string | undefined): FieldChange | null {
  const out: FieldChange = {}
  if (wasText !== nowText) out.formula = nowText ?? null
  if (was?.style !== now?.style) out.style = now?.style ?? null
  if (was?.numberFormat !== now?.numberFormat) out.numberFormat = now?.numberFormat ?? null
  if (was?.note?.text !== now?.note?.text || was?.note?.author !== now?.note?.author) out.note = now?.note ? { ...now.note } : null
  if (was?.link?.url !== now?.link?.url || was?.link?.location !== now?.link?.location || was?.link?.tooltip !== now?.link?.tooltip) out.link = now?.link ? { ...now.link } : null
  return Object.keys(out).length ? out : null
}

/** Whether an input has a field `fieldsBetween` compares beside its value: an input of a value alone differs from none in nothing else. */
export const hasFields = (x: CellInput): boolean =>
  x.formula !== undefined || x.style !== undefined || x.numberFormat !== undefined || x.note !== undefined || x.link !== undefined

/** The fields that changed of edited cells, from their inputs before the edits (`was`) to the book's now. */
export function editFields(s: EngineState, was: { sheet: string; cell: string; input: CellInput | null }[]): Map<number, FieldChange> {
  const out = new Map<number, FieldChange>()
  for (const { sheet, cell, input } of was) {
    const i = sheetIndexOf(s, sheet)!
    const { row, col } = parseAddress(cell)
    const now = inputIn(s.input.sheets[i], cell, row, col)
    if (now === (input ?? undefined)) continue
    // a formula's text as written before, as the state shows it now: one that changes is listed
    const f = fieldsBetween(input ?? undefined, input?.formula, now, now?.formula === input?.formula ? input?.formula : now?.formula === undefined ? undefined : formulaText(s, i, row, col, now))
    if (f) out.set(placeOf(i, row, col), f)
  }
  return out
}

/** Whether two sheets differ in anything but their cells and their looks. */
function layoutDiffers(a: SheetInput, b: SheetInput): boolean {
  for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) {
    if (k === 'cells' || k === 'looks') continue
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

/**
 * What a host paints again: the rectangles whose look changed too widely to list by cell, the cells of the rules that
 * changed, and of every rule when a value changed; none without them.
 */
export function repaintOf(book: BookInput, o: Outcome): RangeRef[] | undefined {
  if (!o.valued && !o.rules.length && !o.looks.length) return undefined
  const out = new Map<string, RangeRef>()
  for (const x of [...o.looks, ...o.rules, ...(o.valued ? book.sheets.flatMap(ruleRanges) : [])]) out.set(`${sheetKey(x.sheet)}!${x.range}`, x)
  return out.size ? [...out.values()] : undefined
}

/** An edit's options checked (`allowed` beside `baseRevision`), and the session's revision against its `baseRevision`. */
export function checkBase<T extends EditOptions>(v: unknown, revision: number, allowed: readonly string[] = [], path = 'opts'): T {
  if (v === undefined) return {} as T
  const o = object(v, path, 'options such as { baseRevision: 3 }')
  onlyOptions(o, path, ['baseRevision', ...allowed])
  const b = o.baseRevision
  if (b !== undefined && !(typeof b === 'number' && Number.isInteger(b) && b >= 0)) throw new InputError(`${path}.baseRevision`, `expected a revision, a whole number of at least 0, got ${JSON.stringify(b)}`)
  if (b !== undefined && b !== revision) throw new CalcError('stale_revision', `the edit was made against revision ${b}, but the session is at revision ${revision}: nothing was changed`)
  return o as T
}
