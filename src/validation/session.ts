/**
 * A cell's data validation in a computed book (`Session.validation`, `Session.validateEntry`): the rule over the cell,
 * its formulas as at the cell, a list's items, and whether a value meets the rule as Excel's `Validation.Value` checks
 * it. A rule's formulas are computed at the cell as a formula of the cell would be, a reference staying a reference.
 */
import type { CellInput, CellValidation, DataValidation, EntryCheck, NumberFormat } from '../contract.js'
import type { Arg } from '../args.js'
import { applyEdits, codegenContext, invalidError, toCellValue, valueAt, type EngineState } from '../engine.js'
import { closureBackend } from '../exec/closures.js'
import { walk } from '../exec/walk.js'
import type { Steps } from '../exec/budget.js'
import { lineAt } from '../layout/spans.js'
import { nameFor } from '../names.js'
import { formatAddress } from '../model.js'
import { sheetKey, validate } from '../syntax.js'
import { tableAt } from '../tables/model.js'
import { Arr } from '../arrays/value.js'
import { E, isErr, Rng, type V } from '../values.js'
import { anchorOf, ruleAt, twoBounds } from './rules.js'
import { formulaAt } from './structure.js'

/** A formula of a rule computed at cell (r, c) of sheet `s`: a reference stays one; a formula that does not compute is its error. */
function computeAt(state: EngineState, s: number, r: number, c: number, text: string): Arg {
  const known = (name: string, sheet: string | undefined) => !!nameFor(state.names, sheet === undefined ? s : state.sheetIdx.get(sheetKey(sheet))!, name)
  const enclosing = tableAt(state.tables, s, r, c)?.name
  const tables = state.tables.list.map((t) => ({ name: t.name, columns: t.cols }))
  const v = validate(text, { sheets: state.sheetNames, known, externalFunctions: state.externalFunctions, tables, ...(enclosing ? { enclosing } : {}) })
  if (!v.ok) return E[invalidError(v.code)]
  return walk(codegenContext(state, s), v.ast, s, r, c, closureBackend(state.rt)).code()
}

/** The rule over a cell and the cell its formulas are written for, or null. */
function ruleOf(state: EngineState, s: number, r: number, c: number): DataValidation | null {
  const rules = state.input.sheets[s].validations
  const i = ruleAt(rules, r, c)
  return i < 0 ? null : rules![i]
}

/** The rule's formulas as at the cell. */
function formulas(state: EngineState, s: number, r: number, c: number, v: DataValidation): { formula1?: string; formula2?: string } {
  const names = new Map(state.sheetNames.map((n) => [sheetKey(n), n]))
  const at = (text: string) => formulaAt(text, state.sheetNames[s], anchorOf(v), { r, c }, names)
  return { ...(v.formula1 !== undefined ? { formula1: at(v.formula1) } : {}), ...(v.formula2 !== undefined ? { formula2: at(v.formula2) } : {}) }
}

/** The number format a cell shows by: its own, else its row's, else its column's. */
function formatOf(state: EngineState, s: number, r: number, c: number): NumberFormat | undefined {
  const sh = state.input.sheets[s]
  return sh.cells[formatAddress(r, c)]?.format ?? lineAt(sh.rows, 'r', r)?.format ?? lineAt(sh.columns, 'c', c)?.format
}

/** A literal list (a text constant): its items, split at commas as a file holds them in every locale. */
const literal = (text: string): string[] | null => (/^\s*"(?:[^"]|"")*"\s*$/.test(text) ? text.trim().slice(1, -1).replace(/""/g, '"').split(',') : null)

type Item = { value: V; text: string }

/** A list's items: a literal's, or the cells of the reference (or the values of the array) its formula gives; an error. */
function listItems(state: EngineState, s: number, r: number, c: number, text: string): Item[] | { error: V } {
  const items = literal(text)
  if (items) return items.map((t) => ({ value: t, text: t }))
  const x = computeAt(state, s, r, c, text)
  const shown = (v: V, at?: [number, number, number]) => String(state.ctx.rules.display(toCellValue(v), at && formatOf(state, ...at), Infinity).text)
  let out: Item[]
  if (x instanceof Rng) {
    out = []
    // a whole column is read as far as its sheet holds values
    const last = Math.min(x.r2, state.ctx.sheets[x.s].rows - 1)
    for (let row = x.r1; row <= last; row++) for (let col = x.c1; col <= x.c2; col++) {
      const v = valueAt(state, x.s, row, col)
      out.push({ value: v, text: v === null ? '' : shown(v, [x.s, row, col]) })
    }
  } else if (x instanceof Arr) out = x.values().map((v) => ({ value: v, text: v === null ? '' : shown(v) }))
  else if (isErr(x)) return { error: x }
  else if (x instanceof Object) return { error: E.VALUE }
  else out = [{ value: x, text: x === null ? '' : shown(x) }]
  // blank cells past the last value are left out
  while (out.length && out[out.length - 1].value === null) out.pop()
  return out
}

/** A cell's data validation: its rule, the rule's formulas as at it and a list's items computed now; null without a rule. */
export function cellValidation(state: EngineState, s: number, r: number, c: number): CellValidation | null {
  const rule = ruleOf(state, s, r, c)
  if (!rule) return null
  const f = formulas(state, s, r, c, rule)
  if (rule.type !== 'list') return { rule, ...f }
  const items = listItems(state, s, r, c, f.formula1!)
  const list = 'error' in items ? { error: isErr(items.error) ? items.error.code : ('VALUE' as const) } : items.map((i) => ({ value: toCellValue(i.value), text: i.text }))
  return { rule, ...f, list }
}

/** A bound as a number: a number, text that reads as one; blank as `blank`; anything else null. */
function bound(state: EngineState, x: Arg): number | 'blank' | null {
  const v = x instanceof Rng || x instanceof Arr ? state.rt.scalar(x) : (x as V)
  if (v === null) return 'blank'
  if (typeof v === 'number') return v
  if (typeof v === 'string') return state.ctx.rules.valueText(v)
  return null
}

const COMPARE: Record<string, (v: number, a: number, b: number) => boolean> = {
  between: (v, a, b) => v >= a && v <= b,
  notBetween: (v, a, b) => v < a || v > b,
  equal: (v, a) => v === a,
  notEqual: (v, a) => v !== a,
  greaterThan: (v, a) => v > a,
  lessThan: (v, a) => v < a,
  greaterThanOrEqual: (v, a) => v >= a,
  lessThanOrEqual: (v, a) => v <= a,
}

/**
 * Whether a list's item is the value: a cell's value of the same type, text in any case; a literal item as written, text
 * in its case (the set `files`), a number the item reads as, a boolean by its word.
 */
function listed(state: EngineState, value: V, item: Item, isLiteral: boolean): boolean {
  const rules = state.ctx.rules
  if (!isLiteral) return typeof item.value === typeof value && !isErr(item.value) && rules.compare(item.value, value) === 0
  const t = item.text
  if (typeof value === 'number') return rules.valueText(t) === value
  if (typeof value === 'boolean') return rules.toText(value) === t
  return value === t
}

/** Whether the cell's value meets its rule (`v`), in the book as it is now. */
function meets(state: EngineState, s: number, r: number, c: number, v: DataValidation): boolean {
  const value = valueAt(state, s, r, c)
  const blankOk = v.allowBlank !== false
  if (v.type === 'any') return true
  if (value === null) return blankOk
  if (isErr(value)) return false
  const f = formulas(state, s, r, c, v)
  if (v.type === 'custom') {
    const x = computeAt(state, s, r, c, f.formula1!)
    const y = x instanceof Rng || x instanceof Arr ? state.rt.scalar(x) : (x as V)
    return y === null ? blankOk : y === true || (typeof y === 'number' && y !== 0)
  }
  if (v.type === 'list') {
    const isLiteral = literal(f.formula1!) !== null
    const items = listItems(state, s, r, c, f.formula1!)
    if ('error' in items) return false
    // a list of blank cells only takes anything when blanks are ignored
    if (blankOk && items.every((i) => i.value === null)) return true
    return items.some((i) => i.value !== null && listed(state, value, i, isLiteral))
  }
  const bounds = [f.formula1!, ...(twoBounds(v) ? [f.formula2!] : [])].map((t) => bound(state, computeAt(state, s, r, c, t)))
  // a bound on a blank cell lets anything in when blanks are ignored
  if (bounds.includes('blank') && blankOk) return true
  if (bounds.includes(null)) return false
  const [a, b] = bounds.map((x) => (x === 'blank' ? 0 : x)) as number[]
  let n: number
  if (v.type === 'textLength') {
    const t = state.ctx.rules.toText(value)
    if (typeof t !== 'string') return false
    n = t.length
  } else {
    if (typeof value !== 'number') return false
    if (v.type === 'whole' && !Number.isInteger(value)) return false
    n = value
  }
  return COMPARE[v.operator ?? 'between'](n, a, b)
}

/**
 * Whether the cell holding `input` meets its rule: the input put into the book, the rule checked, the input taken back
 * (both recompute what depends on the cell). A cell without a rule takes anything.
 */
export function* checkEntry(state: EngineState, s: number, r: number, c: number, input: CellInput | null): Steps<EntryCheck> {
  const rule = ruleOf(state, s, r, c)
  if (!rule) return { ok: true }
  const cells = state.input.sheets[s].cells
  const addr = formatAddress(r, c)
  const old = cells[addr]
  const same = old === input || JSON.stringify(old ?? null) === JSON.stringify(input)
  const sheet = state.sheetNames[s]
  if (!same) yield* applyEdits(state, [{ sheet, cell: addr, input }])
  try {
    return meets(state, s, r, c, rule) ? { ok: true } : { ok: false, rule }
  } finally {
    if (!same) yield* applyEdits(state, [{ sheet, cell: addr, input: old ?? null }])
  }
}
