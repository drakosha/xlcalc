/**
 * A cell's data validation in a computed book (`Session.getValidation`, `Session.validateEntry`): the rule over the cell,
 * its formulas as at the cell, a list's items, and whether a value meets the rule as Excel's `Validation.Value` checks
 * it. A rule's formulas are computed at the cell as a formula of the cell would be, a reference staying a reference.
 */
import type { CellInput, CellReplace, CellValidation, DataValidation, NumberFormat } from '../contract.js'
import type { Arg } from '../args.js'
import { applyEdits, codegenContext, invalidError, toCellValue, valueAt, type EngineState } from '../engine.js'
import { closureBackend } from '../exec/closures.js'
import { walk } from '../exec/walk.js'
import type { Steps } from '../exec/budget.js'
import { formatOf as lineFormat } from '../layout/lines.js'
import { lineAt } from '../layout/spans.js'
import { nameFor } from '../names.js'
import { formatAddress } from '../model.js'
import { parse, sheetKey, validate, walk as walkAst } from '../syntax.js'
import { decode } from '../structure/area.js'
import { tableAt, tableNames } from '../tables/model.js'
import { Arr } from '../arrays/value.js'
import { E, isErr, Rng, type V } from '../values.js'
import { anchorOf, ruleAt, twoBounds } from './rules.js'
import { formulaAt } from './structure.js'

/** A formula of a rule computed at cell (r, c) of sheet `s`: a reference stays one; a formula that does not compute is its error. */
function computeAt(state: EngineState, s: number, r: number, c: number, text: string): Arg {
  const known = (name: string, sheet: string | undefined) => !!nameFor(state.names, sheet === undefined ? s : state.sheetIdx.get(sheetKey(sheet))!, name)
  const enclosing = tableAt(state.tables, s, r, c)?.name
  const v = validate(text, { sheets: state.sheetNames, known, externalFunctions: state.externalFunctions, tables: tableNames(state.tables), ...(enclosing ? { enclosing } : {}) })
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

/** The number format a cell shows by (`layout/lines.ts`). */
function formatOf(state: EngineState, s: number, r: number, c: number): NumberFormat | undefined {
  const sh = state.input.sheets[s]
  return lineFormat(sh.cells[formatAddress(r, c)], lineAt(sh.rows, 'r', r), lineAt(sh.columns, 'c', c))
}

/** A literal list (a text constant): its items, split at commas as a file holds them in every locale. */
const literal = (text: string): string[] | null => (/^\s*"(?:[^"]|"")*"\s*$/.test(text) ? text.trim().slice(1, -1).replace(/""/g, '"').split(',') : null)

type Item = { value: V; text: string }

/**
 * A list's items: a literal's (each without its leading spaces, as Excel takes them: the set `lists`), or the cells of the
 * reference (or the values of the array) its formula gives, blank cells past the last value left out unless `blanks`; an
 * error.
 */
function listItems(state: EngineState, s: number, r: number, c: number, text: string, blanks = false): Item[] | { error: V } {
  const items = literal(text)
  if (items) return items.map((t) => ({ value: t.trimStart(), text: t.trimStart() }))
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
  while (!blanks && out.length && out[out.length - 1].value === null) out.pop()
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

/** How a value compares with the bounds, by the order of Excel's comparison operators (`a`, `b`: the value against each bound). */
const COMPARE: Record<string, (a: number, b: number) => boolean> = {
  // bounds in the wrong order take nothing (the set `entry-held`, column C), where a conditional format's take both ways
  between: (a, b) => a >= 0 && b <= 0,
  notBetween: (a, b) => a < 0 || b > 0,
  equal: (a) => a === 0,
  notEqual: (a) => a !== 0,
  greaterThan: (a) => a > 0,
  lessThan: (a) => a < 0,
  greaterThanOrEqual: (a) => a >= 0,
  lessThanOrEqual: (a) => a <= 0,
}

/**
 * Whether a list's item is the value: a cell's value of the same type, text in any case; a literal item, text in its
 * case, against the value as text (a number as `&` writes it: `1e3` and `+2` are no 1000 and 2), without spaces at
 * either end (the sets `files`, `entry-held`, `lists`, `entry-held-2`, `held-3`).
 */
function listed(state: EngineState, value: V, item: Item, isLiteral: boolean): boolean {
  const rules = state.ctx.rules
  if (!isLiteral) return typeof item.value === typeof value && !isErr(item.value) && rules.compare(item.value, value) === 0
  const text = rules.toText(value)
  return typeof text === 'string' && text.trim() === item.text
}

/** Whether a formula as at the cell refers to a blank cell by a reference to one cell, which with blanks ignored lets any value in. */
function refersToBlank(state: EngineState, s: number, text: string): boolean {
  let ast
  try {
    ast = parse(text)
  } catch {
    return false
  }
  let blank = false
  walkAst(ast, (n) => {
    if (blank || n.k !== 'ref' || n.ref.kind !== 'cell') return
    const i = n.sheet === null ? s : state.sheetIdx.get(sheetKey(n.sheet))
    const a = decode(n.ref)
    if (i !== undefined && valueAt(state, i, a.r.lo.i, a.c.lo.i) === null) blank = true
  })
  return blank
}

/** Whether a list's formula is a defined name alone. */
const isName = (text: string) => {
  try {
    return parse(text).k === 'name'
  } catch {
    return false
  }
}

/** Whether the cell's value meets its rule (`v`), in the book as it is now. */
function meets(state: EngineState, s: number, r: number, c: number, v: DataValidation): boolean {
  const value = valueAt(state, s, r, c)
  const blankOk = v.allowBlank !== false
  if (v.type === 'any') return true
  // a blank cell where blanks are not ignored is checked as a blank: 0, a text of length 0, in no list (the set `entry-held-2`)
  if (value === null && blankOk) return true
  if (isErr(value)) return false
  const f = formulas(state, s, r, c, v)
  const texts = [f.formula1!, ...(twoBounds(v) ? [f.formula2!] : [])]
  // a formula referring to a blank cell lets anything in when blanks are ignored (the sets `entry-held`, `struct-held`)
  if (blankOk && v.type !== 'list' && texts.some((t) => refersToBlank(state, s, t))) return true
  if (v.type === 'custom') {
    const x = computeAt(state, s, r, c, f.formula1!)
    const y = x instanceof Rng || x instanceof Arr ? state.rt.scalar(x) : (x as V)
    return y === null ? blankOk : y === true || (typeof y === 'number' && y !== 0)
  }
  if (v.type === 'list') {
    const isLiteral = literal(f.formula1!) !== null
    const items = listItems(state, s, r, c, f.formula1!, true)
    if ('error' in items) return false
    // a list through a defined name with a blank among its cells takes anything when blanks are ignored (the set `lists`)
    if (blankOk && isName(f.formula1!) && items.some((i) => i.value === null)) return true
    return items.some((i) => i.value !== null && listed(state, value, i, isLiteral))
  }
  const bounds = texts.map((t) => bound(state, computeAt(state, s, r, c, t)))
  if (bounds.includes('blank') && blankOk) return true
  if (bounds.includes(null)) return false
  const [a, b] = bounds.map((x) => (x === 'blank' ? 0 : x)) as number[]
  let n: V
  if (v.type === 'textLength') {
    const t = value === null ? '' : state.ctx.rules.toText(value)
    if (typeof t !== 'string') return false
    n = t.length
  } else {
    // a number, or a boolean, which compares above every number as Excel's operators compare (the set `entry-held`); text does not pass
    if (value !== null && typeof value !== 'number' && typeof value !== 'boolean') return false
    if (v.type === 'whole' && typeof value === 'number' && !Number.isInteger(value)) return false
    n = value ?? 0
  }
  const cmp = (x: number) => Math.sign(state.ctx.rules.compareOp(n, x))
  return COMPARE[v.operator ?? 'between'](cmp(a), b === undefined ? 0 : cmp(b))
}

/** The rule over a cell that its value, in the book as it is now, does not meet; null when it meets it or there is none. */
export function brokenRule(state: EngineState, s: number, r: number, c: number): DataValidation | null {
  const rule = ruleOf(state, s, r, c)
  return rule && !meets(state, s, r, c, rule) ? rule : null
}

/** Whether a cell has a rule. */
export const hasRule = (state: EngineState, s: number, r: number, c: number): boolean => ruleOf(state, s, r, c) !== null

/**
 * What `read` gives with the cell holding `input`: the input put into the book, read, and taken back (both recompute
 * what depends on the cell).
 */
export const withInput = <T>(state: EngineState, s: number, r: number, c: number, input: CellInput | null, read: () => T): Steps<T> =>
  withInputs(state, [{ sheet: state.sheetNames[s], cell: formatAddress(r, c), input }], read)

/** What `read` gives with the cells holding the inputs of `edits` (each cell once), as `withInput`. */
export function* withInputs<T>(state: EngineState, edits: CellReplace[], read: () => T): Steps<T> {
  const put: CellReplace[] = []
  const back: CellReplace[] = []
  for (const e of edits) {
    const old = state.input.sheets[state.sheetIdx.get(sheetKey(e.sheet))!].cells[e.cell]
    if (old === e.input || JSON.stringify(old ?? null) === JSON.stringify(e.input)) continue
    put.push(e)
    back.push({ ...e, input: old ?? null })
  }
  if (put.length) yield* applyEdits(state, put)
  try {
    return read()
  } finally {
    if (back.length) yield* applyEdits(state, back)
  }
}
