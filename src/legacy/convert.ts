/**
 * Formulas of the old kind (an xlsx cell formula without dynamic array metadata) as the dynamic formulas Excel 365
 * shows for them: the implicit intersection operator `@` where Excel puts it, so the engine, which computes every
 * formula as a dynamic one, gives the old formula's value. Measured against Excel's own text (`formula2`) in the
 * sets of `bench/coverage/legacy/`.
 *
 * Excel computes an old formula by values: a range where one value goes takes the cell of the formula's row or
 * column, an array where one value is wanted gives its top-left value. Only an argument that takes an array keeps
 * ranges whole inside it, and so on into the calls under it until one stops it. Each argument has its class
 * (`params.ts`); the walk carries where it is:
 * - `V` one value is wanted: a range gets `@`, and so does an array, a call with an array answer or a reference of
 *   several cells;
 * - `R` inside an argument that takes a reference (SUM's): a range where one value goes gets `@`, an array stays;
 * - `A` inside an argument that takes an array: nothing gets `@`.
 */
import type { NameInput } from '../contract.js'
import { canonical, FormulaError, isRefOp, parse, type Ast } from '../syntax.js'
import { caseKey } from '../collate.js'
import { inputNameFor } from '../names.js'
import { object, onlyKeys } from '../input/basic.js'
import { strict } from '../strict-args.js'
import { PARAMS } from './params.js'

type Where = 'V' | 'R' | 'A'

/**
 * Where a node stands: `v` one value; `p` one value for a range, an array passed on (IFERROR's value); `f` the
 * top-left cell of a range, one value of an array (N, T); `r` a reference or an array; `a` an array.
 */
type Slot = 'v' | 'p' | 'f' | 'r' | 'a'

/**
 * What a converted node gives: one value, a reference of several cells, an array, or ROW and COLUMN over a
 * reference (an array only where an array is computed).
 */
type Kind = 'one' | 'range' | 'array' | 'rowcol'

interface Out {
  ast: Ast
  kind: Kind
}

/** Arguments whose class the sets show only in part: IFERROR's value and TYPE's pass an array on (`TYPE(IFS("x",1))` is 64), N and T take a range's first cell. */
const SLOTS: Record<string, Record<number, Slot>> = { IFERROR: { 0: 'p' }, IFNA: { 0: 'p' }, N: { 0: 'f' }, T: { 0: 'f' }, TYPE: { 0: 'p' } }

/** Calls that pass on one of their arguments as it is. */
const PASSING: Record<string, (i: number) => boolean> = { IF: (i) => i > 0, CHOOSE: (i) => i > 0, IFERROR: () => true, IFNA: () => true }

interface Params {
  classes: string
  rep: number
  array: boolean
}

const PARSED = new Map<string, Params | undefined>()

/** A function's entry of `params.ts`: its arguments' classes, the number that repeat, whether it can answer an array. */
function paramsOf(fn: string): Params | undefined {
  if (!PARSED.has(fn)) {
    const m = PARAMS[fn] === undefined ? null : /^([a-zA-Z]*)(?:\/(\d+))?(!?)$/.exec(PARAMS[fn])
    PARSED.set(fn, m ? { classes: m[1], rep: Number(m[2] ?? 1), array: m[3] === '!' } : undefined)
  }
  return PARSED.get(fn)
}

/** The class of argument `i` of `fn`; repeated arguments take their place in the last group. */
function classAt(fn: string, i: number): string {
  const p = paramsOf(fn)
  if (!p) return 'v'
  const s = p.classes
  if (i < s.length) return s[i]
  return s[s.length - p.rep + ((i - s.length + p.rep) % p.rep)]
}

/** The slot and the evaluation of argument `i` of `fn` (the node `n`) called where `where` says. */
function argument(fn: string, i: number, n: Ast, where: Where): [Slot, Where] {
  const own = SLOTS[fn]?.[i]
  // under an argument that takes an array IFERROR's value takes a range as a reference, TYPE's whole (SUMPRODUCT(TYPE(A1:A3)) is 64)
  if (own) return [own, own === 'p' && where === 'A' && fn !== 'TYPE' ? 'R' : where]
  switch (classAt(fn, i)) {
    case 'v':
      return ['v', where]
    case 'w':
      return ['v', where === 'A' && n.k !== 'ref' ? 'R' : where]
    case 'V':
      return ['v', where === 'A' ? 'R' : where]
    case 'r':
      return ['r', where === 'A' ? 'A' : 'R']
    case 'R':
      return ['r', 'R']
    default:
      return ['a', 'A']
  }
}

/** Whether a node of `kind` in `slot` gets `@` where `where` says. */
function intersects(kind: Kind, slot: Slot, where: Where): boolean {
  if (kind === 'rowcol') return where !== 'A'
  if (kind === 'range') return (slot === 'v' || slot === 'p') && where !== 'A'
  if (kind === 'array') return (slot === 'v' || slot === 'f') && where === 'V'
  return false
}

const wider = (a: Kind, b: Kind): Kind => (a === 'range' || b === 'range' ? 'range' : a === 'array' || b === 'array' || a === 'rowcol' || b === 'rowcol' ? 'array' : 'one')

/** A positive whole number literal (INDEX's row or column that leaves one cell). */
const place = (n: Ast | undefined) => n?.k === 'num' && Number.isInteger(n.v) && n.v >= 1

/** The size of a literal reference: rows and columns, or null for anything else. */
function size(n: Ast): [number, number] | null {
  if (n.k === 'paren') return size(n.a)
  return n.k === 'ref' && !n.ref.spill ? [n.ref.r2 - n.ref.r1 + 1, n.ref.c2 - n.ref.c1 + 1] : null
}

/** A call's answer after its arguments are converted. */
function answer(fn: string, args: Out[], nodes: Ast[]): Kind {
  switch (fn) {
    case 'ROW':
    case 'COLUMN':
      return args.length && args[0].kind !== 'one' ? 'rowcol' : 'one'
    case 'INDEX': {
      const ref = args[0].kind === 'range' || nodes[0].k === 'ref'
      // of one value one value; of several, one cell when every index given is a number that leaves one
      const one = (args[0].kind === 'one' && nodes[0].k !== 'ref') || (place(nodes[1]) && (nodes.length < 3 ? ref : place(nodes[2])))
      return one ? 'one' : ref ? 'range' : 'array'
    }
    case 'VLOOKUP':
    case 'HLOOKUP':
      return args[2].kind === 'one' ? 'one' : 'array'
    case 'XLOOKUP': {
      const look = size(nodes[1])
      const ret = size(nodes[2])
      const line = look && ret && (look[1] === 1 ? ret[1] === 1 : ret[0] === 1)
      return line && (args[3]?.kind ?? 'one') === 'one' ? 'one' : ret ? 'range' : 'array'
    }
    case 'IFS':
    case 'SWITCH':
      return 'array'
  }
  const passing = PASSING[fn]
  if (passing) return args.reduce<Kind>((k, a, i) => (passing(i) ? wider(k, a.kind) : k), 'one')
  return paramsOf(fn)?.array ? 'array' : 'one'
}

/** A name's formula as a sheet sees it: its local name there, else the global one. */
type Names = (name: string, sheet: string | undefined) => { ast: Ast } | undefined

function converter(names: Names) {
  // the kind of each LET name, by slot; the names being read; the number of `@` put in
  let locals: Kind[] = []
  const reading = new Set<string>()
  let added = 0
  const go = (n: Ast, where: Where, slot: Slot, scope: string | undefined): Out => {
    const o = node(n, where, slot, scope)
    if (!intersects(o.kind, slot, where)) return o
    added++
    return { ast: { k: 'at', a: o.ast }, kind: 'one' }
  }
  const node = (n: Ast, where: Where, slot: Slot, scope: string | undefined): Out => {
    switch (n.k) {
      case 'ref':
        return { ast: n, kind: n.ref.kind === 'cell' && !n.ref.spill ? 'one' : 'range' }
      case 'ref3d':
      case 'union':
        return { ast: n, kind: 'range' }
      // one column of the formula's row, of the header or the totals row is one cell; the rest are ranges
      case 'table': {
        const one = n.spec.cols?.length === 1 && n.spec.areas.length === 1 && ['row', 'headers', 'totals'].includes(n.spec.areas[0])
        return { ast: n, kind: one ? 'one' : 'range' }
      }
      case 'arr':
        return { ast: n, kind: 'array' }
      case 'name':
        return { ast: n, kind: nameKind(n, scope) }
      case 'un':
      case 'pct': {
        const a = go(n.a, where, 'v', scope)
        return { ast: { ...n, a: a.ast }, kind: a.kind === 'one' ? 'one' : 'array' }
      }
      case 'bin': {
        // the reference operators take references whole and answer one
        if (isRefOp(n.op)) {
          const [a, b] = [go(n.a, 'R', 'r', scope), go(n.b, 'R', 'r', scope)]
          // an intersection with one cell is one cell at most
          return { ast: { ...n, a: a.ast, b: b.ast }, kind: n.op === ' ' && (a.kind === 'one' || b.kind === 'one') ? 'one' : 'range' }
        }
        const a = go(n.a, where, 'v', scope)
        const b = go(n.b, where, 'v', scope)
        return { ast: { ...n, a: a.ast, b: b.ast }, kind: a.kind === 'one' && b.kind === 'one' ? 'one' : 'array' }
      }
      // `(@A1:A3)`: the operand in its place gets `@`
      case 'paren': {
        const a = go(n.a, where, slot, scope)
        return { ast: { ...n, a: a.ast }, kind: a.kind }
      }
      case 'fn': {
        const args = n.args.map((a, i) => {
          const [s, w] = argument(n.name, i, a, where)
          return go(a, w, s, scope)
        })
        return { ast: { ...n, args: args.map((a) => a.ast) }, kind: answer(n.name, args, n.args) }
      }
      case 'let': {
        // the values and the calculation are references or arrays, as an argument of IF is
        const inner: Where = where === 'A' ? 'A' : 'R'
        const binds = n.binds.map((b) => {
          const v = go(b.value, inner, 'r', scope)
          locals[b.slot] = v.kind
          return { ...b, value: v.ast }
        })
        const body = n.body ? go(n.body, inner, 'r', scope) : undefined
        const kind = body ? body.kind : locals[n.binds[n.binds.length - 1].slot]
        return { ast: body ? { ...n, binds, body: body.ast } : { ...n, binds }, kind }
      }
      case 'local':
        return { ast: n, kind: locals[n.slot] ?? 'one' }
      default:
        return { ast: n, kind: 'one' }
    }
  }
  // a name gives what its formula gives, computed as an array as Excel computes a name
  const nameKind = (n: Extract<Ast, { k: 'name' }>, scope: string | undefined): Kind => {
    const sheet = n.sheet ?? scope
    const def = names(n.name, sheet)
    const key = `${sheet ?? ''}!${caseKey(n.name)}`
    if (!def || reading.has(key)) return 'one'
    reading.add(key)
    const outer = locals
    locals = []
    // a name's formula sees the global names
    const kind = node(def.ast, 'A', 'a', undefined).kind
    locals = outer
    reading.delete(key)
    return kind === 'rowcol' ? 'array' : kind
  }
  return (ast: Ast, sheet: string | undefined): Ast | undefined => {
    added = 0
    const out = go(ast, 'V', 'v', sheet).ast
    return added ? out : undefined
  }
}

/** What `fromLegacy` knows about the workbook: its names, and the formula's sheet (whose local names it can see). */
export interface LegacyContext {
  names?: NameInput[]
  sheet?: string
}

/**
 * The conversion of a book's formulas of the old kind, given the book's names: a formula's tree and its sheet give the
 * dynamic formula's tree, or undefined when it needs no `@` (the formula means the same in both kinds).
 */
export function legacyConverter(names: readonly NameInput[] = []): (ast: Ast, sheet?: string) => Ast | undefined {
  const parsed = new Map<NameInput, Ast | null>()
  const formulaOf = (n: NameInput): Ast | null => {
    if (!parsed.has(n))
      try {
        parsed.set(n, parse(n.ref, { union: true }))
      } catch {
        parsed.set(n, null)
      }
    return parsed.get(n)!
  }
  const lookup: Names = (name, sheet) => {
    const found = inputNameFor(names, sheet, name)
    const ast = found && formulaOf(found)
    return found && ast ? { ast } : undefined
  }
  return converter(lookup)
}

/**
 * A legacy formula (text without `=`, as a file stores it, with the prefixes of newer functions removed) as the dynamic
 * array formula Excel 365 displays for it, in canonical text: `A1:A10*2` becomes `@A1:A10*2`, `INDEX(A1:C3,0,1)`
 * becomes `@INDEX(A1:C3,0,1)`, `SUM(A1:A10)` stays. Excel puts `@` before a name depending on what the name's formula
 * returns, so the workbook's names are needed for formulas that use names. A formula the engine cannot parse is
 * returned as is. The API function (`strict-args.ts`), the same from `xlcalc/xlsx` and from the root.
 */
export const fromLegacy = /* @__PURE__ */ strict(function fromLegacy(formula: string, ctx?: LegacyContext): string {
  if (ctx !== undefined) onlyKeys(object(ctx, 'ctx', "the workbook's names and the formula's sheet, such as { names, sheet: 'Sales' }"), 'ctx', ['names', 'sheet'])
  let ast: Ast
  try {
    ast = parse(formula)
  } catch (e) {
    if (e instanceof FormulaError) return formula
    throw e
  }
  return canonical(legacyConverter(ctx?.names)(ast, ctx?.sheet) ?? ast)
})
