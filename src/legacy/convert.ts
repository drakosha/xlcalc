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
 *   several cells; a lookup's key or place is such an argument wherever the lookup is;
 * - `R` inside an argument that takes a reference (SUM's): a range where one value goes gets `@`, an array stays;
 * - `A` inside an argument that takes an array: nothing gets `@`.
 */
import type { NameInput } from '../contract.js'
import { canonical, FormulaError, isRefOp, parse, type Ast, type RangeRef } from '../syntax.js'
import { inputNameFor } from '../names.js'
import { object, onlyKeys } from '../input/basic.js'
import { strict } from '../strict-args.js'
import { dereferenced } from '../syntax/idioms.js'
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
  /** A place Excel knows leaves one cell as INDEX's row or column: a whole number literal from 1, such a name, MATCH. */
  pos?: true
  /** A name of a formula that is not a reference: a reference where it goes, but INDEX takes it as an array. */
  named?: true
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

/** Functions whose arguments of class `V` take one value wherever the call is (`bench/coverage/legacy/array-held.ts`). */
const LOOKUPS = new Set(['INDEX', 'VLOOKUP', 'HLOOKUP', 'XLOOKUP'])

/** The slot and the evaluation of argument `i` of `fn` (the node `n`) called where `where` says. */
function argument(fn: string, i: number, n: Ast, where: Where): [Slot, Where] {
  const own = SLOTS[fn]?.[i]
  // under an argument that takes an array IFERROR's value takes a range as a reference, TYPE's whole (SUMPRODUCT(TYPE(A1:A3)) is 64)
  if (own) return [own, own === 'p' && where === 'A' && fn !== 'TYPE' ? 'R' : where]
  switch (classAt(fn, i)) {
    case 'v':
      return ['v', where]
    case 'w':
      // VLOOKUP's column: one value inside SUM (`SUM(VLOOKUP(4,…,@{2,3},FALSE))`), an array inside SUMPRODUCT
      return ['v', where === 'A' ? (n.k === 'ref' ? 'A' : 'R') : 'V']
    case 'V':
      // a lookup's key or place is one value even inside SUM: `SUM(XLOOKUP(@{1;2},…))`, but `N(IF(…))` there keeps every
      // value (`SUM(INDEX(B1:B3,N(IF(1,{2;3}))))` is unchanged, `N({2;3})` takes `@`); IF, CHOOSE, TRANSPOSE keep the array
      return ['v', LOOKUPS.has(fn) && !dereferenced(n) ? 'V' : where === 'A' ? 'R' : where]
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

/** A whole number literal from 1 to 65535 (Excel's two-byte integer, 65536 is not one): INDEX's place that leaves one cell. */
const counting = (n: Ast) => n.k === 'num' && Number.isInteger(n.v) && n.v >= 1 && n.v <= 65535

/** `pos` of an `Out` where `yes`. */
const placed = (yes: boolean | undefined): Pick<Out, 'pos'> => (yes ? { pos: true } : {})

/** The size of a literal reference: rows and columns, or null for anything else. */
function size(n: Ast): [number, number] | null {
  if (n.k === 'paren') return size(n.a)
  return (n.k === 'ref' && !n.ref.spill) || n.k === 'xref' ? [n.ref.r2 - n.ref.r1 + 1, n.ref.c2 - n.ref.c1 + 1] : null
}

/**
 * What the conversion reads of a formula's references and numbers outside array constants (`syntax/shape.ts`'s holes):
 * a reference's kind, size and spill mark, whether a number is 1 (OFFSET's size) or else a place (`counting`). Formulas
 * of one shape (its functions and names spelled in it) with one key convert alike, so a writer asks once per key
 * (`xlsx/write/kind.ts`); what else the conversion comes to read goes here.
 */
export const legacyKey = (holes: readonly (RangeRef | number)[]) =>
  holes.map((h) => (typeof h === 'number' ? (h === 1 ? '1' : counting({ k: 'num', v: h }) ? 'p' : 'n') : `${h.kind}${h.r2 - h.r1},${h.c2 - h.c1}${h.spill ? '#' : ''}`)).join(';')

/** The arguments `answer` reads of a call by place. A call with fewer (Excel refuses such a file, other writers make them) is one value. */
const READ_BY_PLACE: Readonly<Record<string, number>> = { INDEX: 1, OFFSET: 1, VLOOKUP: 3, HLOOKUP: 3, XLOOKUP: 3 }

/** A call's answer after its arguments are converted. */
function answer(fn: string, args: Out[], nodes: Ast[]): Kind {
  if (args.length < (READ_BY_PLACE[fn] ?? 0)) return 'one'
  switch (fn) {
    case 'ROW':
    case 'COLUMN':
      return args.length && args[0].kind !== 'one' ? 'rowcol' : 'one'
    case 'INDEX': {
      const ref = (args[0].kind === 'range' && !args[0].named) || nodes[0].k === 'ref' || nodes[0].k === 'xref'
      // of one cell or value one value (`INDEX(B1,0)` too); of several, one cell when every index given is a place that leaves one
      const one = args[0].kind === 'one' || (args[1]?.pos && (nodes.length < 3 ? ref : args[2].pos))
      return one ? 'one' : ref ? 'range' : 'array'
    }
    case 'OFFSET': {
      // one cell when the height and the width are each the literal 1, or left out from a cell (`B1:B1` is a range); the shifts do not matter
      const one = (n: Ast | undefined) => (n === undefined ? args[0].kind === 'one' : n.k === 'num' && n.v === 1)
      return one(nodes[3]) && one(nodes[4]) ? 'one' : 'range'
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

/** An operator's node over its converted operands. */
const joined = (n: Extract<Ast, { k: 'bin' }>, a: Out, b: Out): Out => ({ ast: { ...n, a: a.ast, b: b.ast }, kind: a.kind === 'one' && b.kind === 'one' ? 'one' : 'array' })

function converter(names: Names) {
  // the kind of each LET name, by slot; the number of `@` put in
  const locals: Kind[] = []
  let added = 0
  const go = (n: Ast, where: Where, slot: Slot, scope: string | undefined): Out => marked(node(n, where, slot, scope), where, slot)
  const marked = (o: Out, where: Where, slot: Slot): Out => {
    if (!intersects(o.kind, slot, where)) return o
    added++
    return { ast: { k: 'at', a: o.ast }, kind: 'one' }
  }
  const node = (n: Ast, where: Where, slot: Slot, scope: string | undefined): Out => {
    switch (n.k) {
      case 'ref':
      case 'xref':
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
        return { ast: n, ...nameKind(n, scope) }
      case 'num':
        return { ast: n, kind: 'one', ...placed(counting(n)) }
      case 'un':
      case 'pct': {
        const a = go(n.a, where, 'v', scope)
        // Excel reads `+2` as the literal 2 (`+MATCH(…)` is not MATCH)
        return { ast: { ...n, a: a.ast }, kind: a.kind === 'one' ? 'one' : 'array', ...placed(n.k === 'un' && n.op === '+' && counting(n.a)) }
      }
      case 'bin': {
        // the reference operators take references whole and answer one
        if (isRefOp(n.op)) {
          const [a, b] = [go(n.a, 'R', 'r', scope), go(n.b, 'R', 'r', scope)]
          // an intersection with one cell is one cell at most
          return { ast: { ...n, a: a.ast, b: b.ast }, kind: n.op === ' ' && (a.kind === 'one' || b.kind === 'one') ? 'one' : 'range' }
        }
        // `1+1+1` nests to the left by a node per operator: in a loop, for a chain of thousands
        const chain = [n]
        for (let l = n.a; l.k === 'bin' && !isRefOp(l.op); l = l.a) chain.push(l)
        let a = go(chain[chain.length - 1].a, where, 'v', scope)
        for (let i = chain.length - 1; i > 0; i--) a = marked(joined(chain[i], a, go(chain[i].b, where, 'v', scope)), where, 'v')
        return joined(n, a, go(n.b, where, 'v', scope))
      }
      // `(@A1:A3)`: the operand in its place gets `@`
      case 'paren': {
        const a = go(n.a, where, slot, scope)
        return { ast: { ...n, a: a.ast }, kind: a.kind, ...placed(a.pos) }
      }
      case 'fn': {
        const args = n.args.map((a, i) => {
          const [s, w] = argument(n.name, i, a, where)
          return go(a, w, s, scope)
        })
        // Excel knows MATCH's answer for a place, and no other call's (XMATCH, ROWS, COUNT are not)
        return { ast: { ...n, args: args.map((a) => a.ast) }, kind: answer(n.name, args, n.args), ...placed(n.name === 'MATCH') }
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
  // a name: Excel takes a literal or a cell for one value (a whole number from 1 for INDEX's place, as the literal), any
  // other formula for a reference whatever it computes (`1+1`, `(2)`, ABS(-2), an array, another name get `@`, not in N)
  const nameKind = (n: Extract<Ast, { k: 'name' }>, scope: string | undefined): Omit<Out, 'ast'> => {
    const def = names(n.name, n.sheet ?? scope)?.ast
    if (!def) return { kind: 'one' }
    const lit = def.k === 'un' && def.op === '-' ? def.a : def
    if (lit.k === 'num' || lit.k === 'str' || lit.k === 'bool' || lit.k === 'err') return { kind: 'one', ...placed(counting(def)) }
    if (def.k === 'ref') return { kind: def.ref.kind === 'cell' && !def.ref.spill ? 'one' : 'range' }
    // a range of the reference operators (`'S'!B1:'S'!B8`) is a formula too
    return { kind: 'range', ...(def.k === 'ref3d' ? {} : { named: true }) }
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
 * returned as is. The API function (`strict-args.ts`), from `xlcalc/xlsx`.
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
