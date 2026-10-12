/**
 * Static shapes of a formula's nodes (`docs/arrays-incell.md`): 0 one value, 1 maybe an array (it depends on
 * the data: INDEX, IF, functions with an array answer), 2 an array of more than one value whatever the data.
 * The code generator lifts only nodes above 0, so a formula without arrays compiles as before; a formula of
 * shape 2 has a result that cannot be shown in one cell.
 */
import type { FunctionSpec } from '../args.js'
import type { NameDef } from '../exec/walk.js'
import { FUNCTIONS } from '../functions.js'
import type { ErrorCode } from '../contract.js'
import type { Rect } from '../structure/area.js'
import { isRefOp, type Ast } from '../syntax.js'

/** A name's definition as the names of sheet `scope` show it (`exec/walk.ts`). */
export type NameRef = (n: Extract<Ast, { k: 'name' }>, scope: number) => NameDef | undefined

/** A structured reference's rectangle in the formula's row, or its error. */
export type TableRef = (n: Extract<Ast, { k: 'table' }>) => Rect | ErrorCode

export type Shape = 0 | 1 | 2

interface Info {
  s: Shape
  /** The node is a reference of several cells (a range lifted or refused as such, not an array). */
  ref: boolean
}

const max = (a: Shape, b: Shape): Shape => (a > b ? a : b)

/** Several values that one of is returned (IF's branches, CHOOSE's values): 2 if all are arrays, else the most of 1. */
function either(xs: Shape[]): Shape {
  if (xs.length && xs.every((x) => x === 2)) return 2
  return xs.some((x) => x > 0) ? 1 : 0
}

/** A call: an array among the lifted arguments makes an array (of more than one value if one has more); else one of the passed ones, or maybe an array answer. */
function combine(lifted: Shape[], passed: Shape[], array: boolean): Shape {
  const l = lifted.reduce(max, 0)
  if (l === 2) return 2
  return max(max(l, either(passed)), array ? 1 : 0)
}

const isArea = (r: { kind: string; r1: number; r2: number; c1: number; c2: number }) => r.kind !== 'cell' && !(r.r1 === r.r2 && r.c1 === r.c2)

/**
 * Shapes of every node of a formula on sheet `scope`; a name is judged by its formula (`nameRef`), which sees the
 * global names, a LET name by its value; `inArea`: the formula is an old-kind array formula.
 */
export function shapes(ast: Ast, scope: number, nameRef: NameRef, inArea = false, tableRef: TableRef = () => 'REF'): Map<Ast, Shape> {
  const out = new Map<Ast, Shape>()
  // LET's names by slot (a name's formula has its own), the names being expanded
  let bound: Info[] = []
  const inside = new Set<NameDef>()
  const go = (n: Ast): Info => {
    const i = info(n)
    out.set(n, i.s)
    return i
  }
  const info = (n: Ast): Info => {
    switch (n.k) {
      case 'num':
      case 'str':
      case 'bool':
      case 'err':
      case 'missing':
        return { s: 0, ref: false }
      case 'arr':
        // a 1 × 1 constant is still an array (TYPE({1}) is 64), but one value wherever it goes
        return { s: n.rows.length * n.rows[0].length > 1 ? 2 : 1, ref: false }
      case 'ref':
        // a spill reference has as many cells as its array spilled into
        return n.ref.spill ? { s: 1, ref: true } : isArea(n.ref) ? { s: 2, ref: true } : { s: 0, ref: false }
      case 'xref':
        return isArea(n.ref) ? { s: 2, ref: true } : { s: 0, ref: false }
      // the engine puts another workbook's name's reference in its place before it computes (`links.ts`)
      case 'xname':
        return { s: 0, ref: false }
      case 'table': {
        const x = tableRef(n)
        return typeof x !== 'string' && (x.r2 > x.r1 || x.c2 > x.c1) ? { s: 2, ref: true } : { s: 0, ref: false }
      }
      // several areas are one value nowhere (#VALUE!), and a reference where a function takes them
      case 'ref3d':
        return { s: 0, ref: true }
      case 'union':
        n.items.forEach(go)
        return { s: 0, ref: true }
      // one cell or one value
      case 'at':
        go(n.a)
        return { s: 0, ref: false }
      case 'name': {
        const d = nameRef(n, scope)
        if (!d || inside.has(d)) return { s: 0, ref: false }
        const outer: [Info[], number] = [bound, scope]
        ;[bound, scope] = [[], -1]
        inside.add(d)
        const i = go(d.ast)
        inside.delete(d)
        ;[bound, scope] = outer
        return i
      }
      case 'paren':
        return go(n.a)
      case 'un': {
        const a = go(n.a)
        // unary plus keeps a reference: SUM(+A1:A3) sums the range
        return n.op === '+' ? a : { s: a.s, ref: false }
      }
      case 'pct':
        return { s: go(n.a).s, ref: false }
      case 'bin':
        // the intersection and the range operator answer a reference of one cell or more
        if (isRefOp(n.op)) {
          go(n.a)
          go(n.b)
          return { s: 1, ref: true }
        }
        return { s: max(go(n.a).s, go(n.b).s), ref: false }
      case 'fn':
        return call(n.name, n.args.map(go), n.args)
      case 'let':
        for (const b of n.binds) bound[b.slot] = go(b.value)
        return n.body ? go(n.body) : bound[n.binds[n.binds.length - 1].slot]
      case 'local':
        return bound[n.slot]
    }
  }
  const call = (name: string, args: Info[], nodes: Ast[]): Info => {
    const spec: FunctionSpec | undefined = FUNCTIONS[name]
    if (!spec) return { s: 0, ref: false }
    if (name === 'ROW' || name === 'COLUMN') {
      const a = nodes[0]
      const r = (a?.k === 'ref' && !a.ref.spill) || a?.k === 'xref' ? a.ref : null
      const many = r ? (name === 'ROW' ? r.r2 > r.r1 : r.c2 > r.c1) : args.length > 0 && args[0].s > 0
      // in an old-kind array formula ROW() and COLUMN() are its area's rows and columns
      return { s: r ? (many ? 2 : 0) : args.length || inArea ? 1 : 0, ref: false }
    }
    if (spec.lazy) {
      const lifted: Shape[] = []
      const rest: Shape[] = []
      args.forEach((a, i) => (spec.liftAt(i, args.length) ? lifted : rest).push(a.s))
      // IF without its third argument answers FALSE
      if (name === 'IF' && args.length === 2) rest.push(0)
      return { s: combine(lifted, rest, false), ref: false }
    }
    const lifted: Shape[] = []
    const passed: Shape[] = []
    args.forEach((a, i) => {
      const mode = spec.mode(i)
      if (mode === 'whole') return
      if (mode === 'pass') passed.push(a.s)
      else if (!a.ref || (mode === 'lift' && spec.rangeAt(i) === 'lift')) lifted.push(a.s)
    })
    return { s: combine(lifted, passed, spec.array === true), ref: false }
  }
  go(ast)
  return out
}
