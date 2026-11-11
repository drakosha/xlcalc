/** Compile-only selection of CSE calls; ordinary formulas and forced array contexts keep the dynamic path. */
import type { FunctionSpec } from '../args.js'
import type { NameRef, Shape } from '../arrays/shape.js'
import { FUNCTIONS } from '../functions.js'
import type { Ast, BinOp } from '../syntax.js'

/** Scalar arguments of the two functions whose CSE evaluation is positional (measured against Excel). */
export const cseCall = (name: string) => name === 'XLOOKUP' || name === 'INDEX'

/** A multi-cell CSE enters array evaluation through value slots and array-producing whole slots. */
export const cseArray = (name: string, spec: FunctionSpec | undefined, i: number, multiple: boolean) => {
  if (!multiple || !spec) return false
  const mode = spec.mode(i)
  if (mode === 'lift' && cseCall(name)) return false
  return (!spec.refs && mode === 'lift' && spec.rangeAt(i) !== 'first') || (spec.array === true && (mode === 'whole' || mode === 'pass'))
}

/** Arithmetic enters the CSE array context; comparisons keep their positional operands. */
export const cseBinary = (op: BinOp, multiple: boolean) => multiple && '+-*/^&'.includes(op)

/** Whether a call actually needs the positional path, through the values of lazy LET bindings and defined names. */
export function needsCse(ast: Ast, shape: Map<Ast, Shape>, name: NameRef, scope: number, multiple: boolean): boolean {
  type Binding = { value: Ast; full: boolean }
  let locals: (Binding | undefined)[] = []
  const inside = new Set<Ast>()
  const go = (n: Ast, full: boolean, scope: number): boolean => {
    switch (n.k) {
      case 'fn': {
        const spec = FUNCTIONS[n.name]
        if (!full && spec && cseCall(n.name) && n.args.some((a, i) => spec.mode(i) === 'lift' && shape.get(a))) return true
        return n.args.some((a, i) => go(a, full || cseArray(n.name, spec, i, multiple), scope))
      }
      case 'bin': {
        const array = full || cseBinary(n.op, multiple)
        return go(n.a, array, scope) || go(n.b, array, scope)
      }
      case 'un':
      case 'pct':
        return go(n.a, full || (multiple && (n.k === 'pct' || n.op !== '+')), scope)
      case 'paren':
      case 'at':
        return go(n.a, full, scope)
      case 'union':
        return n.items.some((x) => go(x, full, scope))
      case 'local': {
        const b = locals[n.slot]
        if (!b || inside.has(b.value)) return false
        inside.add(b.value)
        const found = go(b.value, b.full, scope)
        inside.delete(b.value)
        return found
      }
      case 'let': {
        const outer = locals
        locals = locals.slice()
        for (const b of n.binds) locals[b.slot] = { value: b.value, full }
        const found = go(n.body ?? n.binds[n.binds.length - 1].value, full, scope)
        locals = outer
        return found
      }
      case 'lambda': {
        const outer = locals
        locals = locals.slice()
        for (const p of n.params) locals[p.slot] = undefined
        const found = go(n.body, true, scope)
        locals = outer
        return found
      }
      case 'call':
        return go(n.fn, full, scope) || n.args.some((a) => go(a, full, scope))
      case 'name': {
        const d = name(n, scope)
        if (!d || inside.has(d.ast)) return false
        inside.add(d.ast)
        const outer = locals
        locals = []
        const found = go(d.ast, full, -1)
        locals = outer
        inside.delete(d.ast)
        return found
      }
      default:
        return false
    }
  }
  return go(ast, false, scope)
}
