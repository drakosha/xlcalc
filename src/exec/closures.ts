/**
 * The `closures` mode: each formula becomes a tree of closures over the runtime, built by the same walk as
 * the `eval` text, without generating code from strings (so it runs under a Content Security Policy
 * without 'unsafe-eval'). Constants are built once per formula, as the `eval` factory hoists them.
 */
import type { Arg } from '../args.js'
import type { Runtime } from '../runtime.js'
import { Areas, type V } from '../values.js'
import { once, OPERATORS, type Backend, type Operator } from './walk.js'

type Fn = (parameters?: readonly V[]) => Arg

export function closureBackend(rt: Runtime, parameters?: V[]): Backend<Fn> {
  const { cell, spill, scalar, cmp, L1, L2, C, F, FL, LZ, ext, I, U, RO } = rt
  // the runtime's scalar operators get what the walk proved to be one value
  const op1 = (name: 'neg' | 'pct') => rt[name] as (a: Arg) => V
  const op2 = (name: Operator) => rt[name] as (a: Arg, b: Arg) => V
  return {
    value: (v) => {
      if (!parameters) return () => v
      const i = parameters.push(v) - 1
      return (p) => p![i]
    },
    // the argument the function takes as not given (its default)
    absent: () => () => undefined as unknown as Arg,
    array: (a) => () => a,
    cell: (s, r, c) => {
      if (!parameters) return () => cell(s, r, c)
      const i = parameters.length
      parameters.push(s, r, c)
      return (p) => cell(p![i] as number, p![i + 1] as number, p![i + 2] as number)
    },
    spill: (s, r, c) => {
      if (!parameters) return () => spill(s, r, c)
      const i = parameters.length
      parameters.push(s, r, c)
      return (p) => spill(p![i] as number, p![i + 1] as number, p![i + 2] as number)
    },
    range: (g, one) => (one ? () => scalar(g) : () => g),
    unary: (op, a, lift) => {
      const f = op1(op)
      return lift ? (p) => L1(f, a(p)) : (p) => f(a(p))
    },
    binary: (op, a, b, lift) => {
      const name = OPERATORS[op]
      const f = name ? op2(name) : undefined
      if (lift) {
        const g = f ?? C[op]
        return (p) => L2(g, a(p), b(p))
      }
      return f ? (p) => f(a(p), b(p)) : (p) => cmp(op, a(p) as V, b(p) as V)
    },
    call: (name, args, lift) => {
      const f = (lift ? FL : F)[name]
      return (p) => f(...args.map((a) => a(p)))
    },
    lazy: (name, args, check) => {
      const thunks = (p: readonly V[] | undefined) => p ? args.map((a) => () => a(p)) : args
      if (check.length) return (p) => LZ(name, check, ...thunks(p))
      const f = F[name]
      return (p) => f(...thunks(p))
    },
    external: (text) => () => ext(text),
    read: (a) => (p) => scalar(a(p)),
    areas: (list) => {
      const a = new Areas(list, true)
      return () => a
    },
    union: (items) => (p) => U(...items.map((a) => a(p))),
    refOp: (op, a, b) => (p) => RO(op, a(p), b(p)),
    at: (a, r, c) => {
      if (!parameters) return (p) => I(a(p), r, c)
      const i = parameters.length
      parameters.push(r, c)
      return (p) => I(a(p), p![i] as number, p![i + 1] as number)
    },
    let: (count, build) => {
      // the uses read the current computation's values
      const cells: Fn[] = new Array(count)
      const { values, body } = build((i) => () => cells[i]())
      return () => {
        for (let i = 0; i < count; i++) cells[i] = once(values[i])
        return body()
      }
    },
  }
}
