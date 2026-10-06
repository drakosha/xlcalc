/**
 * The `closures` mode: each formula becomes a tree of closures over the runtime, built by the same walk as
 * the `eval` text, without generating code from strings (so it runs under a Content Security Policy
 * without 'unsafe-eval'). Constants are built once per formula, as the `eval` factory hoists them.
 */
import type { Arg } from '../args.js'
import type { Runtime } from '../runtime.js'
import type { V } from '../values.js'
import { OPERATORS, type Backend, type Operator } from './walk.js'

type Fn = () => Arg

export function closureBackend(rt: Runtime): Backend<Fn> {
  const { cell, spill, scalar, cmp, L1, L2, C, F, FL, LZ, ext } = rt
  // the runtime's scalar operators get what the walk proved to be one value
  const op1 = (name: 'neg' | 'pct') => rt[name] as (a: Arg) => V
  const op2 = (name: Operator) => rt[name] as (a: Arg, b: Arg) => V
  return {
    value: (v) => () => v,
    array: (a) => () => a,
    cell: (s, r, c) => () => cell(s, r, c),
    spill: (s, r, c) => () => spill(s, r, c),
    range: (g, one) => (one ? () => scalar(g) : () => g),
    unary: (op, a, lift) => {
      const f = op1(op)
      return lift ? () => L1(f, a()) : () => f(a())
    },
    binary: (op, a, b, lift) => {
      const name = OPERATORS[op]
      const f = name ? op2(name) : undefined
      if (lift) {
        const g = f ?? C[op]
        return () => L2(g, a(), b())
      }
      return f ? () => f(a(), b()) : () => cmp(op, a() as V, b() as V)
    },
    call: (name, args, lift) => {
      const f = (lift ? FL : F)[name]
      return () => f(...args.map((a) => a()))
    },
    lazy: (name, args, check) => {
      if (check.length) return () => LZ(name, check, ...args)
      const f = F[name]
      return () => f(...args)
    },
    external: (text) => () => ext(text),
    read: (a) => () => scalar(a()),
  }
}
