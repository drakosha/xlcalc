/**
 * The `closures` mode: each formula becomes a tree of closures over the runtime, built by the same walk as
 * the `eval` text, without generating code from strings (so it runs under a Content Security Policy
 * without 'unsafe-eval'). Constants are built once per formula, as the `eval` factory hoists them.
 */
import type { Arg } from '../args.js'
import type { Runtime } from '../runtime.js'
import type { V } from '../values.js'
import type { Backend } from './walk.js'

type Fn = () => Arg
type Op = (a: Arg, b: Arg) => Arg

export function closureBackend(rt: Runtime): Backend<Fn> {
  const { cell, scalar, cmp, L1, L2, C, F, FL, LZ, ext } = rt
  // the runtime's scalar operators take what the walk proved to be one value
  const OPS = { '+': rt.add, '-': rt.sub, '*': rt.mul, '/': rt.div, '^': rt.pow, '&': rt.cat } as Record<string, Op>
  const UN = { neg: rt.neg, pct: rt.pct } as Record<string, (a: Arg) => Arg>
  return {
    value: (v) => () => v,
    array: (a) => () => a,
    cell: (s, r, c) => () => cell(s, r, c),
    range: (g, one) => (one ? () => scalar(g) : () => g),
    unary: (op, a, lift) => {
      const f = UN[op]
      return lift ? () => L1(f as (a: V) => V, a()) : () => f(a())
    },
    binary: (op, a, b, lift) => {
      const f = OPS[op]
      if (lift) return () => L2((f ?? C[op]) as (a: V, b: V) => V, a(), b())
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
  }
}
