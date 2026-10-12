/**
 * The `closures` mode: each formula becomes a tree of closures over the runtime, built by the same walk as
 * the `eval` text, without generating code from strings (so it runs under a Content Security Policy
 * without 'unsafe-eval'). Constants are built once per formula, as the `eval` factory hoists them.
 */
import type { Arg } from '../args.js'
import type { Runtime } from '../runtime.js'
import { Areas, type V } from '../values.js'
import { once, operator, type Backend, type Operator } from './walk.js'

type Fn = (parameters?: readonly Arg[]) => Arg

export function closureBackend(rt: Runtime, parameters?: Arg[]): Backend<Fn> {
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
      parameters.push(r)
      return (p) => cell(s, p![i] as number, c)
    },
    spill: (s, r, c) => {
      if (!parameters) return () => spill(s, r, c)
      const i = parameters.length
      parameters.push(r)
      return (p) => spill(s, p![i] as number, c)
    },
    range: (g, one) => {
      if (!parameters) return one ? () => scalar(g) : () => g
      const i = parameters.push(g) - 1
      return one ? (p) => scalar(p![i]) : (p) => p![i]
    },
    unary: (op, a, lift) => {
      const f = op1(op)
      return lift ? (p) => L1(f, a(p)) : (p) => f(a(p))
    },
    binary: (op, a, b, lift, last) => {
      const name = operator(op, last)
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
      const thunks = (p: readonly Arg[] | undefined) => p ? args.map((a) => () => a(p)) : args
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
      const frames = parameters ? new WeakMap<readonly Arg[], Fn[]>() : undefined
      const current = (p: readonly Arg[] | undefined) => p && frames ? frames.get(p)! : cells
      const { values, body } = build((i) => (p) => current(p)[i]())
      return (p) => {
        const frame = frames && p ? new Array<Fn>(count) : cells
        if (frames && p) frames.set(p, frame)
        for (let i = 0; i < count; i++) frame[i] = once(() => values[i](p))
        return body(p)
      }
    },
  }
}
