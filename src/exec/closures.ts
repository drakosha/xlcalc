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

/** A leaf the backend made last: a cell or a constant, which an operator over two of them reads itself. */
interface Leaf {
  fn: Fn | undefined
  cell: boolean
  s: number
  r: number
  c: number
  v: V
}

export function closureBackend(rt: Runtime, parameters?: Arg[]): Backend<Fn> {
  const { cell, spill, scalar, cmp, L1, L2, C, F, FL, LZ, ext, I, U, RO } = rt
  // the runtime's scalar operators get what the walk proved to be one value
  const op1 = (name: 'neg' | 'pct') => rt[name] as (a: Arg) => V
  const op2 = (name: Operator) => rt[name] as (a: Arg, b: Arg) => V
  // the last two leaves (records reused, nothing made for them): an operator over two of them is one closure, not three
  const leaves: Leaf[] = [0, 1].map(() => ({ fn: undefined, cell: false, s: 0, r: 0, c: 0, v: null }))
  const leaf = (fn: Fn, cell: boolean, s: number, r: number, c: number, v: V): Fn => {
    const x = leaves[0]
    leaves[0] = leaves[1]
    leaves[1] = x
    x.fn = fn
    x.cell = cell
    x.s = s
    x.r = r
    x.c = c
    x.v = v
    return fn
  }
  const leafOf = (fn: Fn) => (leaves[0].fn === fn ? leaves[0] : leaves[1].fn === fn ? leaves[1] : undefined)
  /** `f` over two leaves, each read where the closure runs, in the order of the operands. */
  const over = (f: (a: Arg, b: Arg) => V, x: Leaf, y: Leaf): Fn => {
    const { s, r, c, v } = x
    const [t, u, w, z] = [y.s, y.r, y.c, y.v]
    if (x.cell && y.cell) return () => f(cell(s, r, c), cell(t, u, w))
    if (x.cell) return () => f(cell(s, r, c), z)
    return y.cell ? () => f(v, cell(t, u, w)) : () => f(v, z)
  }
  return {
    value: (v) => {
      if (!parameters) return leaf(() => v, false, 0, 0, 0, v)
      const i = parameters.push(v) - 1
      return (p) => p![i]
    },
    // the argument the function takes as not given (its default)
    absent: () => () => undefined as unknown as Arg,
    array: (a) => () => a,
    cell: (s, r, c) => {
      if (!parameters) return leaf(() => cell(s, r, c), true, s, r, c, null)
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
      if (!f) return (p) => cmp(op, a(p) as V, b(p) as V)
      const x = leafOf(a), y = x && leafOf(b)
      return x && y && x !== y ? over(f, x, y) : (p) => f(a(p), b(p))
    },
    call: (name, args, lift) => {
      const f = (lift ? FL : F)[name]
      // a call of one, two or three arguments makes no array of them each time
      const [a, b, c] = args
      if (args.length === 1) return (p) => f(a(p))
      if (args.length === 2) return (p) => f(a(p), b(p))
      if (args.length === 3) return (p) => f(a(p), b(p), c(p))
      return (p) => f(...args.map((x) => x(p)))
    },
    lazy: (name, args, check) => {
      const thunks = (p: readonly Arg[] | undefined) => p ? args.map((a) => () => a(p)) : args
      if (check.length) return (p) => LZ(name, check, ...thunks(p))
      const f = F[name]
      return (p) => f(...thunks(p))
    },
    external: (text) => () => ext(text),
    moving: (at, read) => {
      if (read === 'range') return at
      if (read === 'value')
        return () => {
          const g = at()
          return g.r1 === g.r2 && g.c1 === g.c2 ? scalar(g) : g
        }
      const f = read === 'cell' ? cell : spill
      return () => {
        const g = at()
        return f(g.s, g.r1, g.c1)
      }
    },
    computed: (f) => f,
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
