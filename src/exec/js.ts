/**
 * The `eval` mode: each formula becomes a self-contained factory expression in JS, and a batch of formulas
 * is compiled by one `new Function`. Only the package generates code: literals go through `JSON.stringify`,
 * only whitelisted runtime functions are called, sheets and ranges are passed as numbers.
 */
import type { Arg } from '../args.js'
import { Arr } from '../arrays/value.js'
import { PACKAGE_NAME } from '../config.js'
import type { Runtime } from '../runtime.js'
import { Areas, isErr, Rng, type V } from '../values.js'
import { once, OPERATORS, type Backend } from './walk.js'

const lit = (v: V) => (isErr(v) ? `E.${v.code}` : JSON.stringify(v))

/** A backend for one formula: constants are hoisted into its factory, `source` wraps the root expression. */
export function jsBackend(): Backend<string> & { source(root: string): string } {
  const hoisted: string[] = []
  const hoist = (code: string) => `h${hoisted.push(code) - 1}`
  // LET's names of the formula: v0, v1, …
  let names = 0
  return {
    value: lit,
    absent: () => 'undefined',
    array: (a) => hoist(`A(${a.rows},${a.cols},[${a.v.map(lit).join(',')}])`),
    cell: (s, r, c) => `cell(${s},${r},${c})`,
    spill: (s, r, c) => `spill(${s},${r},${c})`,
    range: (g, one) => {
      const h = hoist(`R(${g.s},${g.r1},${g.c1},${g.r2},${g.c2})`)
      return one ? `scalar(${h})` : h
    },
    unary: (op, a, lift) => (lift ? `L1(${op},${a})` : `${op}(${a})`),
    binary: (op, a, b, lift) => {
      const f = OPERATORS[op]
      if (lift) return `L2(${f ?? `C[${JSON.stringify(op)}]`},${a},${b})`
      return f ? `${f}(${a},${b})` : `cmp(${JSON.stringify(op)},${a},${b})`
    },
    call: (name, args, lift) => `${lift ? 'FL' : 'F'}${key(name)}(${args.join(',')})`,
    lazy: (name, args, check) => {
      const thunks = args.map((a) => `()=>${a}`)
      return check.length ? `LZ(${JSON.stringify(name)},[${check.join(',')}],${thunks.join(',')})` : `F${key(name)}(${thunks.join(',')})`
    },
    external: (text) => `ext(${JSON.stringify(text)})`,
    read: (a) => `scalar(${a})`,
    areas: (list) => hoist(`X([${list.map((g) => `R(${g.s},${g.r1},${g.c1},${g.r2},${g.c2})`).join(',')}])`),
    union: (items) => `U(${items.join(',')})`,
    refOp: (op, a, b) => `RO(${JSON.stringify(op)},${a},${b})`,
    at: (a, r, c) => `I(${a},${r},${c})`,
    let: (count, build) => {
      const first = names
      names += count
      const { values, body } = build((i) => `v${first + i}()`)
      return `(()=>{const ${values.map((v, i) => `v${first + i}=M(()=>${v})`).join(',')};return ${body}})()`
    },
    source: (root) => `(()=>{${hoisted.length ? `const ${hoisted.map((h, i) => `h${i}=${h}`).join(',')};` : ''}return ()=>${root}})()`,
  }
}

const key = (name: string) => (/^[A-Z][A-Z0-9_]*$/.test(name) ? `.${name}` : `[${JSON.stringify(name)}]`)

export const PRELUDE = '"use strict";const {cell,spill,scalar,add,sub,mul,div,pow,neg,pct,cat,cmp,L1,L2,C,F,FL,LZ,ext,I,U,RO,E,R,A,X,M}=rt;'

/** The functions of a batch of formula sources, by one `new Function`. */
export function compileSources(rt: Runtime, sources: string[]): (() => Arg)[] {
  const R = (s: number, r1: number, c1: number, r2: number, c2: number) => new Rng(s, r1, c1, r2, c2)
  const A = (rows: number, cols: number, v: V[]) => new Arr(rows, cols, rows, cols, v)
  const X = (list: Rng[]) => new Areas(list, true)
  let make: (rt: unknown) => unknown
  try {
    make = new Function('rt', `${PRELUDE}return [${sources.join(',\n')}]`) as typeof make
  } catch (e) {
    // a Content Security Policy without 'unsafe-eval' blocks new Function with an EvalError
    if (e instanceof EvalError) throw new EvalError(`${PACKAGE_NAME} compiles formulas with new Function for execution 'eval', and this environment blocks it (a Content Security Policy without 'unsafe-eval'?); execution 'closures' does not need it`, { cause: e })
    throw e
  }
  return make({ ...rt, R, A, X, M: once }) as (() => Arg)[]
}
