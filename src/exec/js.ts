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
import { cseAt, cseMap } from './cse.js'
import { callValue, lambdaOf, OMITTED, parameters as given } from './lambda.js'
import { once, operator, type Backend } from './walk.js'

const lit = (v: V) => (isErr(v) ? `E.${v.code}` : JSON.stringify(v))

/** A backend for one formula: constants are hoisted into its factory, `source` wraps the root expression. */
export function jsBackend(parameters?: Arg[]): Backend<string> & { source(root: string): string } {
  const hoisted: string[] = []
  const hoist = (code: string) => `h${hoisted.push(code) - 1}`
  // LET's names and LAMBDA's parameters of the formula: v0, v1, …; the definitions that call themselves: r0, r1, …
  let names = 0
  // the variables of `count` new names: v<i> for the i-th
  const variables = (count: number) => {
    const first = names
    names += count
    return (i: number) => `v${first + i}`
  }
  let selves = 0
  const bound = (v: Arg) => parameters ? `p[${parameters.push(v) - 1}]` : lit(v as V)
  return {
    value: bound,
    absent: () => 'undefined',
    array: (a) => hoist(`A(${a.rows},${a.cols},[${a.v.map(lit).join(',')}])`),
    cell: (s, r, c) => `cell(${s},${bound(r)},${c})`,
    spill: (s, r, c) => `spill(${s},${bound(r)},${c})`,
    range: (g, one) => {
      const h = parameters ? bound(g) : hoist(`R(${g.s},${g.r1},${g.c1},${g.r2},${g.c2})`)
      return one ? `scalar(${h})` : h
    },
    unary: (op, a, lift) => (lift ? `L1(${op},${a})` : `${op}(${a})`),
    binary: (op, a, b, lift, last) => {
      const f = operator(op, last)
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
    at: (a, r, c) => `I(${a},${bound(r)},${bound(c)})`,
    cse: (rows, cols, build) => {
      const cursor = variables(2)
      const kept: string[] = []
      const body = build(
        (value) => `CS(${value},${cursor(0)},${cursor(1)},cell)`,
        (value, local) => {
          const v = variables(1)(0)
          kept.push(local ? v : `${v}=M(()=>${value})`)
          return local ? `(${v}??=M(()=>${value}))()` : `${v}()`
        },
      )
      const code = `CM(${bound(rows)},${bound(cols)},(${cursor(0)},${cursor(1)})=>${body},cell)`
      return kept.length ? `(()=>{let ${kept.join(',')};return ${code}})()` : code
    },
    let: (count, build) => {
      const v = variables(count)
      const { values, body } = build((i) => `${v(i)}()`)
      return `(()=>{const ${values.map((x, i) => `${v(i)}=M(()=>${x})`).join(',')};return ${body}})()`
    },
    lambda: (count, required, build) => {
      const v = variables(count)
      const { body, height } = build((i) => `${v(i)}()`, (i) => `(${v(i)}===OM)`)
      return `LAM(${count},${height},(${Array.from({ length: count }, (_, i) => v(i)).join(',')})=>${body}${required < count ? `,${required}` : ''})`
    },
    invoke: (fn, args) => `CALL(${fn},[${args.join(',')}])`,
    recursive: (build) => {
      const self = `r${selves++}`
      return `(()=>{const ${self}=()=>${build(`${self}()`)};return ${self}()})()`
    },
    eta: (name) => `ETA(${JSON.stringify(name)})`,
    source: (root) => `${parameters ? '(p)=>' : '(()=>'}{${hoisted.length ? `const ${hoisted.map((h, i) => `h${i}=${h}`).join(',')};` : ''}return ()=>${root}}${parameters ? '' : ')()'}`,
  }
}

const key = (name: string) => (/^[A-Z][A-Z0-9_]*$/.test(name) ? `.${name}` : `[${JSON.stringify(name)}]`)

export const PRELUDE = '"use strict";const {cell,spill,scalar,add,sub,addLast,subLast,mul,div,pow,neg,pct,cat,cmp,L1,L2,C,F,FL,LZ,ext,I,U,RO,E,R,A,X,M,LAM,CALL,ETA,OM}=rt;'

/** The functions of a batch of formula sources, by one `new Function`. */
export function compileSources<T = () => Arg>(rt: Runtime, sources: string[]): T[] {
  const R = (s: number, r1: number, c1: number, r2: number, c2: number) => new Rng(s, r1, c1, r2, c2)
  const A = (rows: number, cols: number, v: V[]) => new Arr(rows, cols, rows, cols, v)
  const X = (list: Rng[]) => new Areas(list, true)
  let make: (rt: unknown, CS: typeof cseAt, CM: typeof cseMap) => unknown
  try {
    make = new Function('rt', 'CS', 'CM', `${PRELUDE}return [${sources.join(',\n')}]`) as typeof make
  } catch (e) {
    // a Content Security Policy without 'unsafe-eval' blocks new Function with an EvalError
    if (e instanceof EvalError) throw new EvalError(`${PACKAGE_NAME} compiles formulas with new Function for execution 'eval', and this environment blocks it (a Content Security Policy without 'unsafe-eval'?); execution 'closures' does not need it`, { cause: e })
    throw e
  }
  // a LAMBDA's calculation gets its parameters as the uses of LET's names read them
  const LAM = (count: number, height: number, f: (...vals: (() => Arg)[]) => Arg, required?: number) => lambdaOf(count, height, (args) => f(...given(args, count)), required)
  return make({ ...rt, R, A, X, M: once, LAM, CALL: callValue, OM: OMITTED }, cseAt, cseMap) as T[]
}
