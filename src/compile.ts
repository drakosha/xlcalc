/**
 * Compilation of a formula to JS: a self-contained factory expression per formula, and one `new Function`
 * per batch. Only the package generates code: literals go through `JSON.stringify`, only whitelisted runtime
 * functions are called, sheets and ranges are passed as numbers.
 */
import type { Arg } from './args.js'
import { shapes } from './arrays/shape.js'
import { Arr } from './arrays/value.js'
import { PACKAGE_NAME } from './config.js'
import { FUNCTIONS } from './functions.js'
import type { Runtime } from './runtime.js'
import { canonical, type Ast } from './syntax.js'
import { Rng, type V } from './values.js'

/** What the code generator needs to know of the book, seen from the formula's sheet. */
export interface CodegenContext {
  /** The index of a sheet (null: the formula's own). */
  sheet(name: string | null): number
  /** A name visible from the formula's sheet: its reference, and whether it has several cells. */
  name(name: string): { ast: Extract<Ast, { k: 'ref' }>; area: boolean } | undefined
}

const OPS: Record<string, string> = { '+': 'add', '-': 'sub', '*': 'mul', '/': 'div', '^': 'pow', '&': 'cat' }

/**
 * JS code of one formula (a self-contained factory expression) and its references. Nodes that can be arrays
 * (`arrays/shape.ts`) go through the lifting operators (`L1`, `L2`) and callers (`FL`, `LZ`); the rest is
 * plain scalar code. A range of several cells in a one-value position is passed as a reference to be lifted.
 */
export function codegen(ctx: CodegenContext, ast: Ast, s: number, r: number, c: number): { source: string; refs: Rng[]; volatile: boolean; nested: boolean } {
  const refs: Rng[] = []
  const hoisted: string[] = []
  let volatile = false
  let nested = false
  const sheetOf = ctx.sheet
  const deref = (n: Ast): Ast => (n.k === 'name' ? (ctx.name(n.name)?.ast ?? n) : n)
  const shape = shapes(ast, (name) => ctx.name(name)?.area)
  const many = (n: Ast) => (shape.get(n) ?? 0) > 0
  const hoist = (code: string) => `h${hoisted.push(code) - 1}`
  // rangeOk: a position where the function expects a range (a range object goes there, even for a single cell)
  const go = (n: Ast, rangeOk = false): string => {
    switch (n.k) {
      case 'num':
      case 'str':
      case 'bool':
        return JSON.stringify(n.v)
      case 'err':
        return `E.${n.code}`
      case 'arr': {
        const v = n.rows.flat().map((x) => (x.k === 'err' ? `E.${x.code}` : JSON.stringify(x.v)))
        return hoist(`A(${n.rows.length},${n.rows[0].length},[${v.join(',')}])`)
      }
      case 'ref': {
        const g = new Rng(sheetOf(n.sheet), n.ref.r1, n.ref.c1, n.ref.r2, n.ref.c2)
        refs.push(g)
        if (n.ref.kind === 'cell' && !rangeOk) return `cell(${g.s},${g.r1},${g.c1})`
        const h = hoist(`R(${g.s},${g.r1},${g.c1},${g.r2},${g.c2})`)
        return rangeOk || g.r2 > g.r1 || g.c2 > g.c1 ? h : `scalar(${h})`
      }
      case 'name': {
        const d = deref(n)
        return d.k === 'name' ? 'E.NAME' : go(d, rangeOk)
      }
      case 'paren':
        return rangeOk ? go(n.a, true) : `(${go(n.a)})`
      case 'un':
        // unary plus keeps a reference (SUM(+A1:A3) sums the range)
        if (n.op === '+') return go(n.a)
        return many(n.a) ? `L1(neg,${go(n.a)})` : `neg(${go(n.a)})`
      case 'pct':
        return many(n.a) ? `L1(pct,${go(n.a)})` : `pct(${go(n.a)})`
      case 'bin': {
        const a = go(n.a)
        const b = go(n.b)
        const op = OPS[n.op] ?? `C[${JSON.stringify(n.op)}]`
        if (many(n.a) || many(n.b)) return `L2(${op},${a},${b})`
        return OPS[n.op] ? `${op}(${a},${b})` : `cmp(${JSON.stringify(n.op)},${a},${b})`
      }
      case 'fn': {
        const spec = FUNCTIONS[n.name]
        if (!spec) return `ext(${JSON.stringify(canonical(n))})`
        if (spec.volatile) volatile = true
        if (n.name === 'SUBTOTAL' || n.name === 'AGGREGATE') nested = true
        const key = /^[A-Z][A-Z0-9_]*$/.test(n.name) ? `.${n.name}` : `[${JSON.stringify(n.name)}]`
        if (spec.lazy) {
          // lazy functions (IF, IFERROR) get each argument as a function; lifted when a lifting argument is an array
          const thunks = n.args.map((a) => `()=>${go(a)}`)
          const check = n.args.flatMap((a, i) => (many(a) && spec.liftAt(i, n.args.length) ? [i] : []))
          return check.length ? `LZ(${JSON.stringify(n.name)},[${check.join(',')}],${thunks.join(',')})` : `F${key}(${thunks.join(',')})`
        }
        const args = n.args.map((a, i) => go(a, spec.range(i)))
        const lifted = n.args.some((a, i) => many(a) && spec.mode(i) !== 'whole')
        // RAND gets its cell
        if (spec.cell) args.unshift(String(s), String(r), String(c))
        return `${lifted ? 'FL' : 'F'}${key}(${args.join(',')})`
      }
    }
  }
  const expr = go(ast)
  const decl = hoisted.length ? `const ${hoisted.map((h, i) => `h${i}=${h}`).join(',')};` : ''
  return { source: `(()=>{${decl}return ()=>${expr}})()`, refs, volatile, nested }
}

export const PRELUDE = '"use strict";const {cell,scalar,add,sub,mul,div,pow,neg,pct,cat,cmp,L1,L2,C,F,FL,LZ,ext,E,R,A}=rt;'

export function compileSources(rt: Runtime, sources: string[]): (() => Arg)[] {
  const R = (s: number, r1: number, c1: number, r2: number, c2: number) => new Rng(s, r1, c1, r2, c2)
  const A = (rows: number, cols: number, v: V[]) => new Arr(rows, cols, rows, cols, v)
  let make: (rt: unknown) => unknown
  try {
    make = new Function('rt', `${PRELUDE}return [${sources.join(',\n')}]`) as typeof make
  } catch (e) {
    // a Content Security Policy without 'unsafe-eval' blocks new Function with an EvalError
    if (e instanceof EvalError) throw new EvalError(`${PACKAGE_NAME} compiles formulas with new Function, and this environment blocks it (a Content Security Policy without 'unsafe-eval'?)`, { cause: e })
    throw e
  }
  return make({ ...rt, R, A }) as (() => Arg)[]
}
