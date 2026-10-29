/**
 * The functions that call a function for each value (`exec/lambda.ts`): MAP, REDUCE, SCAN, BYROW, BYCOL and MAKEARRAY,
 * this engine's own. A function given where one goes is called with as many arguments as the function takes (a LAMBDA's
 * parameters, a built-in function's counts), else the call is #VALUE!, as Microsoft's documentation says. A range gives
 * each cell, row or column as a reference, an array its values; each answer that has to be one value and is not (an
 * array of several, a function) makes the whole call #CALC!.
 */
import type { Arg, Cx, FnDef } from '../args.js'
import { arr, Arr } from '../arrays/value.js'
import { allot } from '../exec/budget.js'
import { callValue } from '../exec/lambda.js'
import { Areas, asValue, E, isErr, Lambda, Rng, type V } from '../values.js'
import { NEWER_MAX_ARGS } from './common.js'

/** A function that takes `n` arguments, or undefined: a LAMBDA with `n` parameters, a built-in function within its counts. */
export const taking = (f: Arg | undefined, n: number): Lambda | undefined => (f instanceof Lambda && f.takes(n) ? f : undefined)

/** An argument's rows and columns, and its element at a row and a column: a range's cell as a reference, an array's value. */
export interface Items {
  rows: number
  cols: number
  at(r: number, c: number): Arg
}

/** The items of an argument; undefined for several areas. */
export function itemsOf(x: Arg): Items | undefined {
  if (x instanceof Areas) return undefined
  if (x instanceof Rng) return { rows: x.r2 - x.r1 + 1, cols: x.c2 - x.c1 + 1, at: (r, c) => new Rng(x.s, x.r1 + r, x.c1 + c, x.r1 + r, x.c1 + c) }
  if (x instanceof Arr) return { rows: x.rows, cols: x.cols, at: (r, c) => x.at(r, c) }
  return { rows: 1, cols: 1, at: () => x }
}

/** A row (`row`) or a column of an argument at `i`: of a range a reference, of an array an array. */
function line(x: Rng | Arr, i: number, row: boolean): Arg {
  if (x instanceof Rng) return row ? new Rng(x.s, x.r1 + i, x.c1, x.r1 + i, x.c2) : new Rng(x.s, x.r1, x.c1 + i, x.r2, x.c1 + i)
  const n = row ? x.cols : x.rows
  return arr(row ? 1 : n, row ? n : 1, Array.from({ length: n }, (_, k) => (row ? x.at(i, k) : x.at(k, i))))
}

/** The one value an answer is (a function too, which an array holds), or undefined for an array of several values. */
export function one(x: Arg, cx: Cx): V | undefined {
  if (x instanceof Lambda) return asValue(x)
  if (x instanceof Areas) return undefined
  if (x instanceof Rng) return x.r1 === x.r2 && x.c1 === x.c2 ? cx.scalar(x) : undefined
  if (x instanceof Arr) return x.rows === 1 && x.cols === 1 ? x.at(0, 0) : undefined
  return x
}

/** An array of `rows` × `cols` answers, each one value, or #CALC! when one is not. */
function answers(cx: Cx, rows: number, cols: number, answer: (r: number, c: number) => Arg): Arr | typeof E.CALC {
  allot(rows * cols)
  const v: V[] = new Array(rows * cols)
  for (let r = 0; r < rows; r++)
    for (let c = 0; c < cols; c++) {
      const x = one(answer(r, c), cx)
      if (x === undefined) return E.CALC
      v[r * cols + c] = x
    }
  return arr(rows, cols, v)
}

/**
 * MAP(array1, …, function): the function of the arrays' values at each position; where a smaller array has none, #N/A
 * without a call. The first array an error is the answer, before the function is looked at; a later one #VALUE!.
 */
function map(cx: Cx, first: Arg, rest: Arg[]): Arg {
  if (isErr(first)) return first
  const arrays = [first, ...rest.slice(0, -1)]
  if (arrays.some(isErr)) return E.VALUE
  const f = taking(rest.at(-1), rest.length)
  if (!f) return E.VALUE
  const all = arrays.map(itemsOf)
  if (all.some((x) => !x)) return E.VALUE
  const items = all as Items[]
  const rows = Math.max(...items.map((x) => x.rows))
  const cols = Math.max(...items.map((x) => x.cols))
  return answers(cx, rows, cols, (r, c) => (items.every((x) => r < x.rows && c < x.cols) ? callValue(f, items.map((x) => x.at(r, c))) : E.NA))
}

/**
 * REDUCE and SCAN: the value accumulated over the array's values row by row, or with `each` (SCAN) the array of each
 * step's value. Without an initial value the first value starts it, as Microsoft's documentation says.
 */
function accumulate(cx: Cx, initial: Arg | undefined, array: Arg, fn: Arg, each?: V[]): Arg {
  if (isErr(array)) return array
  const f = taking(fn, 2)
  const items = itemsOf(array)
  if (!f || !items) return E.VALUE
  const n = items.rows * items.cols
  if (each) allot(n)
  const at = (i: number) => items.at(Math.floor(i / items.cols), i % items.cols)
  // SCAN keeps each step's value, one value or the whole call is #CALC!
  const kept = (x: Arg) => {
    const v = each && one(x, cx)
    if (v !== undefined) each!.push(v)
    return !each || v !== undefined
  }
  let i = initial === undefined ? 1 : 0
  let acc = initial ?? at(0)
  if (i && !kept(acc)) return E.CALC
  for (; i < n; i++) {
    acc = callValue(f, [acc, at(i)])
    if (!kept(acc)) return E.CALC
  }
  return each ? arr(items.rows, items.cols, each) : acc
}

/** BYROW and BYCOL: the function of each row (a column of answers) or each column (a row of them). */
const by =
  (row: boolean) =>
  (cx: Cx, x: Arg, fn: Arg): Arg => {
    const f = taking(fn, 1)
    if (isErr(x)) return x
    if (!f || x instanceof Areas) return E.VALUE
    const lines = x instanceof Rng || x instanceof Arr ? x : arr(1, 1, [x instanceof Lambda ? asValue(x) : x])
    const n = row ? (lines instanceof Rng ? lines.r2 - lines.r1 + 1 : lines.rows) : lines instanceof Rng ? lines.c2 - lines.c1 + 1 : lines.cols
    return answers(cx, row ? n : 1, row ? 1 : n, (r, c) => callValue(f, [line(lines, row ? r : c, row)]))
  }

export const functions: Record<string, FnDef> = {
  BYCOL: { args: 'R R', cx: true, array: true, fn: by(false) },
  BYROW: { args: 'R R', cx: true, array: true, fn: by(true) },
  // a size is a whole number from 1, else #VALUE! (2.7 too)
  MAKEARRAY: {
    args: 'n n R',
    cx: true,
    array: true,
    fn: (cx: Cx, rows: number, cols: number, fn: Arg) => {
      const f = taking(fn, 2)
      if (!f || !Number.isInteger(rows) || !Number.isInteger(cols) || rows < 1 || cols < 1) return E.VALUE
      return answers(cx, rows, cols, (r, c) => callValue(f, [r + 1, c + 1]))
    },
  },
  MAP: { args: 'R R', rep: 1, max: NEWER_MAX_ARGS, cx: true, array: true, fn: map },
  REDUCE: { args: 'R R R', omit: { at: { 0: 'default' } }, cx: true, array: true, fn: (cx: Cx, initial: Arg | undefined, array: Arg, fn: Arg) => accumulate(cx, initial, array, fn) },
  SCAN: { args: 'R R R', omit: { at: { 0: 'default' } }, cx: true, array: true, fn: (cx: Cx, initial: Arg | undefined, array: Arg, fn: Arg) => accumulate(cx, initial, array, fn, []) },
}
