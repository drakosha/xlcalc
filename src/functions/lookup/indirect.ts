/**
 * INDIRECT and OFFSET, this engine's own: references found as the formula computes. Each tells the graph what it found
 * (`Cx.found`), so the formula is computed again when a cell there changes (`graph/dynamic.ts`).
 */
import type { Arg, Cx, FnDef } from '../../args.js'
import { MAX_COLS, MAX_ROWS } from '../../model.js'
import { E, isErr, Rng } from '../../values.js'

/** A reference found: given to the graph. */
function found(cx: Cx, g: Rng | ReturnType<Cx['refer']>): Arg {
  if (g instanceof Rng) cx.found(g)
  return g
}

/** One axis of OFFSET: from `at` moved by `by`, `size` long (a negative size reaches back), or undefined off the sheet. */
function moved(at: number, by: number, size: number, max: number): [number, number] | undefined {
  const from = at + by
  const [lo, hi] = size > 0 ? [from, from + size - 1] : [from + size + 1, from]
  return lo >= 0 && hi < max ? [lo, hi] : undefined
}

/** OFFSET: a reference of `height` × `width` cells (the reference's by default) `rows` down and `cols` right of its top-left cell. */
function offset(cx: Cx, ref: Arg, rows: number, cols: number, height?: number, width?: number): Arg {
  if (isErr(ref)) return ref
  if (!(ref instanceof Rng)) return E.VALUE
  const h = height ?? ref.r2 - ref.r1 + 1
  const w = width ?? ref.c2 - ref.c1 + 1
  if (h === 0 || w === 0) return E.REF
  const r = moved(ref.r1, rows, h, MAX_ROWS)
  const c = moved(ref.c1, cols, w, MAX_COLS)
  return r && c ? found(cx, new Rng(ref.s, r[0], c[0], r[1], c[1])) : E.REF
}

export const functions: Record<string, FnDef> = {
  INDIRECT: {
    args: 's b?',
    cx: true,
    cell: true,
    dynamic: true,
    array: true,
    fn: (cx: Cx, s: number, r: number, c: number, text: string, a1 = true) => found(cx, cx.refer(text, a1, s, r, c)),
  },
  OFFSET: { args: 'R w w w? w?', cx: true, dynamic: true, array: true, fn: offset },
}
