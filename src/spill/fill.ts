/**
 * The cells of a large spilled array past its listed part (`docs/spill.md`): they all hold one value and are
 * not stored, so `A:A*2` costs its data rows, not 1,048,576 cells. The runtime reads them through these.
 */
import type { V } from '../values.js'

/** A rectangle of a sheet whose every cell holds `v`. */
export interface Fill {
  r1: number
  c1: number
  r2: number
  c2: number
  v: V
}

/** The value of a cell that holds nothing stored: a fill's value, else blank. */
export function fillAt(fills: readonly Fill[], r: number, c: number): V {
  for (const f of fills) if (r >= f.r1 && r <= f.r2 && c >= f.c1 && c <= f.c2) return f.v
  return null
}

/** Bounds of a range's occupied part `r2`, `c2` grown over the fills inside the range. */
export function fillBounds(fills: readonly Fill[], g: { r1: number; c1: number; r2: number; c2: number }, r2: number, c2: number): [number, number] {
  for (const f of fills)
    if (f.r1 <= g.r2 && f.r2 >= g.r1 && f.c1 <= g.c2 && f.c2 >= g.c1) {
      r2 = Math.max(r2, Math.min(f.r2, g.r2))
      c2 = Math.max(c2, Math.min(f.c2, g.c2))
    }
  return [r2, c2]
}

/** Values of a range's occupied part (`b`, row by row, `out` filled from the stored cells): the fills' cells get their value. */
export function fillFlat(fills: readonly Fill[], b: { r1: number; c1: number; h: number; w: number }, out: V[]) {
  for (const f of fills) {
    const [r1, r2] = [Math.max(f.r1, b.r1), Math.min(f.r2, b.r1 + b.h - 1)]
    const [c1, c2] = [Math.max(f.c1, b.c1), Math.min(f.c2, b.c1 + b.w - 1)]
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) out[(r - b.r1) * b.w + (c - b.c1)] ??= f.v
  }
}
