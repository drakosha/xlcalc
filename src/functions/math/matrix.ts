/** MDETERM: the determinant by Gaussian elimination with partial pivoting. This engine's own. */
import { E, type CErr } from '../../values.js'
import type { FnDef, Grid } from '../../args.js'

function mdeterm(g: Grid): number | CErr {
  const n = g.rows
  // a square range filled with numbers only (a blank, text or a boolean is #VALUE!)
  if (g.cols !== n || g.h !== n || g.w !== n || g.v.some((v) => typeof v !== 'number')) return E.VALUE
  const a = Array.from({ length: n }, (_, i) => g.v.slice(i * n, (i + 1) * n) as number[])
  let det = 1
  for (let k = 0; k < n; k++) {
    let p = k
    for (let i = k + 1; i < n; i++) if (Math.abs(a[i][k]) > Math.abs(a[p][k])) p = i
    if (a[p][k] === 0) return 0
    if (p !== k) {
      ;[a[p], a[k]] = [a[k], a[p]]
      det = -det
    }
    det *= a[k][k]
    for (let i = k + 1; i < n; i++) {
      const f = a[i][k] / a[k][k]
      for (let j = k + 1; j < n; j++) a[i][j] -= f * a[k][j]
    }
  }
  return det
}

export const functions: Record<string, FnDef> = {
  MDETERM: { args: 'M', fn: mdeterm },
}
