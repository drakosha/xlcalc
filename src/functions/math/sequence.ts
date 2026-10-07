/** SEQUENCE and RANDARRAY (Excel 365), this engine's own. An empty array is #CALC! in Excel, #VALUE! here. */
import type { Cx, FnDef } from '../../args.js'
import { Arr, EMPTY } from '../../arrays/value.js'
import { E, type V } from '../../values.js'

/** Excel's grid: an array larger than a sheet cannot be made. */
const MAX_CELLS = 1_048_576 * 16_384

/** The size of a new array: negative #VALUE!, zero #CALC!, else rows × cols. */
function size(rows: number, cols: number): [number, number] | typeof E.VALUE {
  const r = Math.trunc(rows)
  const c = Math.trunc(cols)
  if (r < 0 || c < 0 || r * c > MAX_CELLS) return E.VALUE
  return r === 0 || c === 0 ? EMPTY : [r, c]
}

function build(rows: number, cols: number, at: (k: number) => V): Arr | typeof E.VALUE {
  const s = size(rows, cols)
  if (!Array.isArray(s)) return s
  const [r, c] = s
  return new Arr(r, c, r, c, Array.from({ length: r * c }, (_, k) => at(k)))
}

export const functions: Record<string, FnDef> = {
  RANDARRAY: {
    args: 'n? n? n? n? b?',
    cx: true,
    cell: true,
    volatile: true,
    array: true,
    // each element draws from the clock seed, the cell and its position, as RAND does
    fn: (cx: Cx, s: number, r: number, c: number, rows = 1, cols = 1, lo = 0, hi = 1, whole = false) => {
      if (lo > hi || (whole && (!Number.isInteger(lo) || !Number.isInteger(hi)))) return E.VALUE
      return build(rows, cols, (k) => (whole ? lo + Math.floor(cx.rand(s, r, c, k) * (hi - lo + 1)) : lo + cx.rand(s, r, c, k) * (hi - lo)))
    },
  },
  SEQUENCE: { args: 'n n? n? n?', array: true, fn: (rows = 1, cols = 1, start = 1, step = 1) => build(rows, cols, (k) => start + k * step) },
}
