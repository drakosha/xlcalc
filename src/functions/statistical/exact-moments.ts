/** Exact dyadic sums for the exceptional cases where double-precision moments lose their scale or significant bits. */
import { spend } from '../../exec/budget.js'

const bits = new DataView(new ArrayBuffer(8))
const FRACTION = (1n << 52n) - 1n

/** A finite double as an integer times a power of two, including subnormal inputs. */
function dyadic(x: number): [bigint, number] {
  bits.setFloat64(0, x)
  const raw = bits.getBigUint64(0)
  const exponent = Number((raw >> 52n) & 2047n)
  const mantissa = (raw & FRACTION) | (exponent ? 1n << 52n : 0n)
  return [raw >> 63n ? -mantissa : mantissa, exponent ? exponent - 1075 : -1074]
}

export interface ExactMoments {
  n: bigint
  sx: bigint
  sy: bigint
  /** n times each central sum; their scales are 2*ex, 2*ey and ex+ey. */
  xx: bigint
  yy: bigint
  xy: bigint
  ex: number
  ey: number
}

export function exactMoments(xs: readonly number[], ys?: readonly number[]): ExactMoments | undefined {
  let sx = 0n, sy = 0n, xx = 0n, yy = 0n, xy = 0n
  let ex = 0, ey = 0, haveX = false, haveY = false
  for (let i = 0; i < xs.length; i++) {
    spend(1)
    const x = xs[i], y = ys?.[i] ?? 0
    if (!Number.isFinite(x) || !Number.isFinite(y)) return undefined
    const [mx, px] = dyadic(x), [my, py] = dyadic(y)
    if (mx && !haveX) { ex = px; haveX = true }
    if (my && !haveY) { ey = py; haveY = true }
    if (mx && px < ex) {
      const d = BigInt(ex - px)
      sx <<= d; xx <<= 2n * d; xy <<= d; ex = px
    }
    if (my && py < ey) {
      const d = BigInt(ey - py)
      sy <<= d; yy <<= 2n * d; xy <<= d; ey = py
    }
    const ix = mx ? mx << BigInt(px - ex) : 0n
    const iy = my ? my << BigInt(py - ey) : 0n
    sx += ix; sy += iy; xx += ix * ix; yy += iy * iy; xy += ix * iy
  }
  const n = BigInt(xs.length)
  return { n, sx, sy, xx: n * xx - sx * sx, yy: n * yy - sy * sy, xy: n * xy - sx * sy, ex, ey }
}

/** Keep more than a double's significand before the final division, without converting an enormous integer to Infinity. */
function fraction(numerator: bigint, denominator: bigint, exponent: number): [number, number] {
  const sign = numerator < 0n ? -1 : 1
  const n = numerator < 0n ? -numerator : numerator
  const ns = Math.max(0, n.toString(2).length - 62), ds = Math.max(0, denominator.toString(2).length - 62)
  return [sign * Number(n >> BigInt(ns)) / Number(denominator >> BigInt(ds)), exponent + ns - ds]
}

function scale(x: number, exponent: number): number {
  if (x === 0 || !Number.isFinite(x)) return x
  while (exponent > 1023) { x *= 2 ** 1023; exponent -= 1023 }
  while (exponent < -1022) { x *= 2 ** -1022; exponent += 1022 }
  return x * 2 ** exponent
}

export function ratio(n: bigint, d: bigint, exponent = 0): number {
  const [value, e] = fraction(n, d, exponent)
  return scale(value, e)
}

/** Take the root before applying the exponent: the variance may overflow/underflow while its root is representable. */
export function rootRatio(n: bigint, d: bigint, exponent = 0): number {
  let [value, e] = fraction(n, d, exponent)
  if (e % 2 !== 0) { value *= 2; e-- }
  return scale(Math.sqrt(value), e / 2)
}

/** The fitted y at x, with the means and the slope combined before rounding. */
export function predict(f: ExactMoments, x: number): number {
  if (!Number.isFinite(x)) return NaN
  const [m, p] = dyadic(x)
  const e = m ? Math.min(f.ex, p) : f.ex
  const delta = (m ? m << BigInt(p - e) : 0n) * f.n - (f.sx << BigInt(f.ex - e))
  const numerator = (f.sy * f.xx << BigInt(f.ex - e)) + f.xy * delta
  return ratio(numerator, f.n * f.xx, f.ey - f.ex + e)
}
