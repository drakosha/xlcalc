/**
 * BESSELI, BESSELJ, BESSELK, BESSELY. Ported from bessel 1.0.2 (bessel.js), Copyright (C) 2013-present SheetJS,
 * Apache License 2.0 (THIRD_PARTY.md). Changed: Excel's J1 above 8 and start of the I recurrence, argument checks, loop caps.
 */
import { E, MIN_NORMAL } from '../../values.js'
import { horner, MAX_TRIG } from '../common.js'
import { gammaln } from '../jstat.js'
import type { FnDef } from '../../args.js'

/** Recurrence steps allowed: enough for any order below MAX_TRIG, Excel's limit for J and Y; beyond it #NUM! rather than a long loop. */
const MAX_STEPS = 1.5e8
/** Below this (the smallest normal number) Excel's answer is 0. */
const LOG_TINY = Math.log(MIN_NORMAL)

/** Forward recurrence from orders 0 and 1 (sign −1 for J and Y, +1 for K); stops once it overflows. */
function forward(x: number, n: number, f0: number, f1: number, sign: number): number {
  if (n === 0) return f0
  if (n === 1 || (f0 === 0 && f1 === 0)) return f1
  if (n > MAX_STEPS) return NaN
  const tdx = 2 / x
  let f2 = f1
  for (let o = 1; o < n && Number.isFinite(f2); ++o) {
    f2 = f1 * o * tdx + sign * f0
    f0 = f1
    f1 = f2
  }
  return f2
}

/** Whether (x/2)^n/n!·e^extra, a bound of |J_n(x)| (extra 0) and I_n(x) (extra x), is below Excel's smallest number. */
const underflows = (x: number, n: number, extra: number) => n * Math.log(x / 2) - gammaln(n + 1) + extra < LOG_TINY

const W = 0.636619772

const J0_A1 = [57568490574.0, -13362590354.0, 651619640.7, -11214424.18, 77392.33017, -184.9052456].reverse()
const J0_A2 = [57568490411.0, 1029532985.0, 9494680.718, 59272.64853, 267.8532712, 1.0].reverse()
const J0_B1 = [1.0, -0.1098628627e-2, 0.2734510407e-4, -0.2073370639e-5, 0.2093887211e-6].reverse()
const J0_B2 = [-0.1562499995e-1, 0.1430488765e-3, -0.6911147651e-5, 0.7621095161e-6, -0.934935152e-7].reverse()

function j0(x: number): number {
  if (x < 8) return horner(J0_A1, x * x) / horner(J0_A2, x * x)
  const xx = x - 0.785398164
  const y = 64 / (x * x)
  return Math.sqrt(W / x) * (Math.cos(xx) * horner(J0_B1, y) - (Math.sin(xx) * horner(J0_B2, y) * 8) / x)
}

const J1_A1 = [72362614232.0, -7895059235.0, 242396853.1, -2972611.439, 15704.4826, -30.16036606].reverse()
const J1_A2 = [144725228442.0, 2300535178.0, 18583304.74, 99447.43394, 376.9991397, 1.0].reverse()
const J1_B1 = [1.0, 0.183105e-2, -0.3516396496e-4, 0.2457520174e-5, -0.240337019e-6].reverse()
/** Excel's J1 above 8 repeats the y² coefficient (bit for bit at x = 8, 8.5, 10, 15, 30); Y1 has the right one. */
const J1_B1_EXCEL = [1.0, 0.183105e-2, -0.3516396496e-4, -0.3516396496e-4, 0.2457520174e-5, -0.240337019e-6].reverse()
const J1_B2 = [0.04687499995, -0.2002690873e-3, 0.8449199096e-5, -0.88228987e-6, 0.105787412e-6].reverse()

function j1(x: number): number {
  const ax = Math.abs(x)
  if (ax < 8) return (x * horner(J1_A1, x * x)) / horner(J1_A2, x * x)
  const xx = ax - 2.356194491
  const y = 64 / (x * x)
  const a = Math.sqrt(W / ax) * (Math.cos(xx) * horner(J1_B1_EXCEL, y) - (Math.sin(xx) * horner(J1_B2, y) * 8) / ax)
  return x < 0 ? -a : a
}

function besselj(x: number, n: number): number {
  if (x < 0) return (n % 2 ? -1 : 1) * besselj(-x, n)
  if (n === 0) return j0(x)
  if (n === 1) return j1(x)
  if (x === 0) return 0
  if (x > n) return forward(x, n, j0(x), j1(x), -1)
  if (underflows(x, n, 0)) return 0
  if (n > MAX_STEPS) return NaN
  // Miller's downward recurrence, normalised by J0 + 2(J2 + J4 + …) = 1
  const m = 2 * Math.floor((n + Math.floor(Math.sqrt(40 * n))) / 2)
  let jsum = false
  let bjp = 0
  let sum = 0
  let bj = 1
  let ret = 0
  const tox = 2 / x
  for (let j = m; j > 0; j--) {
    const bjm = j * tox * bj - bjp
    bjp = bj
    bj = bjm
    if (Math.abs(bj) > 1e10) {
      bj *= 1e-10
      bjp *= 1e-10
      ret *= 1e-10
      sum *= 1e-10
    }
    if (jsum) sum += bj
    jsum = !jsum
    if (j === n) ret = bjp
  }
  return ret / (2 * sum - bj)
}

const Y0_A1 = [-2957821389.0, 7062834065.0, -512359803.6, 10879881.29, -86327.92757, 228.4622733].reverse()
const Y0_A2 = [40076544269.0, 745249964.8, 7189466.438, 47447.2647, 226.1030244, 1.0].reverse()
const Y0_B2 = [-0.1562499995e-1, 0.1430488765e-3, -0.6911147651e-5, 0.7621095161e-6, -0.934945152e-7].reverse()

function y0(x: number): number {
  if (x < 8) return horner(Y0_A1, x * x) / horner(Y0_A2, x * x) + W * j0(x) * Math.log(x)
  const xx = x - 0.785398164
  const y = 64 / (x * x)
  return Math.sqrt(W / x) * (Math.sin(xx) * horner(J0_B1, y) + (Math.cos(xx) * horner(Y0_B2, y) * 8) / x)
}

const Y1_A1 = [-0.4900604943e13, 0.127527439e13, -0.5153438139e11, 0.7349264551e9, -0.4237922726e7, 0.8511937935e4].reverse()
const Y1_A2 = [0.249958057e14, 0.4244419664e12, 0.3733650367e10, 0.2245904002e8, 0.102042605e6, 0.3549632885e3, 1].reverse()

function y1(x: number): number {
  if (x < 8) return (x * horner(Y1_A1, x * x)) / horner(Y1_A2, x * x) + W * (j1(x) * Math.log(x) - 1 / x)
  const xx = x - 2.356194491
  const y = 64 / (x * x)
  return Math.sqrt(W / x) * (Math.sin(xx) * horner(J1_B1, y) + (Math.cos(xx) * horner(J1_B2, y) * 8) / x)
}

const I0_A = [1.0, 3.5156229, 3.0899424, 1.2067492, 0.2659732, 0.360768e-1, 0.45813e-2].reverse()
const I0_B = [0.39894228, 0.1328592e-1, 0.225319e-2, -0.157565e-2, 0.916281e-2, -0.2057706e-1, 0.2635537e-1, -0.1647633e-1, 0.392377e-2].reverse()

function i0(x: number): number {
  const ax = Math.abs(x)
  if (ax < 3.75) return horner(I0_A, (x * x) / (3.75 * 3.75))
  return (Math.exp(ax) / Math.sqrt(ax)) * horner(I0_B, 3.75 / ax)
}

const I1_A = [0.5, 0.87890594, 0.51498869, 0.15084934, 0.2658733e-1, 0.301532e-2, 0.32411e-3].reverse()
const I1_B = [0.39894228, -0.3988024e-1, -0.362018e-2, 0.163801e-2, -0.1031555e-1, 0.2282967e-1, -0.2895312e-1, 0.1787654e-1, -0.420059e-2].reverse()

function i1(x: number): number {
  const ax = Math.abs(x)
  if (ax < 3.75) return x * horner(I1_A, (x * x) / (3.75 * 3.75))
  return ((x < 0 ? -1 : 1) * Math.exp(ax) * horner(I1_B, 3.75 / ax)) / Math.sqrt(ax)
}

function besseli(x: number, n: number): number {
  if (n === 0) return i0(x)
  if (n === 1) return i1(x)
  const ax = Math.abs(x)
  if (ax === 0) return 0
  if (underflows(ax, n, ax)) return 0
  if (n > MAX_STEPS) return NaN
  // Miller's downward recurrence, normalised by I0
  const m = 2 * (n + Math.floor(Math.sqrt(40 * n)))
  const tox = 2 / ax
  let bip = 0
  let bi = 1
  let ret = 0
  for (let j = m; j > 0; j--) {
    const bim = j * tox * bi + bip
    bip = bi
    bi = bim
    if (Math.abs(bi) > 1e10) {
      bi *= 1e-10
      bip *= 1e-10
      ret *= 1e-10
    }
    if (j === n) ret = bip
  }
  ret *= i0(x) / bi
  return x < 0 && n % 2 ? -ret : ret
}

const K0_A = [-0.57721566, 0.4227842, 0.23069756, 0.348859e-1, 0.262698e-2, 0.1075e-3, 0.74e-5].reverse()
const K0_B = [1.25331414, -0.7832358e-1, 0.2189568e-1, -0.1062446e-1, 0.587872e-2, -0.25154e-2, 0.53208e-3].reverse()

function k0(x: number): number {
  if (x <= 2) return -Math.log(x / 2) * i0(x) + horner(K0_A, (x * x) / 4)
  return (Math.exp(-x) / Math.sqrt(x)) * horner(K0_B, 2 / x)
}

const K1_A = [1.0, 0.15443144, -0.67278579, -0.18156897, -0.1919402e-1, -0.110404e-2, -0.4686e-4].reverse()
const K1_B = [1.25331414, 0.23498619, -0.365562e-1, 0.1504268e-1, -0.780353e-2, 0.325614e-2, -0.68245e-3].reverse()

function k1(x: number): number {
  if (x <= 2) return Math.log(x / 2) * i1(x) + (1 / x) * horner(K1_A, (x * x) / 4)
  return (Math.exp(-x) / Math.sqrt(x)) * horner(K1_B, 2 / x)
}

/** Excel: the order is truncated, a negative order is #NUM!. */
const order = (n: number, f: (n: number) => number) => (n < 0 ? E.NUM : f(Math.trunc(n)))

export const functions: Record<string, FnDef> = {
  BESSELI: { args: 'x x', fn: (x: number, n: number) => order(n, (k) => besseli(x, k)) },
  BESSELJ: { args: 'x x', fn: (x: number, n: number) => (Math.abs(x) >= MAX_TRIG ? E.NUM : order(n, (k) => besselj(x, k))) },
  BESSELK: { args: 'x x', fn: (x: number, n: number) => (x <= 0 ? E.NUM : order(n, (k) => forward(x, k, k0(x), k1(x), 1))) },
  BESSELY: { args: 'x x', fn: (x: number, n: number) => (x <= 0 || x >= MAX_TRIG ? E.NUM : order(n, (k) => forward(x, k, y0(x), y1(x), -1))) },
}
