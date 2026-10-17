/**
 * Trigonometric and hyperbolic functions. ACOS … SINH ported from formula.js 4.6.1 (src/math-trig.js), Copyright (c)
 * 2014 Sutoiku, Inc., MIT License (THIRD_PARTY.md); the rest this engine's own.
 */
import { E, type CErr } from '../../values.js'
import type { FnDef } from '../../args.js'
import { MAX_TRIG } from '../common.js'

/** f(x) within Excel's range; `zero` is the answer at 0 where f has a pole there. */
const trig = (f: (x: number) => number, zero?: CErr): FnDef => ({
  args: 'n',
  fn: (x: number) => (Math.abs(x) >= MAX_TRIG ? E.NUM : zero && x === 0 ? zero : f(x)),
})

export const functions: Record<string, FnDef> = {
  ACOS: { args: 'n', fn: Math.acos },
  ACOSH: { args: 'n', fn: (x: number) => Math.log(x + Math.sqrt(x * x - 1)) },
  // atan2: in (0, π) and exact for huge x (ACOT(1E10) is 1E-10)
  ACOT: { args: 'n', fn: (x: number) => Math.atan2(1, x) },
  ACOTH: { args: 'n', fn: (x: number) => 0.5 * Math.log((x + 1) / (x - 1)) },
  ASIN: { args: 'n', fn: Math.asin },
  ASINH: { args: 'n', fn: (x: number) => Math.log(x + Math.sqrt(x * x + 1)) },
  ATAN: { args: 'n', fn: Math.atan },
  ATAN2: { args: 'n n', fn: (x: number, y: number) => (x === 0 && y === 0 ? E.DIV0 : Math.atan2(y, x)) },
  ATANH: { args: 'n', fn: (x: number) => Math.log((1 + x) / (1 - x)) / 2 },
  COS: trig(Math.cos),
  COSH: { args: 'n', fn: (x: number) => (Math.exp(x) + Math.exp(-x)) / 2 },
  COT: trig((x) => 1 / Math.tan(x), E.DIV0),
  COTH: { args: 'n', fn: (x: number) => (x === 0 ? E.DIV0 : 1 / Math.tanh(x)) },
  CSC: trig((x) => 1 / Math.sin(x), E.DIV0),
  // 1/sinh keeps the digits of a small x that the difference of exponentials loses (CSCH(1E-12) is 1E12)
  CSCH: { args: 'n', fn: (x: number) => (x === 0 ? E.DIV0 : 1 / Math.sinh(x)) },
  DEGREES: { args: 'n', fn: (x: number) => (x * 180) / Math.PI },
  PI: { args: '', fn: () => Math.PI },
  RADIANS: { args: 'n', fn: (x: number) => (x * Math.PI) / 180 },
  SEC: trig((x) => 1 / Math.cos(x)),
  SECH: { args: 'n', fn: (x: number) => 2 / (Math.exp(x) + Math.exp(-x)) },
  SIN: trig(Math.sin),
  SINH: { args: 'n', fn: (x: number) => (Math.exp(x) - Math.exp(-x)) / 2 },
  TAN: trig(Math.tan),
  TANH: { args: 'n', fn: Math.tanh },
}
