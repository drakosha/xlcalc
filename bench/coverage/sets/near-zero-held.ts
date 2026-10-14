/**
 * Held out: a sum or a difference near 0. Excel gave 0 for `=0.5-0.4-0.1` and `=0.1+0.2-Z1` (Z1 0.3) in a user's
 * model, and -2.78E-17 for `=1*(0.5-0.4-0.1)` and `=(0.5-0.4-0.1)`, where the engine gave the double's answer each
 * time. The rule predicted: where the formula's last operation, outside any parentheses, is + or -, an answer within
 * 2^-49 of its larger operand's magnitude is 0; parentheses, a later operation, a function's argument or a unary
 * minus around it keep the double. The cases: `last` the formulas themselves, `stored` the same formulas in cells of
 * sheet P (formulas of the old kind) read by another cell times 2^60, so the comparison's tolerance sees 0 apart from
 * the double, `array` a spilled array's elements, and the sweeps `x+m*u-x` for m units u of x's last place (x 1, 1.9,
 * 9.5, 1E+15, 1E-200), `1+2^-k-1`, and the operands the other way round, each read as m times u.
 * Answers: `test/excel/cases/near-zero-held.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector, xl } from './arrays/common.js'
import { n } from '../vals.js'

/** The rule predicted: a sum or difference within 2^-49 of its larger operand's magnitude is 0. */
function last(a: number, b: number, op: '+' | '-'): number {
  const r = op === '+' ? a + b : a - b
  return Math.abs(r) <= 2 ** -49 * Math.max(Math.abs(a), Math.abs(b)) ? 0 : r
}

/** A formula and its answer predicted; `cell`: also put in a cell of P and read from there. */
interface Case {
  f: string
  v: Val
  cell?: true
}

/** P!B1 0.3 and P!B2 0.1 for the formulas that read a cell. */
const LAST: Case[] = [
  // the user's
  { f: '0.5-0.4-0.1', v: n(last(0.5 - 0.4, 0.1, '-')), cell: true },
  { f: '1*(0.5-0.4-0.1)', v: n(1 * (0.5 - 0.4 - 0.1)), cell: true },
  { f: '(0.5-0.4-0.1)', v: n(0.5 - 0.4 - 0.1), cell: true },
  { f: '0.1+0.2-P!B1', v: n(last(0.1 + 0.2, 0.3, '-')), cell: true },
  { f: 'SUM(0.1,0.2)-P!B1', v: n(last(0.1 + 0.2, 0.3, '-')), cell: true },
  { f: '-0.1-0.2+P!B1', v: n(last(-0.1 - 0.2, 0.3, '+')), cell: true },
  { f: '(0.1+0.2-P!B1)/1', v: n(0.1 + 0.2 - 0.3), cell: true },
  { f: '0.1+0.2-P!B1+0', v: n(last(0.1 + 0.2 - 0.3, 0, '+')), cell: true },
  // the last operation with other operands before it, in parentheses of its own operand, a cell's
  { f: '0+0.1+0.2-0.3', v: n(last(0 + 0.1 + 0.2, 0.3, '-')), cell: true },
  { f: '0.1+0.2-0.3*1', v: n(last(0.1 + 0.2, 0.3, '-')), cell: true },
  { f: '0.1+0.2-(0.3)', v: n(last(0.1 + 0.2, 0.3, '-')), cell: true },
  { f: '(0.1+0.2)-0.3', v: n(last(0.1 + 0.2, 0.3, '-')), cell: true },
  { f: '0.1-(0.3-0.2)', v: n(last(0.1, 0.3 - 0.2, '-')), cell: true },
  { f: 'P!B2+0.2-P!B1', v: n(last(0.1 + 0.2, 0.3, '-')), cell: true },
  { f: '1-0.9-0.1', v: n(last(1 - 0.9, 0.1, '-')), cell: true },
  { f: '0.3-0.2-0.1', v: n(last(0.3 - 0.2, 0.1, '-')), cell: true },
  { f: '-0.3+0.1+0.2', v: n(last(-0.3 + 0.1, 0.2, '+')), cell: true },
  { f: '0.1*3-0.3', v: n(last(0.1 * 3, 0.3, '-')), cell: true },
  // operands of other types: a boolean, text, a percent
  { f: 'TRUE-0.9-0.1', v: n(last(1 - 0.9, 0.1, '-')), cell: true },
  { f: '"0.1"+0.2-0.3', v: n(last(0.1 + 0.2, 0.3, '-')), cell: true },
  { f: '50%-0.4-0.1', v: n(last(0.5 - 0.4, 0.1, '-')), cell: true },
  // unary signs: a minus or a plus around it keeps the double, a plus on the first operand does not
  { f: '-(0.1+0.2-0.3)', v: n(-(0.1 + 0.2 - 0.3)), cell: true },
  { f: '+0.1+0.2-0.3', v: n(last(0.1 + 0.2, 0.3, '-')), cell: true },
  { f: '+(0.1+0.2-0.3)', v: n(0.1 + 0.2 - 0.3), cell: true },
  // a later operation, no cancellation in the last one
  { f: '(0.1+0.2-0.3)+0', v: n(last(0.1 + 0.2 - 0.3, 0, '+')), cell: true },
  { f: '0-(0.1+0.2-0.3)', v: n(last(0, 0.1 + 0.2 - 0.3, '-')), cell: true },
  { f: '2E-17-1E-17', v: n(last(2e-17, 1e-17, '-')), cell: true },
  // a large operand: 0.3 is 0.25 at 1E+15, which the rule takes as 0
  { f: '1E+15+0.3-1E+15', v: n(last(1e15 + 0.3, 1e15, '-')), cell: true },
  // inside a function: SUM's own sum, an argument, LET's calculation
  { f: 'SUM(0.1,0.2,-0.3)', v: n(0.1 + 0.2 - 0.3), cell: true },
  { f: 'SUM(0.1,0.2)-0.3', v: n(last(0.1 + 0.2, 0.3, '-')), cell: true },
  { f: 'ROUND(0.1+0.2-0.3,30)', v: n(0.1 + 0.2 - 0.3), cell: true },
  { f: 'ABS(0.1+0.2-0.3)', v: n(0.1 + 0.2 - 0.3), cell: true },
  { f: 'IF(TRUE,0.1+0.2-0.3)', v: n(0.1 + 0.2 - 0.3), cell: true },
  { f: 'LET(_xlpm.x,0.1,_xlpm.x+0.2-0.3)', v: n(0.1 + 0.2 - 0.3), cell: true },
  // as text: the operation before & keeps the double
  { f: '0.1+0.2-0.3&""', v: { t: 's', v: '5.55111512312578E-17' } },
]

/** The unit in the last place of a positive double. */
const ulp = (x: number) => 2 ** (Math.floor(Math.log2(x)) - 52)

/** The sweeps: a formula in P, its reading (the value over its unit) and the reading predicted. */
function sweeps(): [group: string, f: string, unit: string, v: number][] {
  const out: [string, string, string, number][] = []
  for (const [x, text, unit] of [[1, '1', '2^-52'], [1.9, '1.9', '2^-52'], [9.5, '9.5', '2^-49'], [1e15, '1E+15', '0.125'], [1e-200, '1E-200', '2^-717']] as const) {
    const u = ulp(x)
    for (let m = 1; m <= 20; m++) out.push([`sweep:${text}`, `${text}+${m}*${unit}-${text}`, unit, last(x + m * u, x, '-') / u])
  }
  for (let k = 40; k <= 48; k++) out.push(['sweep:k', `1+2^-${k}-1`, `2^-${k}`, last(1 + 2 ** -k, 1, '-') / 2 ** -k])
  for (const m of [4, 8, 9, 16]) out.push(['sweep:back', `1-(1+${m}*2^-52)`, '2^-52', last(1, 1 + m * 2 ** -52, '-') / 2 ** -52])
  for (const m of [4, 8, 9, 16]) out.push(['sweep:back', `-1+(1+${m}*2^-52)`, '2^-52', last(-1, 1 + m * 2 ** -52, '+') / 2 ** -52])
  return out
}

/** The rows of P's column A: the formulas of `LAST` put in a cell, then the sweeps'. */
const STORED = LAST.filter((c) => c.cell)
const SWEEPS = sweeps()

export const SHEETS: SheetCells = {
  P: {
    B1: 0.3,
    B2: 0.1,
    // a spilled array whose last operation is a difference, and the user's formula as a dynamic one
    D1: { formula: '{0.1,0.5}+0.2-0.3', dynamic: true },
    D3: { formula: '0.5-0.4-0.1', dynamic: true },
    ...Object.fromEntries([...STORED.map((c) => c.f.replace(/P!/g, '')), ...SWEEPS.map((s) => s[1])].map((f, i) => [`A${i + 1}`, { formula: xl(f) }])),
  },
}

/** The cases with their predicted answers. */
function build(): { out: Probe[]; predicted: Val[] } {
  const { out, add } = collector()
  const predicted: Val[] = []
  const put = (group: string, f: string, v: Val) => {
    add(group, 'operators', f)
    predicted.push(v)
  }
  for (const c of LAST) put('last', c.f, c.v)
  STORED.forEach((c, i) => put('stored', `P!A${i + 1}*2^60`, n((c.v as { v: number }).v * 2 ** 60)))
  put('array', 'P!D1*2^60', n(last(0.1 + 0.2, 0.3, '-') * 2 ** 60))
  put('array', 'P!E1*2^60', n(last(0.5 + 0.2, 0.3, '-') * 2 ** 60))
  put('array', 'P!D3*2^60', n(last(0.5 - 0.4, 0.1, '-') * 2 ** 60))
  SWEEPS.forEach(([group, , unit, v], i) => put(group, `P!A${STORED.length + i + 1}/${unit}`, n(v)))
  return { out, predicted }
}

export const buildProbes = (): Probe[] => build().out

/** The predicted answers by case id. */
export const PREDICTED: Readonly<Record<string, Val>> = (({ out, predicted }) => Object.fromEntries(out.map((p, i) => [p.id, predicted[i]])))(build())

/**
 * The predictions Excel did not confirm (the Mac and Windows alike): the bound is eight units in the last place of the
 * operand, 2^(e-49) for an operand in [2^e, 2^(e+1)), not 2^-49 of its magnitude, so 1.9 + 8 units - 1.9 keeps its 8
 * units at every x; and SUM(0.1,0.2,-0.3) in a cell is 0 too (`sum-zero-held` measured SUM).
 */
export const MISSED: readonly string[] = [
  'stored#29',
  'sweep:1#8',
  ...[8, 9, 10, 11, 12, 13, 14, 15].map((m) => `sweep:1.9#${m}`),
  'sweep:9.5#8',
  'sweep:9.5#9',
  ...[8, 9, 10, 11, 12, 13, 14].map((m) => `sweep:1E+15#${m}`),
  ...[8, 9, 10, 11, 12].map((m) => `sweep:1E-200#${m}`),
  'sweep:back#2',
  'sweep:back#6',
]
