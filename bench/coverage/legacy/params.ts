/**
 * Where Excel 365 puts the implicit intersection operator `@` in formulas of the old kind (`run.ts legacy params`).
 * Each whitelisted function's call (one per function, frozen in `calls.json` from the stored cases) as it is, then
 * with one argument at a time replaced by a range (`D!B1:B3`) and by an expression over a range (`(D!B1:B3+0)`): a
 * range gets `@` where the argument takes one value, an expression gets `@` inside where the argument is not an
 * array (Excel computes it by values there). Size arguments (`isSize`) stay as they are; positions that take only a
 * reference (`isRefOnly`) get no expression, which Excel would not load. Then forms outside calls: operators, array
 * constants, whole columns and rows, two-dimensional ranges, nesting.
 */
import CALLS from './calls.json' with { type: 'json' }
import type { Probe } from '../probes.js'
import { xl } from '../sets/arrays/common.js'
import { xlpm } from '../sets/let/xlpm.js'
import { isRefOnly, isSize } from '../sets/fuzz/pools.js'

const RANGE = 'D!B1:B3'
const EXPR = `(${RANGE}+0)`

/** The top-level arguments of a call `NAME(...)`. */
function split(call: string, name: string): string[] {
  const out: string[] = []
  let depth = 0
  let cur = ''
  let quote: string | null = null
  for (const ch of call.slice(name.length + 1, -1)) {
    if (quote) {
      if (ch === quote) quote = null
    } else if (ch === '"') quote = ch
    else if (ch === '(' || ch === '{') depth++
    else if (ch === ')' || ch === '}') depth--
    else if (ch === ',' && depth === 0) {
      out.push(cur)
      cur = ''
      continue
    }
    cur += ch
  }
  return [...out, cur]
}

/** Forms outside one call: operators, constants, references of every shape, nesting. */
const FORMS = [
  RANGE, `+${RANGE}`, `-${RANGE}`, `${RANGE}%`, `${RANGE}+1`, `${RANGE}&"a"`, `${RANGE}=4`, `(${RANGE})`, `${RANGE}*D!C1:C3`,
  'D!B:B', 'D!B:B*2', 'D!1:1', 'D!B1:C3', 'D!B1:C3*2', 'D!B1:B1', 'D!B1', 'D!B1*2',
  '{1,2,3}', '{1,2,3}*2', '{1;2;3}+D!B1', `{1;2;3}+${RANGE}`, '{5}',
  `SUM(ABS(${RANGE}))`, `SUM(${RANGE}*2)`, `SUM(${RANGE},${RANGE}*2)`, `SUMPRODUCT(ABS(${RANGE}))`, `SUMPRODUCT(IF(${RANGE}>5,1,0))`, `SUM(IF(${RANGE}>5,1,0))`,
  'SUM(INDEX(D!B1:C3,0,1))', 'ABS(INDEX(D!B1:C3,0,1))', 'INDEX(D!B1:C3,0,1)+1', 'INDEX(D!B1:C3,2,2)', 'INDEX(D!B1:C3,2,2)*1',
  `SUM(ROW(${RANGE}))`, `ROW(${RANGE})+1`, `ROWS(${RANGE}*2)`, `INDEX(${RANGE}*2,2)`, `MATCH(8,${RANGE}*1,0)`,
  `IF(TRUE,${RANGE})`, `IF(TRUE,${RANGE})+0`, `SUM(IF(TRUE,${RANGE}))`, `CHOOSE(1,${RANGE},D!C1:C3)`, `SUM(CHOOSE(1,${RANGE},D!C1:C3))`,
  `IFERROR(${RANGE},0)`, `IFERROR(1/${RANGE},0)`, `SUM(IFERROR(1/${RANGE},0))`,
  `LET(x,${RANGE},x)`, `LET(x,${RANGE},x*2)`, `LET(x,${RANGE},SUM(x))`, `LET(x,${RANGE}*2,SUM(x))`,
  `SEQUENCE(3)`, `SEQUENCE(3)*2`, `SUM(SEQUENCE(3))`, `TRANSPOSE(${RANGE})`, `MMULT(${RANGE},TRANSPOSE(${RANGE}))`, `FILTER(${RANGE},${RANGE}>5)`, `XLOOKUP(8,${RANGE},D!C1:C3)`,
  `XLOOKUP(8,${RANGE},D!C1:C3)+0`, `ABS(SEQUENCE(3))`, `SEQUENCE(3)+${RANGE}`, `N(${RANGE})`, `T(D!F1:F3)`, `COLUMN(D!B1:D1)`, `ROW()`, `NOW()`, `RAND()>2`, `TODAY()>1`,
]

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const add = (id: string, group: string, fn: string, formula: string) => out.push({ id, group, fn, formula: xl(xlpm(formula)) })
  for (const [name, call] of CALLS as [string, string][]) {
    add(`${name}#call`, 'call', name, call)
    // LET's arguments are names and values: its forms are below
    const args = name === 'LET' ? [] : split(call, name)
    args.forEach((_, i) => {
      if (isSize(name, i)) return
      const with_ = (x: string) => `${name}(${args.map((a, j) => (j === i ? x : a)).join(',')})`
      add(`${name}#r${i + 1}`, 'range', name, with_(RANGE))
      if (!isRefOnly(name, i)) add(`${name}#e${i + 1}`, 'expr', name, with_(EXPR))
    })
  }
  FORMS.forEach((f, i) => add(`form#${i + 1}`, 'form', 'form', f))
  return out
}
