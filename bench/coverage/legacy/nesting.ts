/**
 * How Excel 365 carries the evaluation of formulas of the old kind into nested calls (`run.ts legacy nesting`): each
 * case of `params` inside `SUMPRODUCT`, whose argument is computed as an array (does the call keep its arguments as
 * arrays there, or still take one value?), and each function's call inside `SUM`, whose argument is a reference
 * (does a call that answers an array there get `@`?).
 */
import type { Probe } from '../probes.js'
import { eachCall, probe } from './params.js'

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  eachCall((name, call, v) => {
    const tag = v ? `${v.kind}${v.at}` : 'call'
    out.push(probe(`${name}#array-${tag}`, 'array', name, `SUMPRODUCT(${call})`))
    if (!v) out.push(probe(`${name}#ref-call`, 'ref', name, `SUM(${call})`))
  })
  return out
}
