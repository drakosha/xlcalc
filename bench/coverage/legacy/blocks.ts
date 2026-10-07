/**
 * Where an array evaluation stops in formulas of the old kind (`run.ts legacy blocks`): each argument that keeps `@`
 * inside SUMPRODUCT (the set `nesting`) under the other arguments that take arrays: INDEX's array, ROWS, MMULT. The
 * coverage measurement computed `INDEX(TRANSPOSE(range),1,2)` without intersection, where SUMPRODUCT intersects.
 */
import type { Probe } from '../probes.js'
import { eachCall, probe } from './params.js'
import NESTING from '../../../test/excel/legacy/nesting.json' with { type: 'json' }
import PARAMS from '../../../test/excel/legacy/params.json' with { type: 'json' }

const at = (f: string | undefined, kind: string) => (f === undefined ? false : kind === 'r' ? /@D!B1:B3/.test(f) : /\(@D!B1:B3\+0\)/.test(f))

export function buildProbes(): Probe[] {
  const top = PARAMS.formulas as Record<string, string>
  const nested = NESTING.formulas as Record<string, string>
  const out: Probe[] = []
  eachCall((name, call, v) => {
    if (!v) return
    const tag = `${v.kind}${v.at}`
    // the arguments whose @ stays inside SUMPRODUCT
    if (!at(top[`${name}#${tag}`], v.kind) || !at(nested[`${name}#array-${tag}`], v.kind)) return
    out.push(probe(`${name}#index-${tag}`, 'index', name, `INDEX(${call},1,1)`), probe(`${name}#rows-${tag}`, 'rows', name, `ROWS(${call})`), probe(`${name}#mmult-${tag}`, 'mmult', name, `MMULT(${call},1)`))
  })
  for (const f of ['INDEX(TRANSPOSE(D!P4:R5),1,2)', 'ROWS(TRANSPOSE(D!P4:R5))', 'COLUMNS(TRANSPOSE(D!B1:B8))', 'SUMPRODUCT(TRANSPOSE(D!P4:R5))', 'SUM(TRANSPOSE(D!P4:R5))', 'INDEX(IF(D!B1:B3>5,1,0),2)', 'ROWS(IF(D!B1:B3>5,1,0))', 'INDEX(D!B1:B3*2,2)', 'MATCH(8,TRANSPOSE(D!B1:B3),0)', 'LOOKUP(8,TRANSPOSE(D!B1:B3))'])
    out.push(probe(`form#${out.filter((p) => p.group === 'form').length + 1}`, 'form', 'form', f))
  return out
}
