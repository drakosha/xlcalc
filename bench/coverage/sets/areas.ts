/**
 * References of several areas: unions (`(D!B1:B3,D!C1:C3)`) and references across sheets (`S1:S3!A1:A3`, the sheets
 * from S1 to S3 in the book's order, S2 between them). Every whitelisted function's typical call (`legacy/calls.json`)
 * with each argument that takes a range given a union and a reference across sheets: which functions take them,
 * and what the others answer. Then the forms around them: alone, under operators, INDEX with an area number,
 * AREAS, a union of other sheets, sheets in reverse order, whole columns. Answers: `test/excel/cases/areas.json`.
 */
import type { Probe, SheetCells } from '../probes.js'
import { DEFS } from '../../../src/functions.js'
import { collector } from './arrays/common.js'
import { eachCall } from '../legacy/params.js'

export const UNION = '(D!B1:B3,D!C1:C3)'
export const CUBE = 'S1:S3!A1:A3'

/** Three sheets before the data sheet: S1 and S3 at the ends, S2 between them. */
export const SHEETS: SheetCells = {
  S1: { A1: 1, A2: 2, A3: 3, B1: 10, B2: 'x', B3: true },
  S2: { A1: 4, A2: 5, A3: 6, B1: 20, B2: 40 },
  S3: { A1: 7, A2: 8, A3: 9, B1: 30, B2: 60, B3: 90 },
}

/** Argument kinds that take a range (`args.ts`). */
const RANGES = 'NAYZDVGMRrpP'

const kindAt = (fn: string, i: number): string | undefined => {
  const def = DEFS[fn]
  if (!def) return undefined
  const ks = def.args.split(' ').filter(Boolean).map((t) => t[0])
  const rep = def.rep ?? 0
  const fixed = ks.length - rep
  return i < ks.length ? ks[i] : rep ? ks[fixed + ((i - fixed) % rep)] : undefined
}

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  eachCall((name, call, v) => {
    // the range variants of `legacy/params`, the range made a union or a reference across sheets
    if (v?.kind !== 'r' || !RANGES.includes(kindAt(name, v.at - 1) ?? '')) return
    add(`fn:union`, name, call.replace('D!B1:B3', UNION))
    add(`fn:cube`, name, call.replace('D!B1:B3', CUBE))
  })
  const forms = (group: string, ...fs: string[]) => add(group, group, ...fs)
  forms('cube', CUBE, `${CUBE}+1`, 'S1:S3!A1', 'S1:S3!A1+1', 'SUM(S1:S3!A1)', 'SUM(S3:S1!A1:A3)', 'SUM(S1:S1!A1:A3)', 'SUM(S2:S3!A1:A3)', 'SUM(S1:S3!A:A)', 'SUM(S1:S3!1:1)', "SUM('S1:S3'!A1:A3)", 'SUM(S1:S3!A1:B3)', 'COUNTA(S1:S3!B1:B3)')
  forms('cube', 'INDEX(S1:S3!A1:A3,2)', 'INDEX(S1:S3!A1:A3,2,1,2)', 'AREAS(S1:S3!A1)', 'ISREF(S1:S3!A1)', 'ROWS(S1:S3!A1:A3)', 'SUM(CHOOSE(1,S1:S3!A1:A3))', 'SUM(IF(TRUE,S1:S3!A1:A3))', 'SUM(S1:S3!A1:A3,10)', 'SUM(S1:S3!A1:A3*2)')
  forms('union', UNION, `${UNION}+1`, `SUM(${UNION})`, 'SUM((D!B1:B3,D!B1))', 'COUNT((D!B1:B3,D!B1:B3))', 'SUM((D!B1:B3))', 'SUM((D!B1,S1!A1))', 'SUM((D!B1:B3,D!C1:C3,D!E1:E2))', `SUM(${UNION},10)`, `SUM(${UNION}*2)`)
  forms('union', `INDEX(${UNION},2,1,2)`, `INDEX(${UNION},2,1,3)`, `INDEX(${UNION},2,1,0)`, `INDEX(${UNION},2)`, `INDEX(${UNION},2,1)`, 'INDEX(D!B1:B3,2,1,1)', 'INDEX(D!B1:B3,2,1,2)', `SUM(INDEX(${UNION},0,0,2))`, `INDEX(${UNION},0,1,1)`, `ROWS(INDEX(${UNION},0,0,2))`)
  forms('union', `AREAS(${UNION})`, 'AREAS(D!B1:B3)', 'AREAS((D!B1,D!C1,D!E1))', `ISREF(${UNION})`, 'ISREF(D!B1:B3)', 'ISREF(1)', `SUM(CHOOSE(1,${UNION}))`, `SUM(IF(TRUE,${UNION}))`, `CHOOSE(1,${UNION})`, `ROWS(${UNION})`, `COLUMNS(${UNION})`)
  return out
}
