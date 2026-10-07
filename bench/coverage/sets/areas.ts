/**
 * References of several areas: unions (`(D!B1:B3,D!C1:C3)`) and references across sheets (`Jan:Mar!A1:A3`, the
 * sheets from Jan to Mar in the book's order, Feb between them). Every whitelisted function's typical call (`legacy/calls.json`)
 * with each argument that takes a range given a union and a reference across sheets: which functions take them,
 * and what the others answer. Then the forms around them: alone, under operators, INDEX with an area number,
 * AREAS, a union of other sheets, sheets in reverse order, whole columns. Answers: `test/excel/cases/areas.json`.
 */
import type { Probe, SheetCells } from '../probes.js'
import { DEFS } from '../../../src/functions.js'
import { collector } from './arrays/common.js'
import { eachCall } from '../legacy/params.js'

export const UNION = '(D!B1:B3,D!C1:C3)'
export const CUBE = 'Jan:Mar!A1:A3'

/**
 * Three sheets before the data sheet: Jan and Mar at the ends, Feb between them. Their names do not read as cell
 * addresses: with sheets S1, S2, S3 Excel computed every unquoted `S1:S3!A1:A3` as #VALUE! (`'S1:S3'!A1:A3` worked).
 */
export const SHEETS: SheetCells = {
  Jan: { A1: 1, A2: 2, A3: 3, B1: 10, B2: 'x', B3: true },
  Feb: { A1: 4, A2: 5, A3: 6, B1: 20, B2: 40 },
  Mar: { A1: 7, A2: 8, A3: 9, B1: 30, B2: 60, B3: 90 },
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
  forms('cube', CUBE, `${CUBE}+1`, 'Jan:Mar!A1', 'Jan:Mar!A1+1', 'SUM(Jan:Mar!A1)', 'SUM(Mar:Jan!A1:A3)', 'SUM(Jan:Jan!A1:A3)', 'SUM(Feb:Mar!A1:A3)', 'SUM(Jan:Mar!A:A)', 'SUM(Jan:Mar!1:1)', "SUM('Jan:Mar'!A1:A3)", 'SUM(Jan:Mar!A1:B3)', 'COUNTA(Jan:Mar!B1:B3)')
  forms('cube', 'INDEX(Jan:Mar!A1:A3,2)', 'INDEX(Jan:Mar!A1:A3,2,1,2)', 'AREAS(Jan:Mar!A1)', 'ISREF(Jan:Mar!A1)', 'ROWS(Jan:Mar!A1:A3)', 'SUM(CHOOSE(1,Jan:Mar!A1:A3))', 'SUM(IF(TRUE,Jan:Mar!A1:A3))', 'SUM(Jan:Mar!A1:A3,10)', 'SUM(Jan:Mar!A1:A3*2)')
  forms('cube', 'COUNT((Jan:Mar!A1,D!B1))', 'SUM(Jan:Mar!B1:B3)', 'COUNTA(Jan:Mar!A1:B3)', 'VSTACK(Jan:Mar!A1:A2)', 'HSTACK(Jan:Mar!A1:A2)', 'ROWS(VSTACK(Jan:Mar!A1:A2))', 'COLUMNS(HSTACK(Jan:Mar!A1:B2))', 'SUM(VSTACK(Jan:Mar!A1:A2))', 'INDEX(VSTACK(Jan:Mar!A1:A2),4)', 'INDEX(HSTACK(Jan:Mar!A1:A2),2,3)', 'ROWS(TOCOL(Jan:Mar!A1:A2))')
  forms('union', UNION, `${UNION}+1`, `SUM(${UNION})`, 'SUM((D!B1:B3,D!B1))', 'COUNT((D!B1:B3,D!B1:B3))', 'SUM((D!B1:B3))', 'SUM((D!B1,Jan!A1))', 'SUM((D!B1:B3,D!C1:C3,D!E1:E2))', `SUM(${UNION},10)`, `SUM(${UNION}*2)`)
  forms('union', `ROWS(VSTACK(${UNION}))`, `INDEX(VSTACK(${UNION}),4)`, `COLUMNS(HSTACK(${UNION}))`, `INDEX(HSTACK(${UNION}),2,2)`, `ROWS(TOCOL(${UNION}))`, `INDEX(TOCOL(${UNION}),4)`, `ROWS(TOROW((D!B1:B2,D!C1:C3)))`, `ROWS(VSTACK((D!B1:C1,D!B2:B3)))`, `INDEX(VSTACK((D!B1:C1,D!B2:B3)),2,2)`)
  forms('union', `INDEX(${UNION},2,1,2)`, `INDEX(${UNION},2,1,3)`, `INDEX(${UNION},2,1,0)`, `INDEX(${UNION},2)`, `INDEX(${UNION},2,1)`, 'INDEX(D!B1:B3,2,1,1)', 'INDEX(D!B1:B3,2,1,2)', `SUM(INDEX(${UNION},0,0,2))`, `INDEX(${UNION},0,1,1)`, `ROWS(INDEX(${UNION},0,0,2))`)
  forms('union', `AREAS(${UNION})`, 'AREAS(D!B1:B3)', 'AREAS((D!B1,D!C1,D!E1))', `ISREF(${UNION})`, 'ISREF(D!B1:B3)', 'ISREF(1)', `SUM(CHOOSE(1,${UNION}))`, `SUM(IF(TRUE,${UNION}))`, `CHOOSE(1,${UNION})`, `ROWS(${UNION})`, `COLUMNS(${UNION})`)
  return out
}
