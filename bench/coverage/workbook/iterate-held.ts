/**
 * Held out: iterative calculation on books the set `iterate` did not have, the answers predicted by its rules and
 * committed before Excel ran (`iterate-held.predicted.json`): a loan over periods with interest on the average balance,
 * cycles through ranges and functions, a cycle that overflows, a cycle across three sheets with dependents, settings
 * other than the defaults. The first computation on opening only: a recalculation goes on from the values Excel holds.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { CellInput, SheetInput } from '../../../src/contract.js'
import type { WorkbookCase } from './cases.js'

const cell = (v: number | string): CellInput => (typeof v === 'string' && v.startsWith('=') ? { formula: v.slice(1) } : { value: v })
const sheet = (name: string, cells: Record<string, number | string>): SheetInput => ({ name, cells: Object.fromEntries(Object.entries(cells).map(([a, v]) => [a, cell(v)])) })
const zeros = (sheets: SheetInput[]) => Object.fromEntries(sheets.flatMap((s) => Object.entries(s.cells).flatMap(([a, c]) => (c.formula === undefined ? [] : [[`${s.name}!${a}`, 0]]))))
const MIRRORS = ['zMirrorT', 'zMirrorN', 'zMirrorX'] as const

const PREDICTED = JSON.parse(readFileSync(new URL('./iterate-held.predicted.json', import.meta.url), 'utf8')) as Record<string, Record<string, Val>[]>

function held(id: string, about: string, sheets: SheetInput[], read: [string, string][], count?: number, delta?: number): WorkbookCase {
  const c: WorkbookCase = {
    id,
    about,
    book: { sheets },
    read: read.map(([s, range]) => ({ sheet: s, range })),
    file: { iterate: { ...(count !== undefined ? { count } : {}), ...(delta !== undefined ? { delta } : {}) }, cached: zeros(sheets), plain: {}, mirrors: MIRRORS },
    steps: ['read'],
  }
  return PREDICTED[id] ? { ...c, predicted: PREDICTED[id] } : c
}

/** A loan over five periods: the opening balance, a draw, interest at 12% a year on the average balance, the closing balance. */
function loan(): SheetInput {
  const cells: Record<string, number | string> = { A1: 'rate', B1: 0.12 }
  for (let p = 0; p < 5; p++) {
    const r = p + 3
    cells[`A${r}`] = p === 0 ? 1000 : `=D${r - 1}`
    cells[`B${r}`] = 200 * (p % 2)
    cells[`C${r}`] = `=$B$1/12*(A${r}+D${r})/2`
    cells[`D${r}`] = `=A${r}+B${r}+C${r}`
  }
  cells.F1 = '=SUM(C3:C7)'
  return sheet('L', cells)
}

export function buildWorkbookCases(): WorkbookCase[] {
  return [
    held('loan', 'a loan over five periods, interest on the average balance', [loan()], [['L', 'A3:D7'], ['L', 'F1']]),
    held('loan-fine', 'the same with a maximum change of 1e-9 and 500 iterations', [loan()], [['L', 'A3:D7'], ['L', 'F1']], 500, 1e-9),
    held('range', 'A1:A3 each a quarter of their sum plus its row', [sheet('S', { A1: '=SUM($A$1:$A$3)/4+1', A2: '=SUM($A$1:$A$3)/4+2', A3: '=SUM($A$1:$A$3)/4+3', B1: '=AVERAGE(A1:A3)' })], [['S', 'A1:B3']]),
    held('max', 'A1 = MAX(0,100-A1/2), B1 = MIN(A1,B1+10), 20 iterations', [sheet('S', { A1: '=MAX(0,100-A1/2)', B1: '=MIN(A1,B1+10)' })], [['S', 'A1:B1']], 20),
    held('switch', 'A1 = IF(A1<5,A1+2,A1-3), 9 iterations', [sheet('S', { A1: '=IF(A1<5,A1+2,A1-3)' })], [['S', 'A1']], 9),
    held('overflow', 'A1 = A1*1E+100+1, ten iterations', [sheet('S', { A1: '=A1*1E+100+1', B1: '=ISNUMBER(A1)' })], [['S', 'A1:B1']], 10),
    held(
      'three-sheets',
      'a cycle across Q, P and R with dependents on M and Z, seven iterations, a maximum change of 0.01',
      [
        sheet('Q', { A1: '=R!A1/3+1', B1: '=A1*2' }),
        sheet('P', { A1: '=Q!A1+0.5' }),
        sheet('R', { A1: '=P!A1*2' }),
        sheet('M', { A1: '=Q!B1+R!A1' }),
        sheet('Z', { A1: '=M!A1/2' }),
      ],
      [['Q', 'A1:B1'], ['P', 'A1'], ['R', 'A1'], ['M', 'A1'], ['Z', 'A1']],
      7,
      0.01,
    ),
    held('text', 'A1 = A1&"ab" and B1 = LEN(A1)&"-"&B1, three iterations', [sheet('S', { A1: '=A1&"ab"', B1: '=LEN(A1)&"-"&B1' })], [['S', 'A1:B1']], 3),
  ]
}
