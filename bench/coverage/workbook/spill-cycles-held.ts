/**
 * Held out: arrays in a cycle through their own spill, with iterative calculation on, and another array that wants the
 * same cells. The formulas typed one by one as Excel 365 enters them, in different orders, then read after a full
 * calculation and after one that rebuilds the dependencies; and the same book with the formulas in its file. The books:
 * `B3 = SORT(B1:C2)`, whose spill reaches C4, read by C2 in its range, beside `A4 = SEQUENCE(1,4)` over the same row
 * (`sort`); `D11 = SWITCH(E11,…,H1:I4,…)`, whose spill reaches E11, beside `C14 = MUNIT(3)` (`munit`); two sheets each
 * spilling into the other's range (`sheets`); an array waiting for the spill of an array that reads it, read by a cycle
 * (`wait`). Predicted by the engine before Excel ran it (`spill-cycles-held.predicted.json`; a case the engine does not
 * compute is predicted as its error).
 */
import { existsSync, readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { RangeRef, SheetInput } from '../../../src/contract.js'
import type { Step, WorkbookCase } from './cases.js'

const FILE = new URL('./spill-cycles-held.predicted.json', import.meta.url)
const PREDICTED = (existsSync(FILE) ? JSON.parse(readFileSync(FILE, 'utf8')) : {}) as Record<string, Record<string, Val>[]>

/** Mirror sheets named to sort after the cycle's sheets, iterative calculation with Excel's defaults. */
const FILE_PARTS = { iterate: {}, mirrors: ['zMirrorT', 'zMirrorN', 'zMirrorX'] as const }

/** The formulas of a book by sheet and cell. */
type Formulas = Record<string, Record<string, string>>

const BOOKS: Record<string, { formulas: Formulas; read: RangeRef[] }> = {
  sort: { formulas: { S: { B3: 'SORT(B1:C2)', C2: 'C4+1', A4: 'SEQUENCE(1,4)' } }, read: [{ sheet: 'S', range: 'A1:D5' }] },
  munit: { formulas: { S: { C14: 'MUNIT(3)', D11: 'SWITCH(E11,30,1,H1:I4,2)' } }, read: [{ sheet: 'S', range: 'C11:E16' }] },
  sheets: {
    formulas: { Sheet2: { E5: "'Q data'!E5:F8*2", D7: 'TRANSPOSE(D4:D6)' }, 'Q data': { F5: '#DIV/0!', D4: 'LET(x,0.1,Sheet2!E8:F10)' } },
    read: [{ sheet: 'Sheet2', range: 'D4:H10' }, { sheet: 'Q data', range: 'D4:G10' }],
  },
  wait: { formulas: { S: { B1: 'SEQUENCE(2)', A2: '{1,2}+B2', C1: 'B1+C2', C2: 'C1*1' } }, read: [{ sheet: 'S', range: 'A1:C3' }] },
}

/** Read after typing (Excel computes as it goes), after a full calculation, after one that rebuilds the dependencies. */
const AFTER: Step[] = ['read', 'full', 'read', 'rebuild', 'read']

/** The book's sheets in `order` (all of them by default), empty or with their formulas. */
const sheetsOf = (formulas: Formulas, filled: boolean, order = Object.keys(formulas)): SheetInput[] =>
  order.map((name) => ({ name, cells: filled ? Object.fromEntries(Object.entries(formulas[name]).map(([a, formula]) => [a, { formula }])) : {} }))

function held(c: Omit<WorkbookCase, 'file'>): WorkbookCase {
  return { ...c, file: FILE_PARTS, ...(PREDICTED[c.id] ? { predicted: PREDICTED[c.id] } : {}) }
}

/** The formulas typed in `order` (`Sheet!A1`, or `A1` on S) into the empty sheets. */
function typed(book: string, order: string[], sheets?: string[]): WorkbookCase {
  const { formulas, read } = BOOKS[book]
  const steps: Step[] = order.map((k) => {
    const [sheet, cell] = k.includes('!') ? k.split('!') : ['S', k]
    return { formula: { sheet, cell, formula: formulas[sheet][cell], dynamic: true } }
  })
  const id = `${book}-typed-${order.join('-').replaceAll('!', '.').replaceAll(' ', '_')}${sheets ? `-${sheets.join('-').replaceAll(' ', '_')}` : ''}`
  return held({ id, about: `${book}: the formulas typed in the order ${order.join(', ')}`, book: { sheets: sheetsOf(formulas, false, sheets) }, read, steps: [...steps, ...AFTER] })
}

/** The formulas in the book's file, its sheets in `order`. */
function filed(book: string, order?: string[]): WorkbookCase {
  const { formulas, read } = BOOKS[book]
  return held({ id: `${book}-file${order ? `-${order.join('-').replaceAll(' ', '_')}` : ''}`, about: `${book}: the formulas in the file${order ? `, sheets ${order.join(', ')}` : ''}`, book: { sheets: sheetsOf(formulas, true, order) }, read, steps: AFTER })
}

export function buildWorkbookCases(): WorkbookCase[] {
  const sorts = [['B3', 'C2', 'A4'], ['B3', 'A4', 'C2'], ['C2', 'B3', 'A4'], ['C2', 'A4', 'B3'], ['A4', 'B3', 'C2'], ['A4', 'C2', 'B3']]
  return [
    filed('sort'),
    ...sorts.map((o) => typed('sort', o)),
    filed('munit'),
    typed('munit', ['C14', 'D11']),
    typed('munit', ['D11', 'C14']),
    filed('sheets', ['Sheet2', 'Q data']),
    filed('sheets', ['Q data', 'Sheet2']),
    typed('sheets', ['Sheet2!E5', 'Sheet2!D7', 'Q data!F5', 'Q data!D4']),
    typed('sheets', ['Q data!F5', 'Q data!D4', 'Sheet2!E5', 'Sheet2!D7']),
    filed('wait'),
    typed('wait', ['B1', 'A2', 'C1', 'C2']),
    typed('wait', ['C2', 'C1', 'A2', 'B1']),
  ]
}
