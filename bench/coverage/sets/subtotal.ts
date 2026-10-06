/**
 * Excel cases for SUBTOTAL and AGGREGATE: which cells count as nested (a formula with SUBTOTAL or AGGREGATE,
 * also inside an expression; a reference to such a cell), every function and option over numbers, errors,
 * mixed values and blanks, the k argument of 14–19, argument checks. Computed as dynamic array formulas
 * (Excel 365). Answers: `docs/excel-cases/subtotal.json`.
 */
import type { Probe, SheetCells } from '../probes.js'

export const DYNAMIC = true

const AGG = '_xlfn.AGGREGATE'
const f = (formula: string) => ({ formula })

/** Sheet T: small ranges, each with one kind of nested cell; data for the statistics. */
export const SHEETS: SheetCells = {
  T: {
    A1: 1, A2: 2, A3: f('SUBTOTAL(9,A1:A2)'),
    B1: 1, B2: 2, B3: f(`${AGG}(9,4,B1:B2)`),
    C1: 1, C2: 2, C3: f('SUBTOTAL(9,C1:C2)*10'),
    D1: 1, D2: 2, D3: f('IF(TRUE,SUBTOTAL(9,D1:D2),0)'),
    E1: 1, E2: 2, E3: f('A3'),
    F1: 1, F2: 2, F3: f('SUBTOTAL(109,F1:F2)'),
    G1: 1, G2: 2, G3: f(`${AGG}(9,0,G1:G2)+0`),
    // errors, one from a nested SUBTOTAL
    H1: 2, H2: f('1/0'), H3: 3, H4: f('NA()'), H5: 7, H6: f('SUBTOTAL(9,H2)'),
    // a number, text, booleans, a blank, numeric text
    I1: 1, I2: 'x', I3: true, I5: 2, I6: '3', I7: false,
    // J1:J4 blank
    K1: f('SUBTOTAL(9,L1:L3)'), K2: f('SUBTOTAL(1,L1:L3)'),
    L1: 3, L2: 1, L3: 4,
    M1: 3, M2: 1, M3: 4, M4: 1, M5: 5, M6: 9, M7: 2, M8: 6, M9: 5, M10: 3, M11: f(`${AGG}(14,4,M1:M10,1)`),
    N1: 1, O1: 2, N2: 3, O2: 4, N3: f('SUBTOTAL(9,N1:O2)'), O3: 5,
    P1: 4, P2: 8, P3: 15, P4: 16, P5: 23, P6: 42, P7: f('SUBTOTAL(9,P1:P2)'), P8: f(`${AGG}(4,4,P1:P6)`),
  },
}

const NESTED = ['A1:A3', 'B1:B3', 'C1:C3', 'D1:D3', 'E1:E3', 'F1:F3', 'G1:G3', 'N1:O3']
const CODES = [...Array.from({ length: 11 }, (_, i) => i + 1), ...Array.from({ length: 11 }, (_, i) => i + 101)]

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const count = new Map<string, number>()
  const add = (fn: 'SUBTOTAL' | 'AGGREGATE', args: string) => {
    const k = (count.get(fn) ?? 0) + 1
    count.set(fn, k)
    out.push({ id: `${fn}#${k}`, group: fn, fn, formula: `${fn === 'AGGREGATE' ? AGG : fn}(${args})` })
  }

  // which cells are nested
  for (const r of NESTED) {
    add('SUBTOTAL', `9,T!${r}`)
    add('SUBTOTAL', `109,T!${r}`)
    for (const o of [0, 4, 6]) add('AGGREGATE', `9,${o},T!${r}`)
  }
  // every SUBTOTAL code over numbers with nested cells, errors, mixed values, blanks, only nested cells
  for (const r of ['P1:P8', 'H1:H6', 'I1:I7', 'J1:J4', 'K1:K2']) for (const c of CODES) add('SUBTOTAL', `${c},T!${r}`)
  // every AGGREGATE function and option over numbers with nested cells and over errors
  for (const r of ['P1:P8', 'H1:H6']) for (let fn = 1; fn <= 13; fn++) for (let o = 0; o <= 7; o++) add('AGGREGATE', `${fn},${o},T!${r}`)
  for (const r of ['I1:I7', 'J1:J4', 'K1:K2']) for (let fn = 1; fn <= 13; fn++) add('AGGREGATE', `${fn},${fn % 2 ? 0 : 6},T!${r}`)
  // functions with k
  const ks: Record<number, string> = { 14: '3', 15: '3', 16: '0.3', 17: '1', 18: '0.3', 19: '1' }
  for (const r of ['M1:M11', 'H1:H6']) for (let fn = 14; fn <= 19; fn++) for (const o of [0, 2, 4, 6]) add('AGGREGATE', `${fn},${o},T!${r},${ks[fn]}`)
  for (const k of ['1', '0', '10', '11', '"2"', 'TRUE', 'T!J1', '2.9', 'D!A7', '"abc"']) add('AGGREGATE', `14,6,T!M1:M10,${k}`)
  for (const [fn, k] of [[15, '10'], [16, '1.5'], [16, '-0.1'], [17, '4.5'], [17, '5'], [18, '0.05'], [18, '0.95'], [19, '4'], [19, '0.5']]) add('AGGREGATE', `${fn},6,T!M1:M10,${k}`)
  add('AGGREGATE', '14,6,T!M1:M10')
  add('AGGREGATE', '14,6,T!M1:M10,T!A1:A3')
  // several references
  add('SUBTOTAL', '9,T!A1:A3,T!P1:P8')
  add('SUBTOTAL', '3,T!I1:I7,T!H1:H6')
  add('AGGREGATE', '9,0,T!A1:A3,T!B1:B3')
  add('AGGREGATE', '4,6,T!H1:H6,T!M1:M10')
  add('AGGREGATE', '1,4,T!P1:P8,T!L1:L3')
  // whole columns, a single cell, cells of another sheet
  add('SUBTOTAL', '9,T!P:P')
  add('AGGREGATE', '9,0,T!P:P')
  add('SUBTOTAL', '9,T!P7')
  add('SUBTOTAL', '9,T!P1')
  add('SUBTOTAL', '9,D!A7')
  add('AGGREGATE', '9,6,D!A7')
  add('SUBTOTAL', '1,D!B1:B8,D!D1:D8')
  // argument checks
  for (const c of ['"9"', 'TRUE', '0', '12', '100', '112', '2.9', '-1', 'T!J1', '"abc"', '101.5', 'D!A7']) add('SUBTOTAL', `${c},T!P1:P6`)
  for (const c of ['0', '20', '19.5', 'TRUE', '"9"', 'D!A8']) add('AGGREGATE', `${c},0,T!P1:P6`)
  for (const o of ['-1', '8', '2.9', 'TRUE', '"6"', 'T!J1', 'D!A7']) add('AGGREGATE', `9,${o},T!P1:P6`)
  return out
}
