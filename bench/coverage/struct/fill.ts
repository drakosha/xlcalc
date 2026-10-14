/**
 * The fill handle (Excel for Windows, `Range.AutoFill`): numbers of one cell and of several (a step, a trend, halves, going
 * down), dates by days, months and years, a time, text that ends in a number (leading zeros, two prefixes, Cyrillic),
 * names of days and months, text, a formula, a boolean, a column of numbers and text, whole sources and a part of one at
 * the far end, fills down, up, right and left, and the types series and copy. Predicted by the engine before Excel ran it
 * (`fill.predicted.json`).
 */
import type { FillType, StructureOp } from '../../../src/contract.js'
import { caseAdder, type StructCases } from './cases.js'
import type { StoredBook, StoredCell } from '../stored.js'
import { formulaCell as f } from '../stored.js'

/** A date of the 1900 system shown as one. */
const date = (serial: number): StoredCell => ({ value: serial, format: 'yyyy-mm-dd' })

/**
 * S: sources of fills down in rows 1:3 of columns A:Z, of fills up in rows 40:42, of fills right in rows 50:53 and of fills
 * left in rows 55:56.
 */
function fillBook(): StoredBook {
  const cells: Record<string, StoredCell> = {
    A1: { value: 1 },
    A2: { value: 2 },
    B1: { value: 1 },
    C1: { value: 1 },
    C2: { value: 3 },
    C3: { value: 4 },
    D1: { value: 'Item1' },
    E1: { value: 'Item1' },
    E2: { value: 'Item3' },
    F1: { value: 'Товар1' },
    // 2025-01-01; 2025-01-15 and 2025-02-15; 2025-01-31 and 2025-02-28
    G1: date(45658),
    H1: date(45672),
    H2: date(45703),
    I1: date(45688),
    I2: date(45716),
    J1: { value: 'Mon' },
    K1: { value: 'January' },
    L1: { value: 'abc' },
    M1: f('A1*2'),
    N1: { value: 1 },
    N2: { value: 2 },
    N3: { value: 'x' },
    // 9:00
    O1: { value: 0.375, format: 'h:mm' },
    P1: { value: 'Item09' },
    Q1: { value: 'Q1' },
    R1: { value: 0.5 },
    R2: { value: 1 },
    S1: { value: true },
    T1: { value: 'Пн' },
    U1: { value: '1st' },
    V1: { value: 'a1' },
    V2: { value: 'b2' },
    // 2025-01-01 and 2026-01-01
    W1: date(45658),
    W2: date(46023),
    X1: { value: 10 },
    X2: { value: 8 },
    Y1: { value: '1 item' },
    // 2025-01-01 by the short date's code
    Z1: { value: 45658, format: 'm/d/yyyy' },
    A41: { value: 1 },
    A42: { value: 2 },
    B42: { value: 'Item1' },
    C42: date(45658),
    D42: { value: 'Mon' },
    E40: { value: 1 },
    E41: { value: 2 },
    E42: { value: 'x' },
    A50: { value: 1 },
    B50: { value: 2 },
    A51: { value: 'Item1' },
    A52: { value: 'Mon' },
    B52: { value: 'Tue' },
    A53: { value: 5 },
    B53: { value: 'a' },
    Y55: { value: 10 },
    Z55: { value: 20 },
    Z56: { value: 'Mar' },
  }
  return { sheets: [{ name: 'S', cells }] }
}

/** The fill handle over `to` from `range` of the sheet S. */
export const fill = (range: string, to: string, type?: FillType): StructureOp => ({ op: 'fillRange', sheet: 'S', range, to, ...(type ? { type } : {}) })

export function buildStructCases(): StructCases {
  const { cases, add } = caseAdder('fill')
  add('down-numbers', fill('A1:A2', 'A1:A8'))
  add('down-one-number', fill('B1', 'B1:B5'))
  add('down-trend', fill('C1:C3', 'C1:C8'))
  add('down-item1', fill('D1', 'D1:D5'))
  add('down-item1-item3', fill('E1:E2', 'E1:E7'))
  add('down-cyrillic', fill('F1', 'F1:F4'))
  add('down-date', fill('G1', 'G1:G5'))
  add('down-months', fill('H1:H2', 'H1:H6'))
  add('down-month-ends', fill('I1:I2', 'I1:I6'))
  add('down-day-name', fill('J1', 'J1:J9'))
  add('down-month-name', fill('K1', 'K1:K14'))
  add('down-text', fill('L1', 'L1:L4'))
  add('down-formula', fill('M1', 'M1:M4'))
  add('down-numbers-text', fill('N1:N3', 'N1:N10'))
  add('down-numbers-text-part', fill('N1:N3', 'N1:N4'))
  add('down-time', fill('O1', 'O1:O4'))
  add('down-zeros', fill('P1', 'P1:P4'))
  add('down-quarter', fill('Q1', 'Q1:Q6'))
  add('down-halves', fill('R1:R2', 'R1:R6'))
  add('down-boolean', fill('S1', 'S1:S3'))
  add('down-day-name-ru', fill('T1', 'T1:T4'))
  add('down-ordinal', fill('U1', 'U1:U4'))
  add('down-two-prefixes', fill('V1:V2', 'V1:V6'))
  add('down-years', fill('W1:W2', 'W1:W5'))
  add('down-going-down', fill('X1:X2', 'X1:X6'))
  add('down-number-first', fill('Y1', 'Y1:Y4'))
  add('down-short-date', fill('Z1', 'Z1:Z3'))
  add('down-two-columns', fill('A1:B2', 'A1:B6'))
  add('series-one-number', fill('B1', 'B1:B5', 'series'))
  add('series-text', fill('L1', 'L1:L3', 'series'))
  add('series-trend', fill('C1:C3', 'C1:C7', 'series'))
  add('series-item1', fill('D1', 'D1:D4', 'series'))
  add('series-date', fill('G1', 'G1:G4', 'series'))
  add('copy-numbers', fill('A1:A2', 'A1:A6', 'copy'))
  add('copy-item1', fill('D1', 'D1:D4', 'copy'))
  add('copy-date', fill('G1', 'G1:G4', 'copy'))
  add('up-numbers', fill('A41:A42', 'A35:A42'))
  add('up-item1', fill('B42', 'B37:B42'))
  add('up-date', fill('C42', 'C38:C42'))
  add('up-day-name', fill('D42', 'D36:D42'))
  add('up-numbers-text', fill('E40:E42', 'E33:E42'))
  add('right-numbers', fill('A50:B50', 'A50:G50'))
  add('right-item1', fill('A51', 'A51:E51'))
  add('right-two-rows', fill('A52:B53', 'A52:F53'))
  add('left-numbers', fill('Y55:Z55', 'T55:Z55'))
  add('left-month-name', fill('Z56', 'V56:Z56'))
  return { books: { fill: fillBook() }, cases }
}
