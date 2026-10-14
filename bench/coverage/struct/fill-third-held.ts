/**
 * The fill handle, held out after `fill-held` (Excel for Windows, `Range.AutoFill`): a number alone beside a date, text, a
 * boolean and a formula of another line; empty cells inside runs of numbers, text and names; three dates and three times
 * with one step and without; months' ends of different days, the 28th of February a year apart, the 30th over February;
 * ordinals in capitals, numbers first without a space, with zeros, inside text; names of days round the week and without
 * one step, quarters round the year; fills up, right and left; the type series over numbers alone. Predicted by the engine
 * before Excel ran it (`fill-third-held.predicted.json`).
 */
import { caseAdder, type StructCases } from './cases.js'
import type { StoredBook, StoredCell } from '../stored.js'
import { formulaCell as f } from '../stored.js'
import { fill } from './fill.js'

/** A date of the 1900 system shown as one. */
const date = (serial: number): StoredCell => ({ value: serial, format: 'yyyy-mm-dd' })

/** A time of day shown as one. */
const time = (hours: number): StoredCell => ({ value: hours / 24, format: 'h:mm' })

/** S: sources of fills down in rows 1:4, of fills up in rows 41:42, of fills right in rows 50:51 and of fills left in row 56. */
function thirdBook(): StoredBook {
  const cells: Record<string, StoredCell> = {
    A1: { value: 5 },
    // 2025-01-01
    B1: date(45658),
    C1: { value: 5 },
    D1: { value: 'x' },
    E1: { value: 5 },
    F1: f('E1*2'),
    G1: { value: 1 },
    G4: { value: 4 },
    H1: { value: 'Item1' },
    H3: { value: 'Item3' },
    I1: { value: 'Mon' },
    I3: { value: 'Wed' },
    // 2025-01-01, 2025-01-03, 2025-01-05; 2025-01-15, 2025-02-15, 2025-03-20
    J1: date(45658),
    J2: date(45660),
    J3: date(45662),
    K1: date(45672),
    K2: date(45703),
    K3: date(45736),
    L1: time(9),
    L2: time(9.5),
    L3: time(11),
    M1: time(9),
    M2: time(10),
    M3: time(11),
    // 2025-02-28 and 2025-03-31; 2025-02-28 and 2026-02-28; 2024-12-30 and 2025-01-30
    N1: date(45716),
    N2: date(45747),
    O1: date(45716),
    O2: date(46081),
    P1: date(45656),
    P2: date(45687),
    Q1: { value: '1ST' },
    R1: { value: '1abc' },
    S1: { value: '007 x' },
    T1: { value: 'a1b' },
    U1: { value: '5 apples 3' },
    V1: { value: 'Fri' },
    V2: { value: 'Sat' },
    V3: { value: 'Sun' },
    W1: { value: 'Mon' },
    W2: { value: 'Tue' },
    W3: { value: 'Thu' },
    X1: { value: 'Q3' },
    X2: { value: 'Q4' },
    Y1: { value: 'x-1' },
    Z1: { value: 'Item099' },
    A20: { value: 2 },
    A21: { value: 4 },
    A22: { value: 5 },
    A23: { value: 9 },
    B20: { value: 'Item2' },
    B21: { value: 'Item4' },
    B22: { value: 'Item6' },
    C20: { value: 45658, format: 'yyyy-mm-dd hh:mm' },
    C21: { value: 45658.5, format: 'yyyy-mm-dd hh:mm' },
    D20: { value: 3 },
    E20: { value: 4 },
    F20: { value: 5 },
    A40: { value: 1 },
    A42: { value: 3 },
    B41: { value: 'Mon' },
    B42: { value: 1 },
    A50: { value: 5 },
    A51: { value: true },
    Z56: { value: '3rd' },
  }
  return { sheets: [{ name: 'S', cells }] }
}

export function buildStructCases(): StructCases {
  const { cases, add } = caseAdder('third')
  add('down-number-beside-date', fill('A1:B1', 'A1:B4'))
  add('down-number-beside-text', fill('C1:D1', 'C1:D4'))
  add('down-number-beside-formula', fill('E1:F1', 'E1:F4'))
  add('down-numbers-two-gaps', fill('G1:G4', 'G1:G10'))
  add('down-text-gap', fill('H1:H3', 'H1:H9'))
  add('down-day-names-gap', fill('I1:I3', 'I1:I9'))
  add('down-dates-by-two', fill('J1:J3', 'J1:J6'))
  add('down-dates-no-step', fill('K1:K3', 'K1:K6'))
  add('down-times-no-step', fill('L1:L3', 'L1:L6'))
  add('down-times-by-hour', fill('M1:M3', 'M1:M5'))
  add('down-month-ends-other-days', fill('N1:N2', 'N1:N5'))
  add('down-february-years', fill('O1:O2', 'O1:O5'))
  add('down-thirtieth-over-february', fill('P1:P2', 'P1:P5'))
  add('down-ordinal-capitals', fill('Q1', 'Q1:Q4'))
  add('down-number-first-no-space', fill('R1', 'R1:R3'))
  add('down-number-first-zeros-space', fill('S1', 'S1:S3'))
  add('down-number-inside', fill('T1', 'T1:T3'))
  add('down-number-first-and-last', fill('U1', 'U1:U3'))
  add('down-day-names-round', fill('V1:V3', 'V1:V7'))
  add('down-day-names-no-step', fill('W1:W3', 'W1:W6'))
  add('down-quarters-round', fill('X1:X2', 'X1:X6'))
  add('down-text-minus', fill('Y1', 'Y1:Y3'))
  add('down-zeros-widen', fill('Z1', 'Z1:Z3'))
  add('down-number-trend-four', fill('A20:A23', 'A20:A27'))
  add('down-text-by-two', fill('B20:B22', 'B20:B25'))
  add('down-half-days', fill('C20:C21', 'C20:C24'))
  add('series-numbers-alone', fill('D20:F20', 'D20:F23', 'series'))
  add('up-numbers-gap', fill('A40:A42', 'A35:A42'))
  add('up-day-name-number', fill('B41:B42', 'B37:B42'))
  add('right-number-beside-boolean', fill('A50:A51', 'A50:D51'))
  add('left-ordinal', fill('Z56', 'U56:Z56'))
  return { books: { third: thirdBook() }, cases }
}
