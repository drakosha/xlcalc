/**
 * The fill handle, held out after `fill` (Excel for Windows, `Range.AutoFill`): a number alone among text, an empty cell, a
 * formula, a boolean and another number of its line, numbers alone in their lines; quarters, ordinals and numbers before
 * text in other spellings; months over a month's end and the 29th of February, three months apart and going up; the trend
 * of text and of dates; names of days by twos and in other cases; times, dates with a time, numbers with a format; fills up,
 * right and left; the types series and copy. Predicted by the engine before Excel ran it (`fill-held.predicted.json`).
 */
import { caseAdder, type StructCases } from './cases.js'
import type { StoredBook, StoredCell } from '../stored.js'
import { formulaCell as f } from '../stored.js'
import { date, fill } from './fill.js'

/** S: sources of fills down in rows 1:3, of fills up in rows 40:42, of fills right in rows 50:53 and of fills left in row 56. */
function heldBook(): StoredBook {
  const cells: Record<string, StoredCell> = {
    A1: { value: 5 },
    A3: { value: 'a' },
    B1: { value: 'a' },
    B2: { value: 5 },
    C1: { value: 5 },
    C2: { value: true },
    D1: { value: 5 },
    D2: f('A1+1'),
    E1: { value: 1 },
    E2: { value: 'a' },
    E3: { value: 2 },
    F1: { value: 1 },
    F3: { value: 3 },
    G1: { value: 'Qtr1' },
    H1: { value: 'Quarter 3' },
    I1: { value: 'q2' },
    J1: { value: 'Q5' },
    K1: { value: '11th' },
    L1: { value: '2ND' },
    L2: { value: '4TH' },
    M1: { value: '10 kg' },
    N1: { value: '007x' },
    O1: { value: '1 item 2' },
    // 2025-01-30 and 2025-02-28; 2024-02-29 and 2025-02-28; 2025-01-15 and 2025-04-15
    P1: date(45687),
    P2: date(45716),
    Q1: date(45351),
    Q2: date(45716),
    R1: date(45672),
    R2: date(45762),
    // 2025-01-01, 2025-01-02, 2025-01-04
    S1: date(45658),
    S2: date(45659),
    S3: date(45661),
    T1: { value: 'Item1' },
    T2: { value: 'Item2' },
    T3: { value: 'Item4' },
    U1: { value: 'Mon' },
    U2: { value: 'Wed' },
    V1: { value: 'mon' },
    W1: { value: 'MONDAY' },
    X1: { value: 'Январь' },
    // 9:00 and 9:30; 2025-01-01 12:00
    Y1: { value: 0.375, format: 'h:mm' },
    Y2: { value: 9.5 / 24, format: 'h:mm' },
    Z1: { value: 45658.5, format: 'yyyy-mm-dd hh:mm' },
    A20: { value: 1.5, format: '0.00' },
    A21: { value: 2, format: '0.00' },
    B20: { value: 1.5, format: '0.00' },
    C20: { value: true },
    C21: { value: false },
    D20: { value: 0.1 },
    D21: { value: 0.2 },
    E20: { value: -1 },
    F20: { value: 7 },
    G20: { value: 8 },
    H20: f('F20'),
    H21: { value: 3 },
    // 2025-03-15 and 2025-04-15
    A41: date(45731),
    A42: date(45762),
    B42: { value: 'Item0' },
    C41: { value: 1 },
    C42: { value: 'a' },
    D42: { value: 'Sun' },
    A50: date(45658),
    A51: { value: 'Jan' },
    B51: { value: 'Mar' },
    A52: { value: 3 },
    A53: { value: 'b' },
    Z56: { value: 'Item3' },
  }
  return { sheets: [{ name: 'S', cells }] }
}

export function buildStructCases(): StructCases {
  const { cases, add } = caseAdder('held')
  add('down-number-gap-text', fill('A1:A3', 'A1:A9'))
  add('down-text-number', fill('B1:B2', 'B1:B6'))
  add('down-number-boolean', fill('C1:C2', 'C1:C6'))
  add('down-number-formula', fill('D1:D2', 'D1:D6'))
  add('down-number-text-number', fill('E1:E3', 'E1:E9'))
  add('down-number-gap-number', fill('F1:F3', 'F1:F9'))
  add('down-numbers-alone', fill('E20:G20', 'E20:G23'))
  add('down-qtr', fill('G1', 'G1:G6'))
  add('down-quarter-word', fill('H1', 'H1:H4'))
  add('down-quarter-small', fill('I1', 'I1:I5'))
  add('down-q5', fill('J1', 'J1:J3'))
  add('down-eleventh', fill('K1', 'K1:K5'))
  add('down-ordinals-by-two', fill('L1:L2', 'L1:L5'))
  add('down-number-first', fill('M1', 'M1:M3'))
  add('down-number-first-zeros', fill('N1', 'N1:N3'))
  add('down-numbers-both-ends', fill('O1', 'O1:O3'))
  add('down-month-end-30', fill('P1:P2', 'P1:P5'))
  add('down-leap-day-years', fill('Q1:Q2', 'Q1:Q6'))
  add('down-three-months', fill('R1:R2', 'R1:R5'))
  add('down-date-trend', fill('S1:S3', 'S1:S6'))
  add('down-text-trend', fill('T1:T3', 'T1:T6'))
  add('down-day-names-by-two', fill('U1:U2', 'U1:U6'))
  add('down-day-name-small', fill('V1', 'V1:V3'))
  add('down-day-name-capitals', fill('W1', 'W1:W3'))
  add('down-month-name-ru', fill('X1', 'X1:X3'))
  add('down-times', fill('Y1:Y2', 'Y1:Y5'))
  add('down-date-time', fill('Z1', 'Z1:Z3'))
  add('down-formatted-numbers', fill('A20:A21', 'A20:A24'))
  add('down-formatted-number', fill('B20', 'B20:B22'))
  add('down-booleans', fill('C20:C21', 'C20:C24'))
  add('down-tenths', fill('D20:D21', 'D20:D25'))
  add('down-formula-number', fill('H20:H21', 'H20:H25'))
  add('series-negative', fill('E20', 'E20:E23', 'series'))
  add('series-months', fill('R1:R2', 'R1:R4', 'series'))
  add('series-number-gap-text', fill('A1:A3', 'A1:A6', 'series'))
  add('copy-day-name', fill('U1', 'U1:U3', 'copy'))
  add('up-months', fill('A41:A42', 'A37:A42'))
  add('up-item0', fill('B42', 'B38:B42'))
  add('up-number-text', fill('C41:C42', 'C36:C42'))
  add('up-day-name', fill('D42', 'D39:D42'))
  add('right-date', fill('A50', 'A50:E50'))
  add('right-month-names-by-two', fill('A51:B51', 'A51:F51'))
  add('right-number-text', fill('A52:A53', 'A52:D53'))
  add('right-number-alone', fill('A52', 'A52:C52'))
  add('left-item3', fill('Z56', 'U56:Z56'))
  return { books: { held: heldBook() }, cases }
}
