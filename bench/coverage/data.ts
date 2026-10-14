/**
 * The data sheet of the measurement's synthetic book (sheet `D`): single cells for edge arguments and ranges for
 * the functions over arrays. The data are made up; there are no real data here.
 */
import { dateToSerial, formatAddress } from '../../src/model.js'
import { parseRange } from '../../src/structure/area.js'

export type ErrCode = 'NULL' | 'DIV0' | 'VALUE' | 'REF' | 'NAME' | 'NUM' | 'NA' | 'GETTING_DATA' | 'SPILL' | 'CALC' | 'ERROR'

/** A cell's value: a date is kept as a serial number and written with a date format. */
export type Cell = number | string | boolean | { error: ErrCode; formula: string } | { date: number } | { text: string; formula: string }

/** The database of Excel's function reference: trees, their height, age, yield and profit. */
export const TREES: (string | number)[][] = [
  ['Tree', 'Height', 'Age', 'Yield', 'Profit'],
  ['Apple', 18, 20, 14, 105],
  ['Pear', 12, 12, 10, 96],
  ['Cherry', 13, 14, 9, 105],
  ['Apple', 14, 15, 10, 75],
  ['Pear', 9, 8, 8, 76.8],
  ['Apple', 8, 9, 6, 45],
]

export const DATA_SHEET = 'D'

export const dt = (iso: string): number => dateToSerial(iso)

/** Named places of the sheet: the address and the contents (blank cells are not written). */
interface Region {
  ref: string
  /** Values by row (one element per row for a column). */
  rows: (Cell | null)[][]
}

const col = (...xs: (Cell | null)[]) => xs.map((x) => [x])
const dates = (...xs: string[]) => col(...xs.map((x) => ({ date: dt(x) })))

export const REGIONS: Record<string, Region> = {
  // single cells
  empty: { ref: 'A1', rows: [[null]] },
  textnum: { ref: 'A2', rows: [['2']] },
  text: { ref: 'A3', rows: [['abc']] },
  true: { ref: 'A4', rows: [[true]] },
  false: { ref: 'A5', rows: [[false]] },
  date: { ref: 'A6', rows: [[{ date: dt('2025-01-31') }]] },
  div0: { ref: 'A7', rows: [[{ error: 'DIV0', formula: '1/0' }]] },
  na: { ref: 'A8', rows: [[{ error: 'NA', formula: 'NA()' }]] },
  cyr: { ref: 'A9', rows: [['Привет Мир']] },
  dotnum: { ref: 'A10', rows: [['1.5']] },
  commanum: { ref: 'A11', rows: [['1,5']] },
  zero: { ref: 'A12', rows: [[0]] },
  emptystr: { ref: 'A13', rows: [[{ text: '', formula: '""' }]] },
  ctrl: { ref: 'A14', rows: [[{ text: '\tabc\n', formula: 'CHAR(9)&"abc"&CHAR(10)' }]] },
  datetime: { ref: 'A15', rows: [[45688.75]] },
  // numeric columns of 8 rows: interchangeable in paired functions
  nums1: { ref: 'B1:B8', rows: col(4, 8, 15, 16, 23, 42, -3.5, 10.25) },
  nums2: { ref: 'C1:C8', rows: col(2, 7, 11, 19, 21, 40, -1, 9) },
  mixed: { ref: 'D1:D8', rows: col(1, '2', 'abc', true, null, 5, false, 3.5) },
  witherr: { ref: 'E1:E8', rows: col(1, 2, { error: 'DIV0', formula: '1/0' }, 4, 5, 6, 7, 8) },
  blank: { ref: 'K1:K8', rows: col(null, null, null, null, null, null, null, null) },
  nums1short: { ref: 'B1:B7', rows: col(4, 8, 15, 16, 23, 42, -3.5) },
  pos: { ref: 'AV1:AV7', rows: col(4, 5, 8, 7, 11, 4, 3) },
  // criteria: texts, amounts, dates of 7 rows
  txt: { ref: 'F1:F7', rows: col('Аренда', 'аренда', 'Арендй', 'Арeнда*', 'x*y', 'Чай', 'Чаи') }, // the e of 'Арeнда*' is Latin
  amt: { ref: 'G1:G7', rows: col(1, 2, 4, 8, 16, 32, 64) },
  dates: { ref: 'H1:H7', rows: dates('2024-12-31', '2025-01-01', '2025-01-02', '2025-02-01', '2025-03-01', '2024-06-30', '2025-12-31') },
  // cash flows
  cf: { ref: 'I1:I5', rows: col(-10000, 3000, 4200, 6800, 1200) },
  xcf: { ref: 'J1:J5', rows: col(-10000, 2750, 4250, 3250, 2750) },
  cfmix: { ref: 'I7:I11', rows: col(-10000, '3000', 4200, null, 6800) },
  cferr: { ref: 'I13:I17', rows: col(-10000, 3000, { error: 'NA', formula: 'NA()' }, 6800, 1200) },
  xcfpos: { ref: 'J7:J11', rows: col(100, 200, 300, 400, 500) },
  xdatesmix: { ref: 'AU7:AU11', rows: dates('2008-03-01', '2008-01-01', '2008-10-30', '2009-02-15', '2009-04-01') },
  xdates: { ref: 'AU1:AU5', rows: dates('2008-01-01', '2008-03-01', '2008-10-30', '2009-02-15', '2009-04-01') },
  // the lookup table L:N, keys ascending
  tbl: {
    ref: 'L1:N6',
    rows: [
      [1, 'one', 10],
      [2, 'two', 20],
      [3, 'three', 30],
      [5, 'five', 50],
      [8, 'eight', 80],
      [13, 'thirteen', 130],
    ],
  },
  lkeys: { ref: 'L1:L6', rows: col(1, 2, 3, 5, 8, 13) },
  lnames: { ref: 'M1:M6', rows: col('one', 'two', 'three', 'five', 'eight', 'thirteen') },
  lvals: { ref: 'N1:N6', rows: col(10, 20, 30, 50, 80, 130) },
  ntbl: { ref: 'M1:N6', rows: [['one', 10], ['two', 20], ['three', 30], ['five', 50], ['eight', 80], ['thirteen', 130]] },
  // the same table by rows
  htbl: {
    ref: 'L8:Q10',
    rows: [
      [1, 2, 3, 5, 8, 13],
      ['one', 'two', 'three', 'five', 'eight', 'thirteen'],
      [10, 20, 30, 50, 80, 130],
    ],
  },
  // matrices
  m1: { ref: 'P1:Q2', rows: [[1, 2], [3, 4]] },
  m2: { ref: 'R1:S2', rows: [[5, 6], [7, 8]] },
  m23: { ref: 'P4:R5', rows: [[1, 2, 3], [4, 5, 6]] },
  // a database (the example of Excel's function reference) and criteria ranges
  db: {
    ref: 'U1:Y7',
    rows: TREES,
  },
  critApple: { ref: 'AA1:AB2', rows: [['Tree', 'Height'], ['Apple', '>10']] },
  critPear: { ref: 'AC1:AC2', rows: [['Tree'], ['Pear']] },
  critWild: { ref: 'AD1:AD2', rows: [['Tree'], ['a*']] },
  critCherry: { ref: 'AE1:AE2', rows: [['Tree'], ['Cherry']] },
  critNone: { ref: 'AF1:AF2', rows: [['Tree'], ['Plum']] },
  critAll: { ref: 'AG1:AG2', rows: [['Tree'], [null]] },
  // other ranges
  probx: { ref: 'AH1:AH4', rows: col(0, 1, 2, 3) },
  probp: { ref: 'AI1:AI4', rows: col(0.2, 0.3, 0.1, 0.4) },
  actual: { ref: 'AK1:AL3', rows: [[58, 35], [11, 25], [10, 23]] },
  actualmix: { ref: 'AK5:AL7', rows: [[58, '35'], [11, null], [true, 23]] },
  actualerr: { ref: 'AK9:AL11', rows: [[58, 35], [11, { error: 'DIV0', formula: '1/0' }], [10, 23]] },
  expected: { ref: 'AM1:AN3', rows: [[45.35, 47.65], [17.56, 18.44], [16.09, 16.91]] },
  hol: { ref: 'AP1:AP8', rows: dates('2012-11-22', '2012-12-04', '2013-01-21', '2008-11-26', '2008-12-04', '2009-01-21', '2006-01-02', '2006-01-16') },
  sched: { ref: 'AQ1:AQ3', rows: col(0.09, 0.11, 0.1) },
  coef: { ref: 'AR1:AR4', rows: col(1, -0.5, 0.0416666666666667, -0.00138888888888889) },
  uniq: { ref: 'AS1:AS8', rows: col(1, 2, 2, 'a', 'A', 3, 1, 'b') },
  bins: { ref: 'AT1:AT3', rows: col(0, 10, 20) },
  logic: { ref: 'AW1:AW4', rows: col(true, false, 'abc', null) },
  alltrue: { ref: 'AX1:AX3', rows: col(true, true, 1) },
  ly: { ref: 'AY1:AY5', rows: col(3, 5, 7, 10, 11) },
  lx: { ref: 'AZ1:AZ5', rows: col(1, 2, 3, 4, 5) },
  lnew: { ref: 'BA1:BA2', rows: col(6, 7) },
  gy: { ref: 'BB1:BB6', rows: col(33100, 47300, 69000, 102000, 150000, 220000) },
  gx: { ref: 'BC1:BC6', rows: col(11, 12, 13, 14, 15, 16) },
  gnew: { ref: 'BD1:BD2', rows: col(17, 18) },
  ties: { ref: 'BE1:BE7', rows: col(89, 88, 92, 101, 94, 97, 95) },
  dup: { ref: 'BF1:BF6', rows: col(1, 2, 2, 3, 5, 2) },
  modes: { ref: 'BG1:BG12', rows: col(1, 2, 3, 4, 3, 2, 1, 2, 3, 5, 6, 1) },
  texts3: { ref: 'BH1:BH3', rows: col('a', null, 'Б') },
}

/** The sheet's cells by address (A1 → value); blanks are left out. */
export function dataCells(): Map<string, Cell> {
  const out = new Map<string, Cell>()
  for (const [name, r] of Object.entries(REGIONS)) {
    const { r1, c1 } = parseRange(r.ref)!
    r.rows.forEach((row, i) =>
      row.forEach((v, j) => {
        if (v === null) return
        const addr = formatAddress(r1 + i, c1 + j)
        const prev = out.get(addr)
        if (prev !== undefined && JSON.stringify(prev) !== JSON.stringify(v)) throw new Error(`data overlap at ${addr} (${name})`)
        out.set(addr, v)
      }),
    )
  }
  return out
}

