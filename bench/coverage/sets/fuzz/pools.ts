/**
 * What the random formulas are made of: literals, cells and ranges of the measurement's data sheet D
 * (`bench/coverage/data.ts`), small array constants, and the mix of them each argument kind of the engine
 * (`src/args.ts`) draws from. No whole columns: a formula stays small for Excel.
 */

/** A pool: its values. */
type Pool = readonly string[]

export const POOLS = {
  num: ['0', '1', '2', '3', '4', '-1', '0.5', '2.5', '-1.5', '10', '100', '0.25', '1E+3', '7'],
  textNum: ['"3"', '"1.5"', '" 2 "', '"-4"', '"1E2"', '"50%"', '"$1,000"', '"0"'],
  dateText: ['"2025-01-31"', '"1/2/2025"', '"31-Jan-2025"', '"18:30"', '"6:35 AM"', '"2024-02-29"', '"2025-01-31 12:00"'],
  text: ['"abc"', '""', '"Привет"', '"A"', '"a*"', '"?"', '"TRUE"', '"x y"', '"aBc"', '"FALSE"', '"a,b"', '" "'],
  bool: ['TRUE', 'FALSE'],
  err: ['#N/A', '#DIV/0!', '#VALUE!', '#NUM!', '#REF!', '#NAME?'],
  // D!B1 4, B3 15, B7 -3.5, C2 7, G4 8, L4 5, A12 0, A15 45688.75 (date and time), N2 20
  cellNum: ['D!B1', 'D!B3', 'D!B7', 'D!C2', 'D!G4', 'D!L4', 'D!A12', 'D!A15', 'D!N2'],
  // blank, "2", "abc", TRUE, FALSE, #DIV/0!, #N/A, Cyrillic, "1.5", "", "Аренда", "two"
  cellAny: ['D!A1', 'D!A2', 'D!A3', 'D!A4', 'D!A5', 'D!A7', 'D!A8', 'D!A9', 'D!A10', 'D!A13', 'D!F1', 'D!M2'],
  date: ['45658', '45351', '60', '59', '61', '0', '-1', '45688.75', '40000'],
  dateCell: ['D!A6', 'D!H1', 'D!H3', 'D!A15', 'D!AU2', 'D!AP1'],
  time: ['0.5', '0.75', '1.25', '-0.5', '0', '0.999', '"12:00"', '"25:00"'],
  range: ['D!B1:B8', 'D!C1:C8', 'D!G1:G7', 'D!B1:B3', 'D!AY1:AY5', 'D!AZ1:AZ5', 'D!I1:I5', 'D!L1:L6', 'D!N1:N6', 'D!AV1:AV7', 'D!D1:D8', 'D!E1:E8', 'D!K1:K8', 'D!A1:A5', 'D!F1:F7', 'D!M1:M6', 'D!AS1:AS8', 'D!H1:H7', 'D!AW1:AW4', 'D!B2:B3'],
  grid: ['D!P1:Q2', 'D!R1:S2', 'D!P4:R5', 'D!L1:N6', 'D!B1:C8', 'D!L8:Q10', 'D!AK1:AL3', 'D!M1:N6', 'D!U1:Y7', 'D!AA1:AB2'],
  array: ['{1,2,3}', '{1;2;3}', '{1,2;3,4}', '{2,"a";TRUE,#N/A}', '{"a","b"}', '{0.5;-1}', '{1}', '{45658,45689}', '{"x";""}', '{3,1,2}'],
  dateRange: ['D!AP1:AP8', 'D!H1:H7', 'D!AU1:AU5', 'D!AU7:AU11'],
  /** Sizes and counts of arrays, repeats and orders: small numeric literals only (see `FUZZ_SIZE`). */
  size: ['-1', '0', '1', '1', '2', '2', '3', '5'],
} as const satisfies Record<string, Pool>

export type PoolName = keyof typeof POOLS

/** A mix of pools with weights. */
export type Mix = readonly (readonly [PoolName, number])[]

const SCALAR: Mix = [['num', 3], ['cellNum', 2], ['textNum', 1], ['bool', 1], ['cellAny', 2], ['text', 2], ['err', 1], ['date', 1], ['dateText', 1], ['array', 0.5], ['range', 0.5]]
const NUMBER: Mix = [['num', 9], ['cellNum', 3], ['textNum', 1.5], ['bool', 1], ['cellAny', 1.5], ['text', 0.8], ['err', 0.8], ['date', 0.8], ['dateText', 0.6], ['array', 0.6], ['range', 0.4]]
const LIST: Mix = [['range', 9], ['num', 3], ['array', 2.5], ['cellNum', 1.5], ['cellAny', 1.5], ['textNum', 0.8], ['bool', 0.8], ['err', 0.4], ['text', 0.4]]

/** The mix an argument of each engine kind draws from. */
export const KIND_MIX: Readonly<Record<string, Mix>> = {
  n: NUMBER,
  i: NUMBER,
  x: NUMBER,
  d: [['date', 7], ['dateCell', 4], ['dateText', 3], ['num', 2], ['text', 0.8], ['bool', 0.6], ['err', 0.8], ['cellAny', 1], ['array', 0.4], ['range', 0.4]],
  t: [['time', 7], ['dateText', 4], ['dateCell', 3], ['num', 2], ['text', 1], ['bool', 0.6], ['err', 0.6], ['array', 0.5]],
  s: [['text', 9], ['textNum', 2], ['dateText', 1], ['num', 2], ['cellAny', 2.5], ['bool', 1], ['err', 0.8], ['array', 0.8], ['range', 0.8]],
  b: [['bool', 10], ['num', 3], ['text', 1.5], ['cellAny', 2], ['textNum', 1], ['err', 0.8], ['array', 1], ['range', 0.6]],
  v: SCALAR,
  a: SCALAR,
  p: [...SCALAR, ['range', 2], ['grid', 0.5]],
  N: LIST,
  A: LIST,
  Y: LIST,
  Z: LIST,
  V: LIST,
  G: [['grid', 7], ['range', 7], ['array', 4], ['num', 1], ['cellAny', 1]],
  M: [['grid', 7], ['range', 7], ['array', 4], ['num', 1], ['cellAny', 1]],
  R: [['range', 9], ['grid', 5], ['array', 2], ['cellNum', 2], ['cellAny', 2]],
  D: [['dateRange', 8], ['range', 3], ['array', 2], ['date', 4], ['dateText', 2], ['err', 1]],
  r: [['size', 7], ['cellNum', 2], ['cellAny', 1]],
}

/** The categories of Excel's function list a nested call prefers for an argument of each kind (else any function). */
export const KIND_CATEGORY: Readonly<Record<string, readonly string[]>> = {
  n: ['Math and trigonometry', 'Statistical', 'Financial', 'Date and time', 'Compatibility'],
  i: ['Math and trigonometry', 'Statistical', 'Date and time'],
  x: ['Math and trigonometry', 'Engineering', 'Financial'],
  d: ['Date and time'],
  t: ['Date and time'],
  s: ['Text'],
  b: ['Logical', 'Information'],
}

/** Typical calls of the functions that have none in the measurement or in the array set's `BASE`. */
export const FUZZ_BASE: Readonly<Record<string, string>> = {
  FILTER: 'FILTER(D!B1:B8,D!B1:B8>10)',
  MDETERM: 'MDETERM(D!P1:Q2)',
  MINVERSE: 'MINVERSE(D!P1:Q2)',
  SEQUENCE: 'SEQUENCE(2,3)',
  SORTBY: 'SORTBY(D!M1:M6,D!N1:N6,-1)',
  TEXTSPLIT: 'TEXTSPLIT("a,b;c",",",";")',
  TOCOL: 'TOCOL(D!P4:R5)',
  TOROW: 'TOROW(D!P4:R5)',
  WRAPCOLS: 'WRAPCOLS(D!B1:B8,3)',
  WRAPROWS: 'WRAPROWS(D!B1:B8,3)',
}

/**
 * Argument positions where only small numeric literals go (`POOLS.size`): sizes of arrays, repeats, orders
 * and indexes of the functions that build arrays (a date or a reference there can exhaust Excel, see
 * `bench/coverage/README.md`). `rest` covers every position from that one on.
 */
export const FUZZ_SIZE: Readonly<Record<string, { at?: readonly number[]; rest?: number }>> = {
  MUNIT: { at: [0] },
  SEQUENCE: { at: [0, 1] },
  EXPAND: { at: [1, 2] },
  TAKE: { at: [1, 2] },
  DROP: { at: [1, 2] },
  WRAPROWS: { at: [1] },
  WRAPCOLS: { at: [1] },
  REPT: { at: [1] },
  BESSELI: { at: [1] },
  BESSELJ: { at: [1] },
  BESSELK: { at: [1] },
  BESSELY: { at: [1] },
  CHOOSECOLS: { rest: 1 },
  CHOOSEROWS: { rest: 1 },
}

export const isSize = (fn: string, i: number): boolean => {
  const s = FUZZ_SIZE[fn]
  return !!s && ((s.at?.includes(i) ?? false) || (s.rest !== undefined && i >= s.rest))
}

/**
 * Argument positions Excel takes only as a reference (a formula with an array or a value there does not
 * load): the ranges of the criteria functions, ROW and COLUMN, COUNTBLANK, RANK's list, SUBTOTAL's and
 * AGGREGATE's references. `odd`/`even` cover those positions from `from` on.
 */
export const REF_ONLY: Readonly<Record<string, { at?: readonly number[]; from?: number; odd?: boolean; even?: boolean }>> = {
  COUNTIF: { at: [0] },
  COUNTIFS: { from: 0, even: true },
  SUMIF: { at: [0, 2] },
  AVERAGEIF: { at: [0, 2] },
  SUMIFS: { at: [0], from: 1, odd: true },
  AVERAGEIFS: { at: [0], from: 1, odd: true },
  MAXIFS: { at: [0], from: 1, odd: true },
  MINIFS: { at: [0], from: 1, odd: true },
  ROW: { at: [0] },
  COLUMN: { at: [0] },
  COUNTBLANK: { at: [0] },
  RANK: { at: [1] },
  'RANK.EQ': { at: [1] },
  'RANK.AVG': { at: [1] },
  SUBTOTAL: { from: 1 },
  AGGREGATE: { from: 2 },
}

export function isRefOnly(fn: string, i: number): boolean {
  const r = REF_ONLY[fn]
  if (!r) return false
  if (r.at?.includes(i)) return true
  if (r.from === undefined || i < r.from) return false
  return r.odd ? i % 2 === 1 : r.even ? i % 2 === 0 : true
}
