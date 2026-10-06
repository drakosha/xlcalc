/** Data sheets of the text case set: T texts, L lookup data (unsorted and mixed columns too), B a database with criteria. */

export type Cell = number | string | boolean

const column = (col: string, values: (Cell | null)[], from = 1): Record<string, Cell> =>
  Object.fromEntries(values.flatMap((v, i) => (v === null ? [] : [[`${col}${from + i}`, v]])))

const DB: Cell[][] = [
  ['Tree', 'Height', 'Age', 'Yield', 'Profit'],
  ['Apple', 18, 20, 14, 105],
  ['Pear', 12, 12, 10, 96],
  ['Cherry', 13, 14, 9, 105],
  ['Apple', 14, 15, 10, 75],
  ['Pear', 9, 8, 8, 76.8],
  ['Apple', 8, 9, 6, 45],
  ['Applesauce', 10, '12', true, 50],
  ['pineapple', 7, 5, '', 30],
]

/** Criteria blocks of sheet B by column: header, then conditions (null for a blank cell). */
const CRITERIA: Record<string, (Cell | null)[]> = {
  G: ['Tree', 'App'],
  H: ['Tree', '=Apple'],
  I: ['Tree', '=App'],
  J: ['Tree', '*apple'],
  K: ['Tree', '<>Apple'],
  L: ['Tree', '?ear'],
  M: ['Tree', 'Pear', 'Cherry'],
  R: ['Age', 12],
  S: ['Age', '12'],
  T: ['Yield', true],
  U: ['Yield', '='],
  V: ['Yield', '<>'],
  W: ['Weight', 'x'],
  X: ['Weight', null],
  Y: ['Tree'],
  Z: ['Tree', null, 'Pear'],
  AA: ['tree', 'Pear'],
  AB: ['Profit', '>=96'],
  AC: ['Tree', 'PEAR'],
  AD: [null, 'Pear'],
  AE: ['Tree', '=Cherry'],
  AF: ['Tree', '=pineapple'],
  AG: [1, null],
  AH: [1, '>3'],
  AI: ['Weight', true],
  AJ: ['Weight', '='],
  AK: [null, true],
  AL: ['Weight', '<>'],
  AM: ['Weight', 1],
  AN: ['Weight', false],
  AO: ['Weight', 'TRUE'],
  AP: ['Weight', true, 'x'],
}

function database(): Record<string, Cell> {
  const out: Record<string, Cell> = {}
  DB.forEach((row, r) => row.forEach((v, c) => v !== '' && (out[`${'ABCDE'[c]}${r + 1}`] = v)))
  for (const [col, vs] of Object.entries(CRITERIA)) Object.assign(out, column(col, vs))
  // two-column blocks: Tree and Height, Height twice
  Object.assign(out, { N1: 'Tree', O1: 'Height', N2: 'Apple', O2: '>10', P1: 'Height', Q1: 'Height', P2: '>10', Q2: '<15' })
  return out
}

export const SHEETS: Record<string, Record<string, Cell>> = {
  T: {
    ...column('A', ['-', '+']),
    ...column('B', ['b', 'c', 'a', 'ab']),
    ...column('C', ['1\t2', '1 2', ' 12\t', '日本語', 'ｱｲｳ', 'ＡＢＣ１２３', 'abc', 'x\ny', '1\n2']),
    // tildes in criteria, lookups and SEARCH
    ...column('D', ['~', '~~', 'a~c', 'a~~c', '~a', 'a~', 'a*c', 'ab', 'a~bx', 'abx', 'xa']),
  },
  L: {
    ...column('A', [0, 1, 2, 2, 2, 5, 8, 13]),
    ...column('B', ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h']),
    ...column('D', ['apple', 'banana', 'Cherry', 'date', 'fig', 'grape']),
    ...column('E', [1, 2, 3, 4, 5, 6]),
    ...column('G', [1, 3, 5, 'a', 'c', 'e', false, true]),
    ...column('I', [1, null, 3, null, 5, null, 7]),
    ...column('K', [5, 1, 9, 3, 7, 2, 8]),
    ...column('M', [3, 'x', 1, true, 5, 'a', 2]),
    ...column('O', [50, 40, 30, 20, 10, 0]),
    ...column('Q', ['Ёж', 'еж', 'Йод', 'иод']),
    ...column('S', ['a*c', 'a?c', 'a~c', 'abc']),
    // unsorted and mixed columns for the approximate (binary) search
    ...column('U', [7, 3, 9, 1, 5, 8, 2, 6, 4]),
    ...column('V', [2, 2, 1, 3, 3, 1, 2, 3]),
    ...column('W', [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 4, 4]),
    ...column('X', ['a', 5, 'c', 2, 'b', 9, 'd', 1, 'e', 7]),
    ...column('Y', [1, null, null, 4, null, 2, null, null, 9, null, 3]),
    ...column('Z', [true, 3, false, 'q', 6, 1, 'b', 10]),
    ...column('AA', [10, 9, 8, 7, 6, 5, 4, 3, 2, 1]),
    ...column('AB', [5, 5, 5, 5, 5]),
    ...column('AC', [1, 3, 3, 3, 5, 7, 7, 9]),
    ...column('AD', [9, 'x', 1, 'y', 5, 'z', 3, 'w', 7]),
    ...column('AE', [50, 40, 30, 30, 20, 'x', 10, 0]),
    ...column('AG', ['а', 'б', 'е', 'ё', 'ж', 'й', 'Я']),
    // row 20: the unsorted column U along a row, row 21 values to read
    ...Object.fromEntries([7, 3, 9, 1, 5, 8, 2, 6, 4].flatMap((v, i) => [[`${'ABCDEFGHI'[i]}20`, v], [`${'ABCDEFGHI'[i]}21`, `r${i + 1}`]])),
  },
  B: database(),
}

/** Adds cases of one function. */
export type Add = (fn: string, ...formulas: string[]) => void

/** A text literal in a formula. */
export const q = (s: string) => `"${s.replace(/"/g, '""')}"`
