/**
 * The database functions select records through the criteria selection SUMIFS uses. What they cost when they select
 * again: a database of 3,400 records (a direction, an item, an amount) under 50 DSUM, each with its own criteria block of
 * a direction and an item; after the book opened and one edit of an amount, edits of the amount check no record one by
 * one and select again in no formula. Their answers, as the functions' own pass over the records gave them, recorded
 * before they went through the selection: an empty criteria row, a condition under a header that names no field,
 * records past the occupied part, fields by number and by name in any case, alternative rows, the cases of DCOUNT,
 * DCOUNTA and DGET; edits of the database, the criteria and the headers in a session; a date without a year after the
 * clock moves.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, type BookInput, type CellInput, type CellValue, type Scalar } from '../src/index.js'
import { watchSelections } from '../src/criteria.js'
import { watchRecords } from '../src/functions/database.js'
import { datedRows, followsClockYear } from './clock-years.js'
import { err } from './formula-values.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const DIRS = ['Доход', 'Расход']
const ITEMS = Array.from({ length: 200 }, (_, i) => `Товар ${i}`)

const colName = (c: number) => (c < 26 ? String.fromCharCode(65 + c) : String.fromCharCode(64 + Math.floor(c / 26)) + String.fromCharCode(65 + (c % 26)))
/** A cell's input: a value, or `{ f }` a formula. */
type Given = Scalar | { f: string }
const input = (v: Given): CellInput => (v !== null && typeof v === 'object' ? { formula: v.f } : { value: v })

/** Cells of `rows` from column `c` and row `r` down; a blank (null) is left out. */
function block(rows: Given[][], c: number, r: number): Record<string, CellInput> {
  const cells: Record<string, CellInput> = {}
  rows.forEach((row, i) => row.forEach((v, j) => v !== null && (cells[`${colName(c + j)}${r + i}`] = input(v))))
  return cells
}

/** Sheet Data: a header row and 3,400 records; sheet Out: 50 DSUM of an item's income, each over its own criteria block. */
function sums(): BookInput {
  const data = block([['Направление', 'Товар', 'Сумма']], 0, 1)
  for (let r = 2; r <= 3401; r++) Object.assign(data, block([[DIRS[r % 3 === 0 ? 1 : 0], ITEMS[r % ITEMS.length], r % 97]], 0, r))
  const out: Record<string, CellInput> = {}
  for (let k = 0; k < 50; k++) {
    const row = 1 + k * 3
    Object.assign(out, block([['Направление', 'Товар'], [DIRS[0], `=${ITEMS[k]}`]], 0, row))
    out[`D${row}`] = { formula: `DSUM(Data!$A$1:$C$3401,"Сумма",A${row}:B${row + 1})` }
  }
  return { sheets: [{ name: 'Data', cells: data }, { name: 'Out', cells: out }] }
}

/** The database A1:E10: headers a text, a number, TRUE and the text `true`; blanks, text and errors among the records. */
const DB = block(
  [
    ['Fruit', 'Amount', 45688, true, 'true'],
    ['Apple', 10, 0.1, 'x', 5],
    ['Pear', 20, 0.7, 'y', 6],
    ['apple pie', 30, 1.3, 'x', 7],
    ['Banana', null, 0.2, 'z', 8],
    [null, 50, 2.9, 'x', 9],
    ['Apricot', 'n/a', 0.05, 'y', 10],
    ['Pear', 70, 3.3, 'x', 11],
    [{ f: '1/0' }, 80, 0.6, 'y', 12],
    ['Pear', { f: 'NA()' }, 1.1, 'x', 13],
  ],
  0,
  1,
)

/**
 * Criteria blocks, one under another from H with a blank row between, and the formulas over each (from N in its first
 * row; `CR` the block, `HR` its header row) with their answers. A whole column A:E reads down to the last row of the sheet.
 */
const CASES: [Given[][], Record<string, CellValue>][] = [
  [[['Fruit'], [null]], { 'DCOUNTA(A1:E10,,CR)': 9, 'DCOUNT(A1:E10,,CR)': 9, 'DCOUNT(A1:E10,"Amount",CR)': 6, 'DSUM(A1:E9,"Amount",CR)': 260, 'DCOUNTA(A1:E30,,CR)': 29, 'DCOUNTA(A:E,,CR)': 117, 'DSUM(A1:E10,"45688",CR)': 10.25 }],
  [[['Nope'], ['x']], { 'DCOUNTA(A1:E10,,CR)': 0, 'DSUM(A1:E9,2,CR)': 0 }],
  [[['Nope'], [true]], { 'DCOUNTA(A1:E10,,CR)': 9, 'DSUM(A1:E9,2,CR)': 260, 'DCOUNTA(A:E,,CR)': 117 }],
  [[['Nope'], ['TRUE']], { 'DCOUNTA(A1:E10,,CR)': 9 }],
  [[['Nope'], [false]], { 'DCOUNTA(A1:E10,,CR)': 0 }],
  [[['Nope', 'Fruit'], [true, 'Pear']], { 'DCOUNTA(A1:E10,,CR)': 3, 'DSUM(A1:E10,3,CR)': 5.1 }],
  [[['Nope', 'Fruit'], [1, 'Pear']], { 'DCOUNTA(A1:E10,,CR)': 3 }],
  [[['Nope', 'Fruit'], ['x', 'Pear'], [null, 'Apple']], { 'DCOUNTA(A1:E10,,CR)': 2, 'DSUM(A1:E10,3,CR)': 1.4 }],
  [[['Amount'], [{ f: '""' }]], { 'DCOUNTA(A1:E10,,CR)': 1, 'DCOUNTA(A1:E30,,CR)': 21, 'DCOUNTA(A:E,,CR)': 109, 'DCOUNTA(A:E,1,CR)': 1, 'DSUM(A:E,3,CR)': 0.2 }],
  [[['Amount'], ['<>x']], { 'DCOUNTA(A1:E10,,CR)': 9, 'DCOUNTA(A1:E30,,CR)': 29, 'DCOUNTA(A:E,,CR)': 117, 'DCOUNTA(A:E,1,CR)': 8 }],
  [[['Fruit'], ['=']], { 'DCOUNTA(A1:E10,,CR)': 1, 'DCOUNTA(A1:E30,,CR)': 21, 'DCOUNTA(A:E,,CR)': 109, 'DSUM(A:E,3,CR)': 2.9 }],
  [[['Fruit'], ['<>']], { 'DCOUNTA(A1:E10,,CR)': 8, 'DCOUNTA(A1:E30,,CR)': 8, 'DCOUNTA(A:E,,CR)': 8 }],
  [[['Amount', 'Fruit'], ['<>x', null], [null, 'Pear']], { 'DCOUNTA(A1:E30,,CR)': 29, 'DCOUNTA(A:E,,CR)': 117, 'DSUM(A:E,3,CR)': 10.25 }],
  [[['Fruit'], ['=Banana']], { 'DGET(A1:E10,"Fruit",CR)': "Banana", 'DGET(A1:E10,"Amount",CR)': err('VALUE'), 'DGET(A1:E10,4,CR)': "z", 'DCOUNTA(A1:E10,"Amount",CR)': 0 }],
  [[['Fruit'], ['Pear']], { 'DGET(A1:E10,"Fruit",CR)': err('NUM'), 'DCOUNT(A1:E10,,CR)': 3, 'DCOUNTA(A1:E10,,CR)': 3, 'DCOUNTA(A1:E10,"Amount",CR)': 3, 'DCOUNT(A1:E10,"Amount",CR)': 2, 'DSUM(A1:E10,"Amount",CR)': err('NA') }],
  [[['Fruit'], ['Kiwi']], { 'DGET(A1:E10,"Fruit",CR)': err('VALUE'), 'DSUM(A1:E10,"Amount",CR)': 0, 'DAVERAGE(A1:E10,"Amount",CR)': err('DIV0'), 'DMAX(A1:E10,"Amount",CR)': 0, 'DPRODUCT(A1:E10,"Amount",CR)': 0, 'DSTDEV(A1:E10,"Amount",CR)': err('DIV0'), 'DCOUNT(A1:E10,,CR)': 0 }],
  [[['Fruit'], ['App']], { 'DSUM(A1:E10,"amount",CR)': 40, 'DSUM(A1:E10,"AMOUNT",CR)': 40, 'DSUM(A1:E10,2,CR)': 40, 'DSUM(A1:E10,2.9999999,CR)': 1.4, 'DSUM(A1:E10,0,CR)': err('VALUE'), 'DSUM(A1:E10,6,CR)': err('VALUE'), 'DSUM(A1:E10,"45688",CR)': 1.4, 'DSUM(A1:E10,45688,CR)': err('VALUE'), 'DCOUNTA(A1:E10,"true",CR)': 2, 'DCOUNTA(A1:E10,TRUE,CR)': 2, 'DSUM(A1:E10,5,CR)': 12, 'DSUM(A1:E10,B1,CR)': 40, 'DSUM(A1:E10,A9,CR)': err('VALUE'), 'DSUM(A1:E10,B1:B2,CR)': err('VALUE'), 'DSUM(A1:E10,NA(),CR)': err('NA'), 'DSUM(A1:E10,"Nope",CR)': err('VALUE') }],
  [[[45688], ['>1']], { 'DCOUNTA(A1:E10,,CR)': 4, 'DSUM(A1:E10,3,CR)': 8.6 }],
  [[['45688'], ['<1']], { 'DCOUNTA(A1:E10,,CR)': 5 }],
  [[['true'], ['x']], { 'DCOUNTA(A1:E10,,CR)': 5 }],
  [[[true], ['y']], { 'DCOUNTA(A1:E10,,CR)': 3 }],
  [[['FRUIT'], ['pear']], { 'DCOUNTA(A1:E10,,CR)': 3 }],
  [[['Fruit', 'Amount'], ['Pear', null], [null, '>15']], { 'DSUM(A1:E10,"45688",CR)': 9.9, 'DPRODUCT(A1:E10,"45688",CR)': 5.747742, 'DSTDEV(A1:E10,"45688",CR)': 1.15887876846545, 'DSTDEVP(A1:E10,"45688",CR)': 1.05790673817056, 'DVAR(A1:E10,"45688",CR)': 1.343, 'DVARP(A1:E10,"45688",CR)': 1.11916666666667, 'DAVERAGE(A1:E10,"45688",CR)': 1.65, 'DMAX(A1:E10,"45688",CR)': 3.3, 'DMIN(A1:E10,"45688",CR)': 0.6, 'DSUM(A1:E10,"Amount",CR)': err('NA'), 'DCOUNT(A1:E10,,CR)': 6 }],
  [[['Fruit'], ['Pear'], ['Pear']], { 'DCOUNT(A1:E10,,CR)': 3, 'DSUM(A1:E10,3,CR)': 5.1 }],
  [[['Fruit'], ['Pear'], [null]], { 'DCOUNTA(A1:E10,,CR)': 9, 'DCOUNTA(A:E,,CR)': 117 }],
  [[['Amount', 'Amount'], ['>15', '<75']], { 'DSUM(A1:E10,"Amount",CR)': 170, 'DCOUNT(A1:E10,,CR)': 4 }],
  [[['Fruit'], [{ f: '1/0' }]], { 'DCOUNTA(A1:E10,,CR)': 1, 'DGET(A1:E10,2,CR)': 80 }],
  [[['Fruit'], ['#DIV/0!']], { 'DCOUNTA(A1:E10,,CR)': 1 }],
  [[['Amount'], ['#N/A']], { 'DCOUNTA(A1:E10,,CR)': 1, 'DGET(A1:E10,1,CR)': "Pear" }],
  [[['Fruit'], ['?ear']], { 'DCOUNTA(A1:E10,,CR)': 3 }],
  [[['Fruit'], ['*an*']], { 'DGET(A1:E10,1,CR)': "Banana" }],
  [[['Fruit'], ['<>Pear']], { 'DCOUNTA(A1:E10,,CR)': 6, 'DCOUNTA(A:E,,CR)': 114 }],
  [[['Fruit'], ['>b']], { 'DCOUNTA(A1:E10,,CR)': 4 }],
  [[['Amount'], [20]], { 'DCOUNTA(A1:E10,,CR)': 1 }],
  [[['Amount'], ['>=50']], { 'DSUM(A1:E10,3,CR)': 6.8 }],
  [[['Amount'], ['n']], { 'DCOUNTA(A1:E10,,CR)': 1 }],
  [[['Amount'], ['2']], { 'DCOUNTA(A1:E10,,CR)': 0 }],
  [[['Fruit'], ['Pear']], { 'DSUM(A1:E10,"Amount",HR)': err('VALUE'), 'DSUM(A1:E1,"Amount",CR)': err('VALUE'), 'DSUM((A1:E5,A6:E10),"Amount",CR)': err('VALUE'), 'DSUM(\'S:T\'!A1:E10,"Amount",CR)': err('VALUE'), 'DSUM(NA(),"Amount",1/0)': err('NA'), 'DSUM(A1:E10,"Amount",1/0)': err('DIV0'), 'DSUM({"Fruit","Amount";"Pear",5;"Apple",6;"Pear",7},"Amount",CR)': 12, 'DCOUNTA(T!A1:E10,,CR)': 3, 'DSUM(T!A:E,3,CR)': 5.1, 'DSUM(IF(TRUE,A1:E10),3,CR)': 5.1 }],
]

const FRUITS = ['Apple', 'Pear', 'apple pie', 'Banana', 'Apricot', 'Kiwi', 'Plum']

/** Sheet T: 300 records of a fruit, an amount and a factor near 1; sheet S: criteria blocks and the formulas N1:N6 over them. */
function edited(): BookInput {
  const t = block([['Fruit', 'Amount', 'Factor']], 0, 1)
  for (let r = 2; r <= 301; r++) Object.assign(t, block([[FRUITS[r % 7], r % 11, 1 + (r % 3) / 1000]], 0, r))
  const s = {
    ...block([['Fruit', 'Amount'], ['Pear', '>5']], 7, 1),
    ...block([['Fruit'], ['App'], ['=Kiwi']], 7, 4),
    ...block([['Amount', 'Nope'], ['<3', true]], 7, 8),
    ...block([[{ f: 'DSUM(T!A1:C301,"Amount",H1:I2)' }], [{ f: 'DCOUNT(T!A1:C301,,H4:H6)' }], [{ f: 'DPRODUCT(T!A1:C301,3,H8:I9)' }]], 13, 1),
    ...block([[{ f: 'DGET(T!A1:C301,"Fruit",H4:H6)' }], [{ f: 'DSUM(T!A:C,"Amount",H1:I2)' }], [{ f: 'DAVERAGE(T!A1:C301,"Amount",H4:H6)' }]], 13, 4),
  }
  return { sheets: [{ name: 'S', cells: s }, { name: 'T', cells: t }] }
}

/** Edits of sheet T's amounts, fruits and header, beyond its range, and of the criteria's values and header, in turn. */
const EDITS: [string, string, Scalar][] = [
  ['T', 'B150', 99],
  ['T', 'B151', 98],
  ['T', 'B152', 97],
  ['T', 'A150', 'Kiwi'],
  ['T', 'A400', 'Pear'],
  ['T', 'B400', 9],
  ['S', 'H2', 'Apple'],
  ['S', 'I1', 'Fruit'],
  ['S', 'I1', 'Amount'],
  ['T', 'B1', 'Qty'],
  ['T', 'B1', 'Amount'],
  ['T', 'A2', null],
  ['S', 'H5', '<B'],
]

/** N1:N6 of `edited()` when it opened and after each edit, as the functions' own pass gave them. */
const EDITED: CellValue[][] = [
  [160, 129, 1.08646638728881, err('NUM'), 160, 5],
  [160, 129, 1.08646638728881, err('NUM'), 160, 5],
  [160, 129, 1.08646638728881, err('NUM'), 160, 5],
  [160, 129, 1.08646638728881, err('NUM'), 160, 5.68217054263566],
  [160, 130, 1.08646638728881, err('NUM'), 160, 6.4],
  [160, 130, 1.08646638728881, err('NUM'), 160, 6.4],
  [160, 130, 1.08646638728881, err('NUM'), 169, 6.4],
  [314, 130, 1.08646638728881, err('NUM'), 314, 6.4],
  [0, 130, 1.08646638728881, err('NUM'), 0, 6.4],
  [314, 130, 1.08646638728881, err('NUM'), 314, 6.4],
  [err('VALUE'), 130, 0, err('NUM'), err('VALUE'), err('VALUE')],
  [314, 130, 1.08646638728881, err('NUM'), 314, 6.4],
  [314, 129, 1.08646638728881, err('NUM'), 314, 6.43410852713178],
  [314, 172, 1.08646638728881, err('NUM'), 314, 6.58139534883721],
]

describe.each(['closures', 'eval'] as const)('database functions, %s', (execution) => {
  const opts = { clock, execution, display: false }

  it('edits of an amount check no record one by one and select again in no formula', () => {
    const s = createSession(sums(), opts)
    const edit = (k: number) => s.setCells([{ sheet: 'Data', cell: 'C1702', input: { value: 100 + k } }])
    edit(1)
    const records = { records: 0 }
    const selections = { selects: 0, checks: 0 }
    watchRecords(records)
    watchSelections(selections)
    try {
      edit(2)
      edit(3)
    } finally {
      watchRecords()
      watchSelections()
    }
    expect([records.records, selections.selects]).toEqual([0, 0])
    expect(s.getHash()).toBe(compute(s.getBook(), opts).hash)
    s.close()
  })

  it('answer as their own pass over the records did', () => {
    const cells = { ...DB }
    const at: [string, string][] = []
    let row = 1
    for (const [crit, answers] of CASES) {
      Object.assign(cells, block(crit, 7, row))
      const right = colName(7 + Math.max(...crit.map((r) => r.length)) - 1)
      Object.keys(answers).forEach((f, i) => {
        const a = `${colName(13 + i)}${row}`
        cells[a] = { formula: f.replaceAll('CR', `H${row}:${right}${row + crit.length - 1}`).replaceAll('HR', `H${row}:${right}${row}`) }
        at.push([a, f])
      })
      row += crit.length + 1
    }
    const got = compute({ sheets: [{ name: 'S', cells }, { name: 'T', cells: DB }] }, opts).sheets[0].cells
    expect(at.map(([a, f]) => [f, got[a]?.value ?? null])).toEqual(CASES.flatMap(([, answers]) => Object.entries(answers)))
  })

  it('answer edits of the database, the criteria and the headers as their own pass did', () => {
    const s = createSession(edited(), opts)
    const values = () => [1, 2, 3, 4, 5, 6].map((r) => s.getCell('S', `N${r}`)?.value ?? null)
    const seen = [values()]
    for (const [sheet, cell, value] of EDITS) {
      s.setCells([{ sheet, cell, input: { value } }])
      seen.push(values())
      expect(s.getHash()).toBe(compute(s.getBook(), opts).hash)
    }
    expect(seen).toEqual(EDITED)
    s.close()
  })

  it("a criterion of a date without a year selects by the clock's year after the clock moves", () => {
    const cells: Record<string, CellInput> = { ...datedRows(2, 301), A1: { value: 'Date' }, B1: { value: 'Amount' }, D1: { value: 'Date' }, D2: { value: '1/5' } }
    cells.E1 = cells.E2 = { formula: 'DSUM(A1:B301,"Amount",D1:D2)' }
    const s = createSession({ sheets: [{ name: 'S', cells }] }, opts)
    followsClockYear(s, ['E1', 'E2'], 2, 301)
    s.close()
  })
})
