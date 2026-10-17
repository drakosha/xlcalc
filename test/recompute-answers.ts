/**
 * The answers of `setCells` calls whose recalculation reaches many formulas (a fanout of 120 and its SUM, a chain of 100
 * with cumulative sums, a lookup and statistics over them, a chain an unchanged INT stops, a spill whose size follows
 * the edit, a batch of 70 inputs, two independent roots, a formula replaced in the middle of the fanout), with undo and
 * redo and the book after each, recorded in `test/fixtures/recompute-answers.json` by a synchronous session;
 * `npx tsx test/recompute-answers.ts --write` writes it again. `recompute-answers.test.ts` compares every kind of session
 * in both executions with it, so a faster recalculation keeps each answer and its order.
 */
import { writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import type { BookInput, CellEdit, CellInput } from '../src/index.js'
import { record } from './edit-answers.js'

export const FIXTURE = fileURLToPath(new URL('./fixtures/recompute-answers.json', import.meta.url))

const rows = (n: number, cell: (r: number) => [string, CellInput]) => Object.fromEntries(Array.from({ length: n }, (_, i) => cell(i + 1)))

export const BOOK: BookInput = {
  sheets: [
    {
      name: 'Calc',
      cells: {
        A1: { value: 1 },
        ...rows(120, (r) => [`B${r}`, { formula: `$A$1*${r}`, ...(r % 10 ? {} : { numberFormat: '0.00' }) }]),
        C1: { formula: 'SUM(B1:B120)' },
        D1: { value: 1 },
        ...rows(99, (r) => [`D${r + 1}`, { formula: `D${r}+1` }]),
        ...rows(100, (r) => [`E${r}`, { formula: `SUM(D$1:D${r})` }]),
        F1: { formula: 'INT(G1)' },
        ...rows(79, (r) => [`F${r + 1}`, { formula: `F${r}+1` }]),
        G1: { value: 1.1 },
        H1: { formula: 'SEQUENCE(A1+1,1,10)' },
        I1: { formula: 'SUM(H1#)' },
        ...rows(70, (r) => [`J${r}`, { value: r }]),
        ...rows(70, (r) => [`K${r}`, { formula: `J${r}*3+SUM($L$1:$L$5)` }]),
        ...rows(5, (r) => [`L${r}`, { value: r }]),
        M1: { formula: 'COUNTIF(B1:B120,">50")' },
        M2: { formula: 'AVERAGE(E1:E100)', numberFormat: '0.0' },
        M3: { formula: 'VLOOKUP(D1+4,D1:E100,2,FALSE)' },
        M4: { formula: 'SUMIFS(B1:B120,B1:B120,">"&D1)' },
      },
    },
    { name: 'Other', cells: { A1: { formula: 'Calc!C1+Calc!D100' }, ...rows(69, (r) => [`A${r + 1}`, { formula: `A${r}+Calc!B${r}` }]) } },
  ],
}

const put = (cell: string, input: CellInput | null): CellEdit => ({ sheet: 'Calc', cell, input })

export const STEPS: CellEdit[][] = [
  [put('A1', { value: 2 })],
  [put('A1', { value: 2 })],
  [put('D1', { value: 5 })],
  [put('G1', { value: 1.9 })],
  [put('G1', { value: 2.1 })],
  Array.from({ length: 70 }, (_, i) => put(`J${i + 1}`, { value: (i + 1) * 2 })),
  [put('L3', { value: 10 })],
  [put('B50', { value: 0 })],
  [put('B50', { formula: '$A$1*50' })],
  [put('A1', { value: 3 }), put('D1', { value: 1 })],
  [{ sheet: 'calc', cell: 'D50', input: { value: 'x' } }],
  [put('D50', null)],
]

if (process.argv.includes('--write')) writeFileSync(FIXTURE, JSON.stringify(await record(BOOK, STEPS, 'closures'), null, 1) + '\n')
