/**
 * The answers of a series of structure operations (rows and columns inserted and deleted, a range moved and one
 * copied, a hundred rows inserted) over a varied book (values, formulas and cumulative sums, a spill, an old-kind array
 * formula, formulas a circular reference stops, texts and dates, a cell of a style alone, a note and a link, a merged
 * area, a table with a calculated column, conditional formats, a validation, a hidden row, looks, another sheet's
 * formulas over it), with every undo and some redos and the whole book after each, recorded in
 * `test/fixtures/structure-answers.json` by a synchronous session; `npx tsx test/structure-answers.ts --write` writes it
 * again. `edit-answers.test.ts` compares every kind of session in both executions with it, so a faster structure
 * operation keeps each answer, its changes and moves, and the book.
 */
import { writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import type { BookInput, CellInput, StructureOp } from '../src/index.js'
import { record, type SeriesOptions } from './edit-answers.js'

export const FIXTURE = fileURLToPath(new URL('./fixtures/structure-answers.json', import.meta.url))

const rows = (n: number, cell: (r: number) => [string, CellInput]) => Object.fromEntries(Array.from({ length: n }, (_, i) => cell(i + 1)))
const RED = { fill: { color: '#FFC7CE' } }

export const BOOK: BookInput = {
  styles: [{ font: { bold: true } }, { font: { italic: true } }],
  sheets: [
    {
      name: 'S',
      cells: {
        ...rows(30, (r) => [`A${r}`, { value: r }]),
        ...rows(30, (r) => [`B${r}`, { formula: `A${r}*2`, ...(r % 7 ? {} : { numberFormat: '0.00' }) }]),
        ...rows(30, (r) => [`C${r}`, { formula: `SUM($A$1:A${r})` }]),
        D1: { formula: 'SEQUENCE(5)' },
        E1: { formula: 'A1:A3*10', array: 'E1:E3' },
        F5: { formula: 'F6+1', held: 5 },
        F6: { formula: 'F5+1', held: 4 },
        ...rows(12, (r) => [`G${r}`, { value: r % 3 ? `t${r}` : r % 2 ? '25%' : { date: '2026-01-15' } }]),
        H2: { value: 2.5, numberFormat: '0.00' },
        I3: { style: 0 },
        J4: { value: 'noted', note: { author: 'Ann', text: 'a note' } },
        K5: { value: 'link', link: { url: 'https://example.com/' } },
        K10: { value: 'merged' },
        M1: { value: 'Qty' },
        N1: { value: 'Price' },
        O1: { value: 'Total' },
        ...rows(7, (r) => [`M${r + 1}`, { value: r }]),
        ...rows(7, (r) => [`N${r + 1}`, { value: r * 1.5 }]),
        ...rows(7, (r) => [`O${r + 1}`, { formula: 'T[[#This Row],[Qty]]*T[[#This Row],[Price]]' }]),
      },
      looks: [{ range: 'H1:H5', style: 1 }],
      merges: ['K10:L11'],
      tables: [{ name: 'T', range: 'M1:O8', columns: [{ name: 'Qty' }, { name: 'Price' }, { name: 'Total', formula: 'T[[#This Row],[Qty]]*T[[#This Row],[Price]]' }] }],
      conditionalFormats: [
        { type: 'cellIs', ranges: ['A1:A30'], operator: 'greaterThan', formulas: ['15'], style: RED },
        { type: 'cellIs', ranges: ['B1:B30'], operator: 'greaterThan', formulas: ['$C$5'], style: RED },
      ],
      validations: [{ range: 'G1:G10', type: 'whole', formula1: '1', formula2: '$A$30' }],
      rows: { 7: { hidden: true } },
    },
    { name: 'Other', cells: { A1: { formula: 'SUM(S!B1:B30)' }, A2: { formula: 'S!C30' }, A3: { formula: 'INDEX(S!A:A,5)' }, A4: { formula: 'ROW(S!A10)' }, A5: { formula: 'SUM(S!D1#)' } } },
  ],
}

export const STEPS: StructureOp[] = [
  { op: 'insertRows', sheet: 'S', row: 5, count: 2 },
  { op: 'deleteRows', sheet: 'S', row: 4 },
  { op: 'insertColumns', sheet: 'S', column: 'B' },
  { op: 'deleteColumns', sheet: 'S', column: 'C' },
  { op: 'moveRange', from: { sheet: 'S', range: 'A1:B4' }, to: { sheet: 'S', cell: 'Q1' } },
  { op: 'copyRange', from: { sheet: 'S', range: 'A10:C12' }, to: { sheet: 'S', range: 'Q10:S12' } },
  { op: 'insertRows', sheet: 'S', row: 15, count: 100 },
  { op: 'deleteRows', sheet: 'S', row: 20, count: 3 },
]

export const OPTIONS: SeriesOptions = { undos: STEPS.length, redos: 4, whole: true }

if (process.argv.includes('--write')) writeFileSync(FIXTURE, JSON.stringify(await record(BOOK, STEPS, 'closures', OPTIONS), null, 1) + '\n')
