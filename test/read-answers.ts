/**
 * The answers of `readCells`, `getCell` and `getCells` over two varied books (input cells with formulas and values,
 * styles and formats, notes and links, cells of a look alone, a spill, an array formula of the old kind, a merged area,
 * a hidden row, conditional formats; one book in ru-RU), before and after a series of edits, recorded in
 * `test/fixtures/read-answers.json` by a synchronous session; `npx tsx test/read-answers.ts --write` writes it again.
 * `read-answers.test.ts` compares every kind of session in both executions with it, so faster reading keeps each answer.
 */
import { writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createSession, type AsyncSession, type BookInput, type CellReadField, type Execution, type RangeRef, type ReadCellsOptions } from '../src/index.js'
import './locales.js'
import { wrapped } from './session-kinds.js'

export const CLOCK = { now: '2026-10-16T12:00:00Z', seed: 1 }
export const FIXTURE = fileURLToPath(new URL('./fixtures/read-answers.json', import.meta.url))

const RED = { fill: { color: '#FFC7CE' } }

export const BOOKS: Record<string, BookInput> = {
  en: {
    styles: [{ font: { bold: true } }, { font: { italic: true } }],
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 1, style: 0, numberFormat: '0.00' },
          A2: { value: 2.5 },
          A3: { value: 'text' },
          A4: { value: true },
          A5: { value: { date: '2026-01-15' }, numberFormat: 14 },
          A6: { formula: 'A1/0' },
          A7: { formula: 'SUM(A1:A2)', numberFormat: '0.0%' },
          A8: { value: -3, note: { author: 'Ann', text: 'a note' } },
          A9: { value: 'link', link: { url: 'https://example.com/' } },
          A10: { formula: 'NOW()', numberFormat: 22 },
          B1: { formula: 'SEQUENCE(3,2)' },
          E1: { formula: 'A1:A3*2', array: 'E1:E3' },
          F5: { formula: 'A1+B2', style: 1 },
          F6: { value: '' },
          G2: { value: 5 },
          K1: { value: 'merged' },
          C10: { value: 1234567.891, numberFormat: '#,##0.00' },
          D10: { value: '=not a formula' },
          D11: { value: '001' },
          H12: { formula: 'T!A1*2' },
        },
        looks: [
          { range: 'G1:H4', style: 1 },
          { range: 'J2:J3', numberFormat: '0.000' },
        ],
        merges: ['K1:L2'],
        rows: { 6: { hidden: true } },
        conditionalFormats: [{ type: 'cellIs', ranges: ['A1:A9'], operator: 'greaterThan', formulas: ['2'], style: RED }],
      },
      { name: 'T', cells: { A1: { formula: 'S!A1*10' }, B2: { value: 'x' }, C3: { formula: 'SEQUENCE(2)' } } },
    ],
  },
  ru: {
    locale: 'ru-RU',
    styles: [{ font: { bold: true } }],
    sheets: [
      {
        name: 'Лист1',
        cells: {
          A1: { value: 1234.5, numberFormat: '#,##0.00' },
          A2: { formula: 'SUM(A1,1)' },
          A3: { value: { date: '2026-03-08' }, numberFormat: 14 },
          A4: { value: 0.5, numberFormat: '0%' },
          B1: { value: 'текст' },
          B2: { formula: 'IF(A1>1,TRUE,FALSE)' },
          B3: { formula: 'A1/0' },
          C1: { formula: 'SEQUENCE(2,2)' },
        },
        looks: [{ range: 'E1:E3', numberFormat: '0.0' }],
      },
    ],
  },
}

const ALL: CellReadField[] = ['value', 'text', 'formula', 'editText', 'style', 'numberFormat', 'note', 'link', 'issue']

/** A query of `readCells`: its ranges on the book's first sheet (or as given) and its options. */
type Query = [ranges: (string | RangeRef)[], opts?: ReadCellsOptions]

function queries(second: string | undefined): Query[] {
  const out: Query[] = [
    [['A1:L12']],
    ...ALL.map((f): Query => [['A1:L12'], { fields: [f] }]),
    [['A1:L12'], { fields: ALL }],
    [['A1:L12'], { fields: [] }],
    [['B:B'], { fields: ['value', 'formula'] }],
    [['1:1'], { fields: ALL }],
    [['A1:C5', 'B2:H8'], { fields: ALL }],
    [['C3:C3', 'A1:L12', 'E2:F6'], { fields: ['value', 'text'] }],
    [['A1:L12'], { rows: [1, 3, 6, 10], fields: ['value', 'text', 'style'] }],
    [['A1:XFD1048576'], { fields: ['value'] }],
    [['A1:XFD1048576'], { fields: ['value'], maxCells: 7, count: true }],
    [['L12:A1'], { fields: ['value'], maxCells: 1 }],
  ]
  if (second) out.push([['A1:C3', { sheet: second, range: 'A1:C3' }], { fields: ALL }])
  return out
}

/** A JSON text, `undefined` for none. */
const json = (x: unknown) => JSON.stringify(x) ?? 'undefined'

/** The answers of the queries, `getCell` of each cell of the window and `getCells` of it, labelled. */
async function read(s: AsyncSession, sheet: string, second: string | undefined, out: [string, string][]) {
  const ref = (r: string | RangeRef): RangeRef => (typeof r === 'string' ? { sheet, range: r } : r)
  for (const [ranges, opts] of queries(second)) {
    const label = `readCells ${json(ranges)} ${json(opts)}`
    let page = await s.readCells(ranges.map(ref), opts)
    out.push([label, json(page)])
    // the pages after the first by the cursor
    for (let n = 1; page.cursor; n++) out.push([`${label} page ${n}`, json((page = await s.readCells(ranges.map(ref), { ...opts, cursor: page.cursor })))])
  }
  for (let r = 1; r <= 12; r++)
    for (let c = 0; c < 12; c++) {
      const cell = String.fromCharCode(65 + c) + r
      out.push([`getCell ${cell}`, json(await s.getCell(sheet, cell))])
    }
  out.push(['getCells A1:L12', json(await s.getCells(sheet, 'A1:L12'))])
  out.push(['getCells [A1:C3, B2:D4]', json(await s.getCells(sheet, ['A1:C3', 'B2:D4']))])
}

/** The answers of a session of book `name`: read, then after each edit of a series (a cell cleared, cells written, a look, rows inserted and undone). */
export async function answers(s: AsyncSession, name: string): Promise<[string, string][]> {
  const book = BOOKS[name]
  const sheet = book.sheets[0].name
  const second = book.sheets[1]?.name
  const out: [string, string][] = []
  await read(s, sheet, second, out)
  const steps: [string, () => Promise<unknown>][] = [
    ['A3 cleared, D2 and Z5 written', () => s.setCells([{ sheet, cell: 'A3', input: null }, { sheet, cell: 'D2', input: { value: 7 } }, { sheet, cell: 'Z5', input: { formula: 'A1*3' } }])],
    ['A2:A4 contents cleared', () => s.clearCells([{ sheet, range: 'A2:A4' }], 'contents')],
    ['a look written on I1', () => s.setCells([{ sheet, cell: 'I1', input: { style: 0 } }])],
    ['two rows inserted before row 2', () => s.editStructure({ op: 'insertRows', sheet, row: 2, count: 2 })],
    ['undone', () => s.undo()],
  ]
  for (const [step, edit] of steps) {
    await edit()
    const before = out.length
    await read(s, sheet, second, out)
    for (let i = before; i < out.length; i++) out[i][0] = `${step}: ${out[i][0]}`
  }
  return out
}

/** The recorded answers: by book and display, each a label and the answer's JSON. */
export type Recorded = Record<string, Record<string, [string, string][]>>

async function record(execution: Execution): Promise<Recorded> {
  const out: Recorded = {}
  for (const name of Object.keys(BOOKS)) {
    out[name] = {}
    for (const display of [false, true]) {
      const s = wrapped(createSession(BOOKS[name], { clock: CLOCK, execution, display }))
      out[name][String(display)] = await answers(s, name)
      await s.close()
    }
  }
  return out
}

if (process.argv.includes('--write')) writeFileSync(FIXTURE, JSON.stringify(await record('closures'), null, 1) + '\n')
