/**
 * The answers of a series of `setCells` calls (fields over a cell's input, content replaced or kept, a cell edited twice
 * in a call, sheets named in another case, looks, notes, links, formulas with held values and array areas, a batch of a
 * column, a refused one) and of undo and redo, with the book after each, recorded in `test/fixtures/edit-answers.json` by a synchronous
 * session; `npx tsx test/edit-answers.ts --write` writes it again. `edit-answers.test.ts` compares every kind of session
 * in both executions with it, so faster editing keeps each answer.
 */
import { writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { createSession, type AsyncSession, type BookInput, type CellEdit, type Execution, type StructureOp } from '../src/index.js'
import type { Labelled } from './recorded.js'
import { wrapped } from './session-kinds.js'

export const CLOCK = { now: '2026-10-16T12:00:00Z', seed: 1 }
export const FIXTURE = fileURLToPath(new URL('./fixtures/edit-answers.json', import.meta.url))

const column = Array.from({ length: 40 }, (_, r) => r + 1)
export const BOOK: BookInput = {
  styles: [{ font: { bold: true } }, { font: { italic: true } }],
  sheets: [
    {
      name: 'Data',
      cells: {
        A1: { value: 1, style: 0, note: { text: 'n' }, link: { url: 'https://example.com/' }, meta: { id: 1 } },
        A2: { formula: 'A1*2', numberFormat: '0.00' },
        A3: { value: { date: '2026-01-15' }, numberFormat: 14 },
        A4: { formula: 'B1:B2*3', array: 'A4:A5' },
        A6: { formula: 'A7+1', held: 5 },
        A7: { formula: 'A6+1', held: 4 },
        B1: { value: 'x' },
        B2: { value: 2 },
        C1: { formula: 'SUM(D1:D40)' },
        ...Object.fromEntries(column.map((r) => [`D${r}`, { value: r }])),
        ...Object.fromEntries(column.map((r) => [`E${r}`, { formula: `D${r}*2+C1` }])),
      },
      looks: [{ range: 'G1:G3', style: 1 }],
    },
    { name: 'Другой', cells: { A1: { formula: "Data!A1+1" } } },
  ],
}

/** The calls of the series: each a list of edits. */
export const STEPS: CellEdit[][] = [
  [{ sheet: 'Data', cell: 'A1', input: { value: 2 } }],
  [{ sheet: 'data', cell: 'A1', input: { formula: 'B2*5' } }],
  [{ sheet: 'Data', cell: 'A1', input: { formula: null } }],
  [{ sheet: 'Data', cell: 'A1', input: { style: null, note: null } }],
  [{ sheet: 'Data', cell: 'A2', input: { value: 3 } }, { sheet: 'Data', cell: 'A2', input: { style: 1 } }],
  [{ sheet: 'Data', cell: 'A3', input: { value: { date: '2026-01-15' } } }],
  [{ sheet: 'Data', cell: 'A3', input: { numberFormat: '0.0' } }],
  [{ sheet: 'Data', cell: 'A3', input: null }],
  [{ sheet: 'Data', cell: 'A5', input: { value: 1 } }],
  [{ sheet: 'Data', cell: 'A4', input: null }, { sheet: 'Data', cell: 'A5', input: null }],
  [{ sheet: 'Data', cell: 'A6', input: { value: 1 } }],
  [{ sheet: 'Data', cell: 'A9', input: { formula: 'D1:D3*2', array: 'A9:A11' } }],
  [{ sheet: 'Data', cell: 'G2', input: { value: 'over a look' } }],
  [{ sheet: 'Data', cell: 'G3', input: { style: null } }],
  [{ sheet: 'Data', cell: 'B1', input: { link: { url: 'https://example.org/' }, meta: { k: [1, 2] } } }],
  [{ sheet: 'Data', cell: 'B1', input: null }],
  [{ sheet: 'Data', cell: 'B2', input: { value: 2 } }],
  [{ sheet: 'Data', cell: 'B2', input: { value: '2' } }],
  [{ sheet: 'Data', cell: 'B2', input: { value: null } }],
  column.map((r) => ({ sheet: 'Data', cell: `D${r}`, input: { value: r * 10 } })),
  [...column.map((r) => ({ sheet: 'DATA', cell: `D${r}`, input: { value: -r } })), { sheet: 'Другой', cell: 'B1', input: { value: true } }, { sheet: 'другой', cell: 'B1', input: { style: 0 } }],
]

/** A JSON text, `undefined` for none. */
const json = (x: unknown) => JSON.stringify(x) ?? 'undefined'

/** A call of a series: edits of cells (`setCells`) or a structure operation (`editStructure`). */
export type Step = CellEdit[] | StructureOp

/** How a series ends and what it records of the book: its undos and redos (four and two), the whole book or its cells (the default). */
export interface SeriesOptions {
  undos?: number
  redos?: number
  whole?: boolean
}

/** A value with its objects' keys in order, but a sheet's cells, whose order is the book's: a thread gives a sheet's fields in another order. */
const ordered = (x: unknown, key = ''): unknown =>
  Array.isArray(x) ? x.map((v) => ordered(v)) : x && typeof x === 'object' && key !== 'cells' ? Object.fromEntries(Object.keys(x).sort().map((k) => [k, ordered((x as Record<string, unknown>)[k], k)])) : x

/** The answers of a session to calls `steps`: each call's, the book after it, then the undos and redos. */
export async function series(s: AsyncSession, steps: Step[], o: SeriesOptions = {}): Promise<Labelled> {
  const out: Labelled = []
  const book = async () => json(o.whole ? ordered(await s.getBook()) : (await s.getBook()).sheets.map((x) => x.cells))
  for (const [i, step] of steps.entries()) {
    const [name, call] = Array.isArray(step) ? ['setCells', () => s.setCells(step)] : ['editStructure', () => s.editStructure(step)]
    // a refused call answers with its error
    out.push([`${name} ${i}`, await (call() as Promise<unknown>).then(json, (e: Error) => `${e.constructor.name}: ${e.message}`)])
    out.push([`book after ${i}`, await book()])
  }
  for (let i = 0; i < (o.undos ?? 4); i++) out.push([`undo ${i}`, json(await s.undo())], [`book after undo ${i}`, await book()])
  for (let i = 0; i < (o.redos ?? 2); i++) out.push([`redo ${i}`, json(await s.redo())], [`book after redo ${i}`, await book()])
  return out
}

/** The recorded answers: by display, each a label and the answer's JSON. */
export type Recorded = Record<string, Labelled>

/** The answers of a synchronous session of `book` to calls `steps`, with display off and on. */
export async function record(book: BookInput, steps: Step[], execution: Execution, o?: SeriesOptions): Promise<Recorded> {
  const out: Recorded = {}
  for (const display of [false, true]) {
    const s = wrapped(createSession(book, { clock: CLOCK, execution, display }))
    out[String(display)] = await series(s, steps, o)
    await s.close()
  }
  return out
}

if (process.argv[1] === fileURLToPath(import.meta.url) && process.argv.includes('--write')) writeFileSync(FIXTURE, JSON.stringify(await record(BOOK, STEPS, 'closures'), null, 1) + '\n')
