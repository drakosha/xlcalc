/**
 * The texts a session keeps for `findCells` (`src/session/search.ts`) change no answer: after each call of a series that
 * changes the book (cells, formulas, formats, widths, hidden rows, a structure operation, undo and redo, names, styles,
 * an array of the old kind, a spill that grows, conditional formats that read another cell, a table, a replacement), every kind of
 * search over the session that searched before gives what a new session over the same book and clock gives. In a
 * synchronous session, in the same thread and in a thread, with closures and eval, with display and without.
 */
import { afterAll, describe, expect, it } from 'vitest'
import type { AsyncSession, BookInput, Execution, FindCellsOptions } from '../src/index.js'
import { createSession } from '../src/index.js'
import './locales.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import { filteredSheet } from './find-books.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-10T12:00:00Z', seed: 1 }

const book: BookInput = {
  locale: 'ru-RU',
  styles: [{ font: { bold: true } }],
  sheets: [
    {
      name: 'S',
      cells: {
        A8: { value: 2.5 },
        A1: { value: 'Hello' }, A2: { value: 'чай 1' }, A3: { value: 0.05, numberFormat: '0%' }, A4: { value: 46096, numberFormat: 14 },
        A5: { value: 1234.5, numberFormat: '#,##0.00' }, A6: { value: true }, A7: { value: '007' }, A10: { value: 'hid 1' },
        B1: { formula: 'SUM(D1:D2)' }, B2: { formula: 'A3*2', numberFormat: '0%' }, B3: { formula: 'TODAY()', numberFormat: 14 },
        C1: { value: 123456789 },
        D1: { value: 1 }, D2: { value: 2 },
        F1: { formula: 'SEQUENCE(D2)*7' },
        H1: { formula: 'D1:D2*5', array: 'H1:H2' },
        J1: { value: 5 }, J2: { value: 50 },
        L1: { value: 'Qty' }, M1: { value: 'Twice' }, L2: { value: 2 }, L3: { value: 3 }, M2: { formula: '[@Qty]*2' }, M3: { formula: '[@Qty]*2' },
      },
      columns: { C: { width: 3 } },
      rows: { 10: { hidden: true } },
      conditionalFormats: [{ type: 'cellIs', ranges: ['J1:J5'], operator: 'greaterThan', formulas: ['$D$1*10'], style: { numberFormat: '0.0' } }],
      tables: [{ name: 'T', range: 'L1:M3', columns: [{ name: 'Qty' }, { name: 'Twice', formula: '[@Qty]*2' }] }],
    },
    filteredSheet(),
  ],
}

const ALL = [{ sheet: 'S', range: 'A1:XFD1048576' }, { sheet: 'F', range: 'A1:XFD1048576' }]

/** The searches compared: each kind, wildcards and literal, in any case and in the cells' own, by rows and by columns, hidden cells skipped. */
const SEARCHES: [string, FindCellsOptions][] = (['text', 'edit', 'formula', 'value'] as const).flatMap((where): [string, FindCellsOptions][] => [
  ['*', { in: where, wildcards: true, skipHidden: true }],
  ['1', { in: where }],
  ['?', { in: where, wildcards: true, wholeCell: true, byColumns: true }],
  ['чай', { in: where, caseSensitive: true }],
  ['x*', { in: where, wildcards: true, skipHidden: true, byColumns: true }],
  // a number by a format of a conditional format's, a number too wide for its column
  [',0', { in: where }],
  ['#', { in: where }],
])

/** What every search gives: its cells with their texts, and the count. */
async function answers(s: AsyncSession) {
  const out: unknown[] = []
  for (const [query, o] of SEARCHES) {
    const p = await s.findCells(ALL, query, { ...o, fields: ['value', 'text', 'editText'], count: true, maxCells: 10000 })
    out.push([query, o, p.total, p.cells])
  }
  return out
}

/** The calls of the series, each changing the book. */
const SERIES: [string, (s: AsyncSession) => Promise<unknown>][] = [
  ['a value others read, a conditional format among them', (s) => s.setCells([{ sheet: 'S', cell: 'D1', input: { value: 6 } }])],
  ['a new cell', (s) => s.setCells([{ sheet: 'S', cell: 'E7', input: { value: 'чайник' } }])],
  ['a formula', (s) => s.setCells([{ sheet: 'S', cell: 'B1', input: { formula: 'MAX(D1:D2)+1' } }])],
  ['a number format', (s) => s.patchCells([{ sheet: 'S', range: 'A5', set: { numberFormat: '0.0%' } }])],
  ['a column narrowed', (s) => s.setLayout([{ sheet: 'S', columns: 'A', set: { width: 2 } }])],
  ['a row hidden', (s) => s.setLayout([{ sheet: 'S', rows: '2', set: { hidden: true } }])],
  ["a row's format", (s) => s.setLayout([{ sheet: 'S', rows: '8', set: { numberFormat: '0.000' } }])],
  ['contents cleared', (s) => s.clearCells([{ sheet: 'S', range: 'A1:A2' }], 'contents')],
  ['a text typed', (s) => s.setCellText('S', 'A3', '7%')],
  ['a text entered into a range', (s) => s.fillCellText('S', 'P1:P3', 'чай=1')],
  ['text pasted', (s) => s.pasteText('S', 'Q1', 'a\t1\nb\t2')],
  ['a row inserted', (s) => s.editStructure({ op: 'insertRows', sheet: 'S', row: 2 })],
  ['undo', (s) => s.undo()],
  ['undo again', (s) => s.undo()],
  ['redo', (s) => s.redo()],
  ['an array of the old kind, its values kept', (s) => s.setCells([{ sheet: 'S', cell: 'H1', input: { formula: 'D1:D2*5+0', array: 'H1:H2' } }])],
  ['a spill that grows', (s) => s.setCells([{ sheet: 'S', cell: 'D2', input: { value: 4 } }])],
  ['a table column', (s) => s.setCells([{ sheet: 'S', cell: 'L3', input: { value: 11 } }])],
  ['names', (s) => s.setNames([{ name: 'Rate', ref: '2' }])],
  ['styles', (s) => s.setStyles([{ font: { bold: true } }, { font: { italic: true } }])],
  ['the clock', (s) => s.setClock({ now: '2027-01-02T12:00:00Z', seed: 1 })],
  ['mixed changes', (s) => s.applyChanges([{ op: 'setCells', edits: [{ sheet: 'F', cell: 'B3', input: { value: 'x9' } }] }, { op: 'replaceCells', edits: [{ sheet: 'S', cell: 'A6', input: { value: 'x false' } }] }])],
  ['a replacement', (s) => s.replaceInCells(ALL, 'x*', 'чай', { wildcards: true })],
  ['a hidden row on the filtered sheet shown', (s) => s.setLayout([{ sheet: 'F', rows: '2', set: { hidden: false } }])],
]

describe.each(['closures', 'eval'] as const)('findCells over the texts it keeps, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    // without display a change lists no cell whose shown text alone changed
    it.each([true, false])('answers after each change as a new session over the same book, display %s', async (display) => {
      const s = await (await opener(kind))(book, { clock, execution, display })
      await answers(s)
      for (const [what, call] of SERIES) {
        await call(s)
        const fresh = createSession(await s.getBook(), { clock: s.clock, execution, display })
        const want = await answers({ findCells: async (...a: Parameters<AsyncSession['findCells']>) => fresh.findCells(...a) } as AsyncSession)
        fresh.close()
        expect([what, await answers(s)]).toEqual([what, want])
      }
      await s.close()
    }, 60_000)
  })
})
