/**
 * What a sheet's window needs beside its cells: `Session.extent` (where a sheet's cells end, spills too) against the
 * cells of the session's own result and the merged areas and tables of its book, while edits grow and shrink it;
 * `Session.layout` (the book without its cells); `Session.cells` over several rectangles, each cell once, as the
 * rectangles read one by one; an old-kind array formula's area on its view. In a synchronous session, in the same
 * thread and in a thread, with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { createInlineCalc, InputError, openSession, type AsyncSession, type BookInput, type Calc, type CellView, type Execution, type LayoutEdit, type Session, type SessionOptions, type SheetExtent, type StructureOp } from '../src/index.js'
import { compareAddress } from '../src/model.js'
import { synthBook } from '../src/synth.js'
import { CLOCK } from './async-books.js'
import { pool } from './fixtures/pool.js'
import { randomBook } from './seeded-books.js'
import { spillBook } from './session-spill-edits.js'
import { extentOf } from './sheet-extent.js'

/** A synchronous session behind the asynchronous API, so that one test covers the three ways. */
function wrapped(s: Session): AsyncSession {
  const out: Record<string, unknown> = {}
  for (const k of ['setCells', 'patchCells', 'setNames', 'setClock', 'restructure', 'setStyles', 'setLayout', 'book', 'layout', 'extent', 'explain', 'cells', 'validation', 'validateEntry', 'result', 'hash', 'close'] as const)
    out[k] = async (...a: unknown[]) => (s[k] as (...a: unknown[]) => unknown)(...a)
  return Object.defineProperty(out, 'clock', { get: () => s.clock }) as unknown as AsyncSession
}

const KINDS = ['a synchronous session', 'the same thread', 'a thread'] as const
type Kind = (typeof KINDS)[number]

const calcs: Calc[] = []
afterAll(() => Promise.all(calcs.map((c) => c.close())))

/** Opens sessions of one kind. */
async function opener(kind: Kind): Promise<(book: BookInput, opts: SessionOptions) => Promise<AsyncSession>> {
  if (kind === 'a synchronous session') return async (book, opts) => wrapped(openSession(book, opts))
  const calc = kind === 'the same thread' ? createInlineCalc({ sliceMs: 2 }) : await pool({ threads: 1 })
  calcs.push(calc)
  return (book, opts) => calc.openSession(book, opts)
}

/** A sheet of values, an old-kind array formula, a spill as tall as A1 says, a merged area and a table; a sheet of nothing; one spilling down a whole column. */
const crafted = (): BookInput => ({
  styles: [{ font: { bold: true } }, { fill: { color: '#FFFF00' } }],
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 4 }, A2: { value: 2 }, A3: { value: 'x' }, A4: { value: true }, A5: { value: 5, style: 0 },
        C1: { formula: 'A1:A3*10', array: 'C1:C3' },
        E1: { formula: 'SEQUENCE(A1)' },
        M1: { value: 'a' }, N1: { value: 'b' }, M2: { value: 1 }, N3: { value: 2 },
      },
      merges: ['J5:K6'],
      tables: [{ name: 'T', range: 'M1:N3', columns: [{ name: 'a' }, { name: 'b' }] }],
    },
    { name: 'Empty', cells: {} },
    { name: 'Far', cells: { A1: { value: 1 }, A2: { value: 2 }, B1: { formula: 'A:A*2' }, Z500: { style: 1 } } },
  ],
})

/** The steps of a session whose extent grows and shrinks, with the extent of `S` after each where it is known by hand. */
const STEPS: [string, (s: AsyncSession) => Promise<unknown>, SheetExtent?][] = [
  ['open', async () => {}, { rows: 6, columns: 14 }],
  ['a taller spill', (s) => s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 12 } }]), { rows: 12, columns: 14 }],
  ['a value blocks the spill', (s) => s.setCells([{ sheet: 'S', cell: 'E8', input: { value: 'x' } }]), { rows: 8, columns: 14 }],
  ['cleared, the spill comes back', (s) => s.setCells([{ sheet: 'S', cell: 'E8', input: null }]), { rows: 12, columns: 14 }],
  ['a shorter spill', (s) => s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 2 } }]), { rows: 6, columns: 14 }],
  ['a style alone, far away', (s) => s.patchCells([{ sheet: 'S', cell: 'P40', set: { style: 1 } }]), { rows: 40, columns: 16 }],
  ['the style removed: no cell is left', (s) => s.patchCells([{ sheet: 'S', cell: 'P40', set: { style: null } }]), { rows: 6, columns: 14 }],
  ['a value further down, then cleared', async (s) => {
    await s.setCells([{ sheet: 'S', cell: 'Q900', input: { value: 1 } }])
    expect(await s.extent('S')).toEqual({ rows: 900, columns: 17 })
    await s.setCells([{ sheet: 'S', cell: 'Q900', input: null }])
  }, { rows: 6, columns: 14 }],
  ['no merged area', (s) => s.setLayout([{ sheet: 'S', merges: [] }]), { rows: 5, columns: 14 }],
  ['the bottom value cleared', (s) => s.setCells([{ sheet: 'S', cell: 'A5', input: null }]), { rows: 4, columns: 14 }],
  ['rows inserted above', (s) => s.restructure({ op: 'insertRows', sheet: 'S', row: 1, count: 3 }), { rows: 7, columns: 14 }],
  ['the table deleted', (s) => s.restructure({ op: 'deleteTable', table: 'T' } satisfies StructureOp), { rows: 7, columns: 5 }],
  ['the far cell of the other sheet cleared, a spill blocked there', (s) => s.setCells([{ sheet: 'Far', cell: 'Z500', input: null }, { sheet: 'Far', cell: 'B9', input: { value: 0 } }])],
  ['merged areas of the other sheet', (s) => s.setLayout([{ sheet: 'Far', merges: ['C2000:D2001'] } satisfies LayoutEdit])],
]

describe.each(['closures', 'eval'] as const)('a window beside the cells, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it('extent: books of the tests, against their results', async () => {
      const open = await opener(kind)
      const books: [string, BookInput][] = [['the synthetic book', synthBook({ seed: 1, rows: 340 }).book], ['the spill book', spillBook()], ['a seeded book', randomBook(11)], ['a crafted book', crafted()]]
      for (const [name, book] of books) {
        const s = await open(book, { clock: CLOCK, execution })
        const [r, b] = [await s.result(), await s.book()]
        for (const sheet of b.sheets) expect(await s.extent(sheet.name), `${name}: ${sheet.name}`).toEqual(extentOf(r, b, sheet.name))
        await s.close()
      }
    }, 120_000)

    it('extent: edits that grow it and shrink it, without the display too', async () => {
      const open = await opener(kind)
      for (const display of [true, false]) {
        const s = await open(crafted(), { clock: CLOCK, execution, display })
        expect(await s.extent('Empty')).toEqual({ rows: 0, columns: 0 })
        expect(await s.extent('Far')).toEqual({ rows: 1_048_576, columns: 26 })
        for (const [step, apply, known] of STEPS) {
          await apply(s)
          const [r, b] = [await s.result(), await s.book()]
          if (known) expect(await s.extent('S'), step).toEqual(known)
          for (const sheet of b.sheets) expect(await s.extent(sheet.name), `${step}: ${sheet.name}`).toEqual(extentOf(r, b, sheet.name))
        }
        expect(await s.extent('Far')).toEqual({ rows: 2001, columns: 4 })
        await s.close()
      }
    }, 60_000)

    it('layout: the book without its cells, after layout edits and a structure operation', async () => {
      const open = await opener(kind)
      const book: BookInput = {
        ...crafted(),
        names: [{ name: 'rate', ref: "'S'!$A$2" }],
        locale: 'ru-RU',
        dateSystem: 1904,
        iteration: { maxIterations: 10 },
      }
      const S = book.sheets[0]
      Object.assign(S, {
        columns: { B: { width: 14, style: 1 }, 'D:F': { hidden: true, outlineLevel: 1 } },
        rows: { '2': { height: 30 }, '7:9': { format: '0.00' } },
        filtered: true,
        frozen: { rows: 1, columns: 2 },
        conditionalFormats: [{ type: 'cellIs', ranges: ['A1:A5'], operator: 'greaterThan', formulas: ['3'], style: { fill: { color: '#FFC7CE' } } }],
        validations: [{ range: 'A1:A10', type: 'whole', formula1: '1', formula2: '100' }],
        outline: { summaryBelow: false },
      })
      const s = await open(book, { clock: CLOCK, execution })
      const without = (b: BookInput) => JSON.parse(JSON.stringify({ ...b, sheets: b.sheets.map(({ cells: _, ...rest }) => rest) }))
      expect(await s.layout()).toEqual(without(book))
      await s.setLayout([{ sheet: 'S', columns: 'B', set: { width: 20 } }, { sheet: 'Far', frozen: { rows: 2 } }])
      expect((await s.restructure({ op: 'insertRows', sheet: 'S', row: 5 })).ok).toBe(true)
      await s.setStyles([...book.styles!, { font: { italic: true } }])
      const now = await s.layout()
      expect(now).toEqual(without(await s.book()))
      expect(now.sheets[0].columns!.B.width).toBe(20)
      expect(now.sheets[0].merges).toEqual(['J6:K7'])
      expect(now.sheets.every((x) => !('cells' in x))).toBe(true)
      await s.close()
    })

    it('cells of several rectangles: each cell once, in the order of addresses, as the rectangles one by one', async () => {
      const open = await opener(kind)
      const book = crafted()
      book.sheets[0].conditionalFormats = [{ type: 'expression', ranges: ['A1:F8'], formula: 'ISNUMBER(A1)', style: { font: { bold: true }, format: '0.0' } }]
      const sets = [['A1:C2', 'B2:F4'], ['E:E', 'A3:E3', 'E2'], ['1:2', 'C1', 'C1'], ['A1:XFD1048576', 'M1:N3'], ['J4:K7', 'A10'], []]
      for (const [name, b] of [['a crafted book', book], ['the spill book', spillBook()]] as [string, BookInput][]) {
        const s = await open(b, { clock: CLOCK, execution })
        for (const sheet of (await s.book()).sheets)
          for (const ranges of sets) {
            const got = await s.cells(sheet.name, ranges)
            const one: Record<string, CellView> = {}
            for (const r of ranges) Object.assign(one, await s.cells(sheet.name, r))
            expect(got, `${name}: ${sheet.name}!${ranges}`).toEqual(one)
            expect(Object.keys(got)).toEqual(Object.keys(one).sort(compareAddress))
          }
        await s.close()
      }
    }, 60_000)

    it("an old-kind array formula's area on its top-left cell, as its spill", async () => {
      const open = await opener(kind)
      const s = await open(crafted(), { clock: CLOCK, execution })
      const w = await s.cells('S', ['C1:C3', 'E1'])
      expect(w.C1).toEqual({ value: 40, text: '40', formula: 'A1:A3*10', spill: 'C1:C3', array: 'C1:C3' })
      expect(w.C2).toEqual({ value: 20, text: '20', spillOf: 'C1' })
      expect(w.E1).toEqual({ value: 1, text: '1', formula: 'SEQUENCE(A1)', spill: 'E1:E4' })
      await s.close()
    })
  })
})

describe('a window beside the cells', () => {
  it('the sheet and the ranges are checked as input', () => {
    const s = openSession(crafted(), { clock: CLOCK })
    expect(() => s.extent('Nope')).toThrow(InputError)
    expect(() => s.cells('S', ['A1', 'A0'])).toThrow('range[1]: expected a range such as "B2", "B2:D9", "B:D" or "2:9" within the sheet, got "A0"')
    expect(() => s.cells('S', [7 as unknown as string])).toThrow('range[0]: expected a string')
    expect(s.cells('S', [])).toEqual({})
    expect(s.extent('far')).toEqual({ rows: 1_048_576, columns: 26 })
  })

  it("a synchronous session's layout is made of its book's objects; a thread's is a copy", async () => {
    const s = openSession(crafted(), { clock: CLOCK })
    const l = s.layout()
    expect(l.sheets[0].merges).toBe(s.book().sheets[0].merges)
    expect(l.styles).toBe(s.book().styles)
    const p = await pool({ threads: 1 })
    calcs.push(p)
    const t = await p.openSession(crafted(), { clock: CLOCK })
    const a = await t.layout()
    a.sheets[0].merges!.push('A20:B20')
    expect((await t.layout()).sheets[0].merges).toEqual(['J5:K6'])
  })
})
