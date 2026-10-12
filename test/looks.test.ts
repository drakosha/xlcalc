/**
 * A book with looks (`SheetInput.looks`) and the same book with every cell of its looks as an entry of a look alone
 * (`expandBook`) are one book: the same hashes, windows, cells, queries, used ranges and book given back, the same xlsx
 * parts, which read back give the looks as the session holds them. In a synchronous session, in the same thread and in a
 * thread, with closures and eval (formats by code that no built-in format has, which a file read gives back by number). Then the looks' cells as a session shows and edits them, by what an entry of a look
 * alone gave, and a book without looks hashing as before.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import { compute, createSession, type AsyncSession, type BookInput, type LookRange } from '../src/index.js'
import { formatAddress } from '../src/model.js'
import { synthBook } from '../src/synth.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { CLOCK } from './async-books.js'
import { cellsOf, expandBook } from './looks-book.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const STYLES = [{ font: { bold: true } }, { fill: { color: '#FFF2CC' } }, { font: { italic: true } }, { border: { bottom: { style: 'thin' as const } } }]

/** Rows in threes with two empty rows between, of style 0 with style 1 and a percent in every 12th column, as the wide shape's sheet, `cols` wide. */
function wideLooks(groups: number, cols: number, top = 20): LookRange[] {
  const out: LookRange[] = []
  for (let g = 0; g < groups; g++) {
    const r = top + g * 5
    out.push({ range: `${formatAddress(r - 1, 0)}:${formatAddress(r + 1, cols - 1)}`, style: 0, numberFormat: '#,##0.0' })
    for (let c = 17; c < cols; c += 12) out.push({ range: `${formatAddress(r - 1, c)}:${formatAddress(r + 1, c)}`, style: 1, numberFormat: '0.0%' })
  }
  return out
}

const BOOKS: [string, BookInput][] = [
  [
    'patterns with gaps, content above and within',
    {
      styles: STYLES,
      sheets: [
        { name: 'G', cells: { A1: { value: 'h', style: 0 }, B1: { value: 2 }, C2: { formula: 'B1*3' }, E20: { value: 0.25 }, R21: { value: 0.5 }, S22: { formula: 'R21*2', numberFormat: '0.00' } }, looks: wideLooks(6, 60) },
        { name: 'Data', cells: { A1: { formula: 'SUM(G!E20:R21)' } } },
      ],
    },
  ],
  [
    'rectangles over each other, one of none, whole rows and columns, single cells',
    {
      styles: STYLES,
      sheets: [{
        name: 'S',
        cells: { B2: { value: 1 }, D4: { value: 'x', style: 3 }, F6: { note: { text: 'a note' } }, G7: {} },
        looks: [{ range: 'A1:H10', style: 0 }, { range: 'C3:E5', numberFormat: '0.000' }, { range: 'D4:G8', style: 2, numberFormat: 14 }, { range: 'E5:E6' }, { range: '12:13', style: 1 }, { range: 'J:J' }, { range: 'J1:J20', numberFormat: 49 }, { range: 'J12', style: 3 }, { range: 'B9' }],
      }],
    },
  ],
  [
    'rows and columns with a look under and over the looks',
    {
      styles: STYLES,
      sheets: [{
        name: 'L',
        cells: { A1: { value: 3 }, C3: { value: 4 }, E5: { value: 5, numberFormat: '0.0' } },
        rows: { 3: { style: 2 }, '5:6': { numberFormat: '0.000' } },
        columns: { C: { style: 1, numberFormat: '0.0%' }, 'E:F': { style: 3 } },
        looks: [{ range: 'B2:F7', style: 0 }, { range: 'C4:C6', numberFormat: '0.0' }, { range: 'F2:F7', style: 1 }],
      }],
    },
  ],
  [
    'spills, a large spill, a merged area, conditional formats and a table over the looks',
    {
      styles: STYLES,
      sheets: [{
        name: 'X',
        cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, B1: { formula: 'A1:A3*1.5' }, D1: { formula: 'A:A*2' }, H1: { value: 'k' }, I1: { value: 'v' }, H2: { value: 1 }, I2: { value: 2 } },
        looks: [{ range: 'B1:B3', numberFormat: '0.000' }, { range: 'D2:D70000', numberFormat: '0.0', style: 1 }, { range: 'F5:G6', style: 2 }, { range: 'H1:I3', style: 0 }],
        merges: ['F5:G6'],
        conditionalFormats: [{ type: 'cellIs', ranges: ['B1:B3', 'D1:D9'], operator: 'greaterThan', formulas: ['3'], style: { font: { color: '#C00000' }, numberFormat: '0.000' } }],
        tables: [{ name: 'T', range: 'H1:I3', columns: [{ name: 'k' }, { name: 'v' }] }],
      }],
    },
  ],
  ['a sheet of the wide shape, smaller', { styles: STYLES, sheets: [{ name: 'G', cells: { A1: { value: 1 }, C25: { value: 2 } }, looks: wideLooks(30, 3000) }] }],
]

/** The windows, cells and ranges each book is read by. */
const WINDOWS = ['A1:L30', 'P15:AF40', 'D1:D20', 'DHX20:DKZ60']
const CELLS = ['A1', 'B2', 'D4', 'E5', 'R21', 'S22', 'C4', 'F5', 'J12', 'J30', 'D69999', 'DIA24']
const RANGES = ['A1:Z40', 'C:E', '20:22']

/** What a session answers of a book: everything the two books are compared by. */
async function answers(s: AsyncSession): Promise<unknown> {
  const book = await s.getBook()
  const out: Record<string, unknown> = { result: await s.getResult().then((r) => [r.hash, r.inputHash]), book: expandBook(book) }
  for (const { name } of book.sheets) {
    out[`used ${name}`] = await s.getUsedRange(name)
    for (const w of WINDOWS) {
      const { revision: _, ...v } = await s.getViewport(name, w, { editText: true, diagnostics: true, formulaLocal: true })
      out[`window ${name}!${w}`] = v
    }
    out[`cells ${name}`] = await Promise.all(CELLS.map((c) => s.getCell(name, c)))
    out[`read ${name}`] = await s.readCells(RANGES.map((range) => ({ sheet: name, range })), { fields: ['value', 'text', 'style', 'numberFormat', 'formula'], maxCells: 5000 })
    out[`find ${name}`] = await s.findCells([{ sheet: name, range: 'A1:Z40' }], '%', { fields: ['style', 'numberFormat', 'text'], maxCells: 2000 })
  }
  return out
}

describe.each(['closures', 'eval'] as const)('a book with looks and the same book with entries, %s', (execution) => {
  const opts = { clock: CLOCK, execution }

  it.each(BOOKS)('%s: the same hashes', (_, book) => {
    const [a, b] = [compute(book, opts), compute(expandBook(book), opts)]
    expect([a.hash, a.inputHash]).toEqual([b.hash, b.inputHash])
  })

  describe.each(KINDS)('%s', (kind) => {
    it.each(BOOKS)('%s: the same answers, book and file', async (_, book) => {
      const open = await opener(kind)
      const [s, e] = [await open(book, opts), await open(expandBook(book), opts)]
      expect(await answers(s)).toEqual(await answers(e))
      // the file: the same parts, and its looks read back as the session holds them
      const [fs, fe] = [await writeXlsx(await s.getSnapshot()), await writeXlsx(await e.getSnapshot())]
      expect(await partsHash(fs.data)).toBe(await partsHash(fe.data))
      // a file read numbers the styles anew: a look by its style's fields
      const looksOf = (b: BookInput) => b.sheets.map((x) => x.looks?.map(({ style, ...rest }) => ({ ...rest, ...(style !== undefined ? { style: b.styles![style] } : {}) })))
      expect(looksOf((await readXlsx(fs.data)).book)).toEqual(looksOf(await s.getBook()))
      await Promise.all([s.close(), e.close()])
    }, 60_000)
  })
})

describe('the looks\' cells in a session', () => {
  const book = (): BookInput => ({ styles: STYLES, sheets: [{ name: 'S', cells: { A1: { value: 1 } }, looks: [{ range: 'B2:D4', style: 1, numberFormat: '0.0%' }, { range: '6:6', style: 2 }] }] })

  it('a cell of the looks shows its look as its own; the result lists it not, its hash is the values\'', () => {
    const s = createSession(book(), { clock: CLOCK })
    expect(s.getCell('S', 'C3')).toEqual({ value: null, text: '', style: 1, numberFormat: '0.0%', numberFormatCode: '0.0%' })
    expect(s.getViewport('S', 'C3:C3').cells.C3).toEqual({ value: null, text: '', style: 1, effectiveStyle: 1, numberFormat: '0.0%', numberFormatSource: 'cell' })
    expect(Object.keys(s.getResult().sheets[0].cells)).toEqual(['A1'])
    expect(s.getUsedRange('S')).toEqual({ rows: 6, columns: 16_384 })
    expect(s.getHash()).toBe(compute({ sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] }, { clock: CLOCK }).hash)
  })

  it('typed over, a cell keeps its look; Delete keeps it, replaced by nothing it has none; patched and cleared, it changes', () => {
    const s = createSession(book(), { clock: CLOCK })
    // read by the cell's percent, as typed into a cell of that format: its look is not a change
    const typed = s.setCellText('S', 'C3', '50')
    expect(typed.ok && typed.changes).toEqual([{ sheet: 'S', cell: 'C3', value: 0.5, text: '50.0%' }])
    expect(cellsOf(s.getBook().sheets[0]).C3).toEqual({ style: 1, numberFormat: '0.0%', value: 0.5 })
    s.setCells([{ sheet: 'S', cell: 'C3', input: null }])
    expect(cellsOf(s.getBook().sheets[0]).C3).toEqual({ style: 1, numberFormat: '0.0%' })
    expect(s.getBook().sheets[0].cells.C3).toBeUndefined()
    s.replaceCells([{ sheet: 'S', cell: 'B2', input: null }])
    expect(s.getCell('S', 'B2')).toBeUndefined()
    s.patchCells([{ sheet: 'S', range: 'D2:D6', set: { style: { font: { bold: true } } } }])
    const st = (a: string) => s.getBook().styles![cellsOf(s.getBook().sheets[0])[a].style!]
    expect([st('D2'), st('D6')]).toEqual([{ fill: { color: '#FFF2CC' }, font: { bold: true } }, { font: { italic: true, bold: true } }])
    s.clearCells([{ sheet: 'S', range: 'B3:B4' }], 'formats')
    expect([cellsOf(s.getBook().sheets[0]).B3, s.getCell('S', 'B3')]).toEqual([undefined, undefined])
    // every step undone: the book opened
    while (s.undo());
    expect(expandBook(s.getBook())).toEqual(expandBook(book()))
  })

  it('a row inserted under a row of looks takes them; a range moved takes its looks and leaves none', () => {
    const s = createSession(book(), { clock: CLOCK })
    const r = s.editStructure({ op: 'insertRows', sheet: 'S', row: 3 })
    expect(r.ok && r.changes.map((c) => [c.cell, c.style])).toEqual([['B3', 1], ['C3', 1], ['D3', 1]])
    expect(['B3', 'B5', 'B6', 'B7'].map((a) => cellsOf(s.getBook().sheets[0])[a]?.style)).toEqual([1, 1, undefined, 2])
    s.editStructure({ op: 'moveRange', from: { sheet: 'S', range: 'B2:B3' }, to: { sheet: 'S', cell: 'F2' } })
    expect(['B2', 'F2', 'F3'].map((a) => cellsOf(s.getBook().sheets[0])[a]?.numberFormat)).toEqual([undefined, '0.0%', '0.0%'])
    while (s.undo());
    expect(expandBook(s.getBook())).toEqual(expandBook(book()))
  })
})

describe('a book without looks', () => {
  it('the synthetic book keeps its hashes', () => {
    const r = compute(synthBook({ seed: 1, rows: 3400 }).book, { clock: { now: '2026-10-04T12:00:00Z', seed: 7 } })
    expect([r.hash.slice(0, 16), r.inputHash.slice(0, 16)]).toEqual(['101578a9c6b600b1', '1fb3575501a5acbf'])
  })
})
