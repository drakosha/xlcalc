/**
 * A book with looks (`SheetInput.looks`) and the same book with every cell of its looks as an entry of a look alone
 * (`expandBook`) are one book: the same hashes, windows, cells, queries, used ranges and book given back, the same xlsx
 * parts, which read back give the looks as the session holds them. In a synchronous session, in the same thread and in a
 * thread, with closures and eval (formats by code that no built-in format has, which a file read gives back by number). Then the looks' cells as a session shows and edits them, by what an entry of a look
 * alone gave.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import { compute, createSession, type AsyncSession, type BookInput, type LookRange } from '../src/index.js'
import { formatAddress } from '../src/model.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { CLOCK } from './async-books.js'
import { cellsOf, expandBook, seeded } from './looks-book.js'
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

/** A session's hashes and book with its looks as entries, and the book. */
async function hashed(s: AsyncSession): Promise<[Record<string, unknown>, BookInput]> {
  const book = await s.getBook()
  return [{ result: await s.getResult().then((r) => [r.hash, r.inputHash]), book: expandBook(book) }, book]
}

/** The files of two sessions written: the same parts. The first's data. */
async function sameFiles(s: AsyncSession, e: AsyncSession): Promise<Uint8Array> {
  const [fs, fe] = [await writeXlsx(await s.getSnapshot()), await writeXlsx(await e.getSnapshot())]
  expect(await partsHash(fs.data)).toBe(await partsHash(fe.data))
  return fs.data
}

/** What a session answers of a book: everything the two books are compared by. */
async function answers(s: AsyncSession): Promise<unknown> {
  const [out, book] = await hashed(s)
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

/**
 * The answers and the file of the expanded books, taken once: a session of any kind and execution mode opens them the
 * same, and the tests below check each kind's session with the looks against this.
 */
const ENTRIES = new Map<string, { answers: unknown; fileHash: string }>()

beforeAll(async () => {
  const open = await opener('a synchronous session')
  for (const [name, book] of BOOKS) {
    const e = await open(expandBook(book), { clock: CLOCK, execution: 'closures' })
    ENTRIES.set(name, { answers: await answers(e), fileHash: await partsHash((await writeXlsx(await e.getSnapshot())).data) })
    await e.close()
  }
}, 60_000)

describe.each(['closures', 'eval'] as const)('a book with looks and the same book with entries, %s', (execution) => {
  const opts = { clock: CLOCK, execution }

  it.each(BOOKS)('%s: the same hashes', (_, book) => {
    const [a, b] = [compute(book, opts), compute(expandBook(book), opts)]
    expect([a.hash, a.inputHash]).toEqual([b.hash, b.inputHash])
  })

  describe.each(KINDS)('%s', (kind) => {
    it.each(BOOKS)('%s: the same answers, book and file', async (name, book) => {
      const open = await opener(kind)
      const s = await open(book, opts)
      const entries = ENTRIES.get(name)!
      expect(await answers(s)).toEqual(entries.answers)
      // the file: the same parts, and its looks read back as the session holds them
      const data = (await writeXlsx(await s.getSnapshot())).data
      expect(await partsHash(data)).toBe(entries.fileHash)
      // a file read numbers the styles anew: a look by its style's fields
      const looksOf = (b: BookInput) => b.sheets.map((x) => x.looks?.map(({ style, ...rest }) => ({ ...rest, ...(style !== undefined ? { style: b.styles![style] } : {}) })))
      expect(looksOf((await readXlsx(data)).book)).toEqual(looksOf(await s.getBook()))
      await s.close()
    }, 60_000)
  })
})

/** What a session answers of a book after an edit, the lighter way: hashes, the book, a window, the used range, cells read. */
async function after(s: AsyncSession): Promise<unknown> {
  const [out, book] = await hashed(s)
  for (const { name } of book.sheets) {
    out[`used ${name}`] = await s.getUsedRange(name)
    const { revision: _, ...v } = await s.getViewport(name, 'A1:Z60', { editText: true })
    out[`window ${name}`] = v
    out[`read ${name}`] = await s.readCells([{ sheet: name, range: 'A15:Z60' }], { fields: ['value', 'text', 'style', 'numberFormat'], maxCells: 2000 })
  }
  return out
}

/** An answer without what a session gives as a function (`mapCell`). */
const plain = (x: unknown) => (x && typeof x === 'object' ? Object.fromEntries(Object.entries(x).filter(([k]) => k !== 'mapCell')) : x)

/** Edits of the looks and over them on the book of patterns: each is taken, undone and done again. */
const EDITS: [string, (s: AsyncSession) => Promise<unknown>][] = [
  ['a patch of a rectangle over the looks', (s) => s.patchCells([{ sheet: 'G', range: 'C19:T24', set: { style: { font: { italic: true } } } }])],
  ['a patch of whole rows', (s) => s.patchCells([{ sheet: 'G', range: '25:26', set: { numberFormat: '0.000' } }])],
  ['a patch of whole columns', (s) => s.patchCells([{ sheet: 'G', range: 'R:S', set: { style: 3 } }])],
  ['formats cleared', (s) => s.clearCells([{ sheet: 'G', range: 'E20:F31' }], 'formats')],
  ['everything cleared', (s) => s.clearCells([{ sheet: 'G', range: 'Q19:S22' }], 'all')],
  ['the style table replaced', async (s) => s.setStyles([...(await s.getBook()).styles!, { font: { color: '#123456' } }])],
  ['a text typed into a cell of the looks', (s) => s.setCellText('G', 'R25', '12')],
  ['a value set in a cell of the looks', (s) => s.setCells([{ sheet: 'G', cell: 'D30', input: { value: 5 } }])],
  ['a cell of the looks deleted', (s) => s.setCells([{ sheet: 'G', cell: 'E25', input: null }])],
  ['a cell of the looks replaced by nothing', (s) => s.replaceCells([{ sheet: 'G', cell: 'E31', input: null }])],
  ['rows inserted inside the looks', (s) => s.editStructure({ op: 'insertRows', sheet: 'G', row: 21, count: 2 })],
  ['rows inserted at their edge', (s) => s.editStructure({ op: 'insertRows', sheet: 'G', row: 20 })],
  ['rows inserted under a row of looks', (s) => s.editStructure({ op: 'insertRows', sheet: 'G', row: 23 })],
  ['rows inserted before them', (s) => s.editStructure({ op: 'insertRows', sheet: 'G', row: 5, count: 3 })],
  ['a row deleted inside', (s) => s.editStructure({ op: 'deleteRows', sheet: 'G', row: 21 })],
  ['rows deleted through them', (s) => s.editStructure({ op: 'deleteRows', sheet: 'G', row: 18, count: 12 })],
  ['columns inserted inside', (s) => s.editStructure({ op: 'insertColumns', sheet: 'G', column: 'S', count: 2 })],
  ['a column inserted before them', (s) => s.editStructure({ op: 'insertColumns', sheet: 'G', column: 'A' })],
  ['columns deleted through them', (s) => s.editStructure({ op: 'deleteColumns', sheet: 'G', column: 'Q', count: 5 })],
  ['a rectangle of looks copied', (s) => s.editStructure({ op: 'copyRange', from: { sheet: 'G', range: 'C19:T23' }, to: { sheet: 'G', range: 'C40:T44' } })],
  ['entries copied over looks', (s) => s.editStructure({ op: 'copyRange', from: { sheet: 'G', range: 'A1:C3' }, to: { sheet: 'G', range: 'Q20' } })],
  ['a rectangle of looks moved to another sheet', (s) => s.editStructure({ op: 'moveRange', from: { sheet: 'G', range: 'D19:H22' }, to: { sheet: 'Data', cell: 'J50' } })],
  ['entries moved over looks', (s) => s.editStructure({ op: 'moveRange', from: { sheet: 'G', range: 'A1:B2' }, to: { sheet: 'G', cell: 'R20' } })],
]

describe.each(['closures', 'eval'] as const)('edits of a book with looks and of the same book with entries, %s', (execution) => {
  const opts = { clock: CLOCK, execution }
  describe.each(KINDS)('%s', (kind) => {
    it('each edit, its undo and its redo: the same answers', async () => {
      const open = await opener(kind)
      const book = BOOKS[0][1]
      const [s, e] = [await open(book, opts), await open(expandBook(book), opts)]
      for (const [name, edit] of EDITS) {
        expect(plain(await edit(s)), name).toEqual(plain(await edit(e)))
        expect(await after(s), name).toEqual(await after(e))
        expect(plain(await s.undo()), `${name}, undone`).toEqual(plain(await e.undo()))
        expect(await after(s), `${name}, undone`).toEqual(await after(e))
        expect(plain(await s.redo()), `${name}, done again`).toEqual(plain(await e.redo()))
      }
      expect(await after(s)).toEqual(await after(e))
      await sameFiles(s, e)
      // every edit undone: the book opened
      while (await s.undo());
      expect(expandBook(await s.getBook())).toEqual(expandBook(book))
      await Promise.all([s.close(), e.close()])
    }, 120_000)
  })
})

/** A random edit of the book of patterns: of looks, of values over them, of its lines and ranges, an undo or a redo. */
function randomEdit(rnd: (n: number) => number): [string, (s: AsyncSession) => Promise<unknown>] {
  const col = () => formatAddress(0, rnd(26)).replace(/\d+$/, '')
  const cell = () => formatAddress(14 + rnd(30), rnd(26))
  const rect = () => {
    const [r, c] = [14 + rnd(30), rnd(24)]
    return `${formatAddress(r, c)}:${formatAddress(r + rnd(6), c + rnd(5))}`
  }
  const sheet = 'G'
  const kinds: [string, (s: AsyncSession) => Promise<unknown>][] = [
    ['patch', (s) => s.patchCells([{ sheet, range: rnd(4) ? rect() : `${16 + rnd(20)}:${17 + rnd(20)}`, set: rnd(2) ? { style: { fill: { color: '#C6EFCE' } } } : { numberFormat: '0.00' } }])],
    ['clear', (s) => s.clearCells([{ sheet, range: rect() }], (['formats', 'all', 'contents'] as const)[rnd(3)])],
    ['value', (s) => s.setCells([{ sheet, cell: cell(), input: rnd(5) ? { value: rnd(100) } : null }])],
    ['text', (s) => s.setCellText(sheet, cell(), String(rnd(1000) / 10))],
    ['nothing', (s) => s.replaceCells([{ sheet, cell: cell(), input: null }])],
    ['rows', (s) => s.editStructure({ op: rnd(2) ? 'insertRows' : 'deleteRows', sheet, row: 1 + rnd(45), count: 1 + rnd(3) })],
    ['columns', (s) => s.editStructure({ op: rnd(2) ? 'insertColumns' : 'deleteColumns', sheet, column: col(), count: 1 + rnd(2) })],
    ['copy', (s) => s.editStructure({ op: 'copyRange', from: { sheet, range: rect() }, to: { sheet, range: cell() } })],
    ['move', (s) => s.editStructure({ op: 'moveRange', from: { sheet, range: rect() }, to: { sheet, cell: cell() } })],
    ['undo', (s) => s.undo()],
    ['redo', (s) => s.redo()],
  ]
  return kinds[rnd(kinds.length)]
}

describe.each(['closures', 'eval'] as const)('a random series of edits of a book with looks, %s', (execution) => {
  const opts = { clock: CLOCK, execution }
  describe.each(KINDS)('%s', (kind) => {
    it('after each step: the session\'s hashes a full computation\'s, its windows those of the book it gives', async () => {
      const open = await opener(kind)
      const book = BOOKS[0][1]
      const s = await open(book, opts)
      const rnd = seeded(execution === 'eval' ? 17 : 5)
      for (let k = 0; k < 40; k++) {
        const [name, edit] = randomEdit(rnd)
        await edit(s)
        const now = await s.getBook()
        const full = compute(now, opts)
        const at = `step ${k}: ${name}`
        expect([await s.getHash(), (await s.getResult()).inputHash], at).toEqual([full.hash, full.inputHash])
        // the book taken again from its entries: the same looks
        const fresh = createSession(expandBook(now), opts)
        expect(await s.getUsedRange('G'), at).toEqual(fresh.getUsedRange('G'))
        const { revision: _a, ...a } = await s.getViewport('G', 'A10:AB50')
        const { revision: _b, ...b } = fresh.getViewport('G', 'A10:AB50')
        expect(a, at).toEqual(b)
        fresh.close()
      }
      while (await s.undo());
      expect(expandBook(await s.getBook())).toEqual(expandBook(book))
      await s.close()
    }, 120_000)
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
