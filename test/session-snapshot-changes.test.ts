import { readFileSync } from 'node:fs'
import { afterAll, describe, expect, it } from 'vitest'
import { compute, parseEntry, type AsyncSession, type BookInput, type CellChange, type EditResult } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const note = { author: 'Ann', text: 'before' }
const link = { url: 'https://example.com/before', location: 'S!A1', tooltip: 'before' }
const nextNote = { text: 'after' }
const nextLink = { location: 'S!B1', tooltip: 'after' }

const book = (): BookInput => ({
  styles: [{ font: { bold: true }, fill: { color: '#FFFF00' } }],
  names: [{ name: 'Rate', ref: "'S'!A1" }],
  sheets: [{
    name: 'S',
    cells: {
      A1: { value: 2, style: 0, note: { ...note }, link: { ...link }, meta: { nested: [1, { id: 7 }] } },
      B1: { formula: 'A1*10' },
      C1: { value: { date: '2026-01-01' } },
    },
    columns: { A: { width: 12 } },
    rows: { 1: { height: 20 } },
    frozen: { rows: 1 },
    conditionalFormats: [{ type: 'cellIs', ranges: ['A1'], operator: 'greaterThan', formulas: ['1'], style: { font: { italic: true } } }],
    validations: [{ range: 'A1', type: 'whole', formula1: '0', formula2: '20' }],
  }],
})

/** Only note and link fields, so value and display changes do not hide a missing notification. */
const attachments = (result: EditResult) => result.changes.map(({ sheet, cell, note, link }: CellChange) => ({ sheet, cell, ...(note !== undefined ? { note } : {}), ...(link !== undefined ? { link } : {}) }))

afterAll(closeCalcs)

describe.each(['closures', 'eval'] as const)('snapshot and cell attachments, %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    const open = openerFor(kind)
    const session = (input = book()) => open(input, { clock, execution })

    it('a snapshot stays at its original book and result through edits, history and closing', async () => {
      const s = await session()
      const saved = await s.getSnapshot()
      const before = structuredClone(saved)
      await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
      await s.patchCells([{ sheet: 'S', cell: 'A1', set: { note: nextNote, link: nextLink } }])
      await s.setLayout([{ sheet: 'S', columns: 'A', set: { width: 18 } }])
      await s.setStyles([{ font: { italic: true } }])
      await s.setNames([{ name: 'Rate', ref: "'S'!B1" }])
      expect((await s.editStructure({ op: 'insertRows', sheet: 'S', row: 1 })).ok).toBe(true)
      await s.undo()
      await s.redo()
      await s.close()
      expect(saved).toEqual(before)
      expect(compute(saved.book, { clock, execution }).hash).toBe(saved.result.hash)
    })

    it('changing nested objects of a snapshot cannot change the session', async () => {
      const s = await session()
      try {
        const saved = await s.getSnapshot()
        saved.book.styles![0].font!.bold = false
        saved.book.sheets[0].cells.A1.note!.text = 'snapshot only'
        saved.book.sheets[0].cells.A1.link!.url = 'https://example.com/snapshot'
        saved.book.sheets[0].cells.C1.value = { date: '2026-02-01' }
        saved.book.sheets[0].columns!.A.width = 1
        saved.book.sheets[0].rows!['1'].height = 1
        saved.book.sheets[0].frozen!.rows = 2
        saved.book.sheets[0].conditionalFormats![0].ranges[0] = 'B1'
        saved.book.sheets[0].validations![0].formula1 = '10'
        saved.book.names![0].ref = "'S'!B1"
        const meta = saved.result.sheets[0].cells.A1.meta as { nested: [number, { id: number }] }
        meta.nested[1].id = 99
        expect(await s.getBook()).toEqual(book())
        expect(s.revision).toBe(0)
      } finally { await s.close() }
    })

    it('exporting a retained snapshot writes its original inputs and computed values', async () => {
      const s = await session()
      try {
        const saved = await s.getSnapshot()
        await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
        const exported = await readXlsx((await writeXlsx(saved)).data)
        expect(exported.book.sheets[0].cells.A1.value).toBe(2)
        expect(exported.cached.sheets[0].cells.B1.value).toBe(20)
      } finally { await s.close() }
    })

    it('snapshot copying preserves non-finite numbers, negative zero and date inputs', async () => {
      const input: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: NaN }, A2: { value: Infinity }, A3: { value: -0 }, A4: { value: { date: '1900-02-29' } } } }] }
      const s = await session(input)
      try {
        const saved = await s.getSnapshot()
        expect(saved.book).toEqual(input)
        expect(Object.is(saved.book.sheets[0].cells.A3.value, -0)).toBe(true)
        expect(compute(saved.book, { clock, execution }).hash).toBe(saved.result.hash)
      } finally { await s.close() }
    })

    const writes: [string, (s: AsyncSession) => Promise<EditResult>, object][] = [
      ['setCells', (s) => s.setCells([{ sheet: 'S', cell: 'A1', input: { note: nextNote, link: nextLink } }]), { note: nextNote, link: nextLink }],
      ['replaceCells', (s) => s.replaceCells([{ sheet: 'S', cell: 'A1', input: { value: 2, style: 0, note: nextNote, link: nextLink } }]), { note: nextNote, link: nextLink }],
      ['patchCells', (s) => s.patchCells([{ sheet: 'S', cell: 'A1', set: { note: nextNote, link: nextLink } }]), { note: nextNote, link: nextLink }],
      ['clearCells contents', (s) => s.clearCells([{ sheet: 'S', range: 'A1' }], 'contents'), { link: null }],
      ['clearCells all', (s) => s.clearCells([{ sheet: 'S', range: 'A1' }], 'all'), { note: null, link: null }],
    ]
    it.each(writes)('%s reports attachments, undo restores them, redo reports them again', async (_, write, expected) => {
      const s = await session()
      try {
        const result = await write(s)
        expect(attachments(result).find((c) => c.cell === 'A1')).toEqual({ sheet: 'S', cell: 'A1', ...expected })
        const undone = (await s.undo())!
        const restored = Object.fromEntries(Object.keys(expected).map((field) => [field, field === 'note' ? note : link]))
        expect(attachments(undone).find((c) => c.cell === 'A1')).toEqual({ sheet: 'S', cell: 'A1', ...restored })
        expect(attachments((await s.redo())!).find((c) => c.cell === 'A1')).toEqual({ sheet: 'S', cell: 'A1', ...expected })
      } finally { await s.close() }
    })

    it.each(['setCells', 'replaceCells', 'patchCells'] as const)('%s reports additions and removals on an otherwise empty cell', async (method) => {
      const s = await session({ sheets: [{ name: 'S', cells: {} }] })
      try {
        const add = method === 'patchCells'
          ? await s.patchCells([{ sheet: 'S', cell: 'D9', set: { note, link } }])
          : await s[method]([{ sheet: 'S', cell: 'D9', input: { note, link } }])
        expect(attachments(add)).toEqual([{ sheet: 'S', cell: 'D9', note, link }])
        expect(attachments((await s.undo())!)).toEqual([{ sheet: 'S', cell: 'D9', note: null, link: null }])
        expect(attachments((await s.redo())!)).toEqual([{ sheet: 'S', cell: 'D9', note, link }])
        const remove = method === 'patchCells'
          ? await s.patchCells([{ sheet: 'S', cell: 'D9', set: { note: null, link: null } }])
          : await s[method]([{ sheet: 'S', cell: 'D9', input: null }])
        const expected = method === 'setCells' ? { link: null } : { note: null, link: null }
        expect(attachments(remove)).toEqual([{ sheet: 'S', cell: 'D9', ...expected }])
      } finally { await s.close() }
    })

    it('equal notes and links in a different property order are not changes; meta stays host data', async () => {
      const s = await session()
      try {
        const result = await s.setCells([{ sheet: 'S', cell: 'A1', input: { note: { text: note.text, author: note.author }, link: { tooltip: link.tooltip, location: link.location, url: link.url }, meta: null } }])
        expect(result.changes).toEqual([])
        expect((await s.getCell('S', 'A1'))!.meta).toBeUndefined()
        expect((await s.undo())!.changes).toEqual([])
        expect((await s.getCell('S', 'A1'))!.meta).toEqual(book().sheets[0].cells.A1.meta)
        expect((await s.clearCells([{ sheet: 'S', range: 'A1' }], 'formats')).changes.every((c) => c.note === undefined && c.link === undefined)).toBe(true)
      } finally { await s.close() }
    })

    it('attachments are reported with display disabled and their returned objects are detached', async () => {
      const s = await open(book(), { clock, execution, display: false })
      try {
        const result = await s.patchCells([{ sheet: 'S', cell: 'A1', set: { note: nextNote, link: nextLink } }])
        expect(result.changes).toEqual([{ sheet: 'S', cell: 'A1', value: 2, note: nextNote, link: nextLink }])
        result.changes[0].note!.text = 'response only'
        result.changes[0].link!.tooltip = 'response only'
        const cell = (await s.getCell('S', 'A1'))!
        expect([cell.note, cell.link]).toEqual([nextNote, nextLink])
      } finally { await s.close() }
    })

    it('a deleted sheet put back by undo reports its notes and links, a cell of a note alone too', async () => {
      const s = await session({ sheets: [{ name: 'D', cells: { A1: { note }, B1: { value: 1, note }, C1: { value: 2, link } } }, { name: 'S', cells: { A1: { formula: 'SUM(D!B1:C1)' } } }] })
      try {
        expect((await s.editStructure({ op: 'deleteSheet', sheet: 'D' })).ok).toBe(true)
        const back = [{ sheet: 'D', cell: 'A1', note }, { sheet: 'D', cell: 'B1', note }, { sheet: 'D', cell: 'C1', link }]
        expect(attachments((await s.undo())!).filter((c) => c.sheet === 'D')).toEqual(back)
        await s.redo()
        expect(attachments((await s.undo())!).filter((c) => c.sheet === 'D')).toEqual(back)
      } finally { await s.close() }
    })

    it.each(['row', 'column'] as const)('the documented entry context accounts for an inherited %s format without changing parsing', async (source) => {
      const readme = readFileSync(new URL('../docs/entry.md', import.meta.url), 'utf8')
      const contract = readFileSync(new URL('../src/contract.ts', import.meta.url), 'utf8')
      expect(contract.slice(contract.indexOf('export interface EntryOptions'), contract.indexOf('export type EntryResult'))).toContain('getViewport')
      expect(readme).toContain("const shown = session.getViewport('S', 'A1').cells.A1")
      const s = await session({ sheets: [{ name: 'S', cells: { A1: { value: 0.1 } }, ...(source === 'row' ? { rows: { 1: { numberFormat: '0%' } } } : { columns: { A: { numberFormat: '0%' } } }) }] })
      try {
        const own = (await s.getCell('S', 'A1'))!
        const shown = (await s.getViewport('S', 'A1')).cells.A1
        expect(parseEntry('5', { cell: own })).toEqual({ ok: true, input: { value: 5 } })
        expect(parseEntry('5', { cell: shown })).toEqual({ ok: true, input: { value: 0.05, numberFormat: '0%' } })
        expect(await s.previewCellText('S', 'A1', '5')).toMatchObject({ ok: true, input: { value: 0.05 } })
      } finally { await s.close() }
    })
  })
})
