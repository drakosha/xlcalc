/**
 * Notes, hyperlinks and frozen panes in the model: checked as input, taking no part in computation or the values hash,
 * in `inputHash` only when a book has them, moved by structure operations (the rules measured in `test/excel/parts/`),
 * and changed through a session without recomputing, in both execution modes.
 */
import { describe, expect, it } from 'vitest'
import { compute, InputError, createSession, editStructure, type BookInput, type Execution, type SheetInput, type StructureOp } from '../src/index.js'
import { sameValuesApart } from './formula-values.js'

const clock = { now: '2026-10-10T12:00:00Z', seed: 1 }

/** S: a sum over A1:A3, a note on A1, a link on B2, a note and a link on an empty C3, rows 1:2 and column A frozen; T: a link to S. */
const book = (): BookInput => ({
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 1, note: { author: 'Ann', text: 'first' } },
        A2: { value: 2 },
        A3: { value: 3 },
        B1: { formula: 'SUM(A1:A3)' },
        B2: { value: 'site', link: { url: 'https://example.com/', tooltip: 'go' } },
        C3: { note: { text: 'alone' }, link: { location: 'S!A1' } },
      },
      frozen: { rows: 2, columns: 1 },
    },
    { name: 'T', cells: { A1: { value: 'back', link: { location: "'S'!B1" } } } },
  ],
})

/** The book without its parts: a cell that held only them is an empty cell, listed in a result as a cell with only a style is. */
function bare(b: BookInput): BookInput {
  return {
    ...b,
    sheets: b.sheets.map(({ frozen: _, ...s }): SheetInput => ({ ...s, cells: Object.fromEntries(Object.entries(s.cells).map(([a, { note: __, link: ___, ...c }]) => [a, c])) })),
  }
}

const parts = (b: BookInput, sheet = 'S') => {
  const s = b.sheets.find((x) => x.name === sheet)!
  return {
    notes: Object.fromEntries(Object.entries(s.cells).flatMap(([a, c]) => (c.note ? [[a, c.note.text]] : []))),
    links: Object.fromEntries(Object.entries(s.cells).flatMap(([a, c]) => (c.link ? [[a, c.link.url ?? c.link.location]] : []))),
    frozen: s.frozen,
  }
}

const after = (op: StructureOp) => {
  const r = editStructure(book(), op)
  if (!r.ok) throw new Error(r.message)
  return r.book
}

describe('notes, links and frozen panes as input', () => {
  it('wrong shapes throw InputError with the path', () => {
    const fails = (cells: SheetInput['cells'], frozen?: unknown, path?: string) => {
      const e = (() => {
        try {
          compute({ sheets: [{ name: 'S', cells, ...(frozen !== undefined ? { frozen } : {}) } as SheetInput] })
        } catch (x) {
          return x
        }
      })()
      expect(e).toBeInstanceOf(InputError)
      if (path) expect((e as Error).message).toContain(path)
    }
    fails({ A1: { note: { author: 'Ann' } as never } }, undefined, 'cells["A1"].note.text')
    fails({ A1: { note: { text: 'x', by: 'Ann' } as never } }, undefined, 'unknown field "by"')
    fails({ A1: { link: {} } }, undefined, 'a hyperlink needs a url, a location or both')
    fails({ A1: { link: { url: 'x', tip: 'y' } as never } }, undefined, 'unknown field "tip"')
    fails({ A1: { link: { location: 5 } as never } }, undefined, 'cells["A1"].link.location')
    fails({}, {}, 'expected rows, columns or both')
    fails({}, { rows: 0 }, 'frozen.rows')
    fails({}, { columns: 16_384 }, 'frozen.columns')
    fails({}, { rows: 1.5 }, 'frozen.rows')
  })

  it('values and the values hash are those of the book without them; inputHash takes them in only when there are any', () => {
    const r = sameValuesApart(book(), bare(book()), clock)
    // a book without them keeps the inputHash it had before they existed: recorded with master e1383d4
    const old: BookInput = { styles: [{ font: { bold: true } }], sheets: [{ name: 'S', cells: { A1: { value: 1, style: 0 }, A2: { value: 2 }, A3: { value: 3 }, B1: { formula: 'SUM(A1:A3)' }, B2: { value: 'site' }, C3: {} }, merges: ['D1:E1'] }, { name: 'T', cells: { A1: { value: 'back' } } }] }
    expect(compute(old, { clock }).inputHash).toBe('0cfa5b2caae351144eb19b829ac9148095834a251d6c6250abbe5643d067802d')
    const noFrozen = { ...book(), sheets: book().sheets.map(({ frozen: _, ...s }) => s) }
    expect(compute(noFrozen, { clock }).inputHash).not.toBe(r.inputHash)
  })
})

describe('notes, links and frozen panes under structure operations', () => {
  it('rows inserted above the split push it on, at the split leave it; notes and links move with their cells, inserted rows get none', () => {
    expect(parts(after({ op: 'insertRows', sheet: 'S', row: 1 }))).toEqual({ notes: { A2: 'first', C4: 'alone' }, links: { B3: 'https://example.com/', C4: 'S!A1' }, frozen: { rows: 3, columns: 1 } })
    expect(parts(after({ op: 'insertRows', sheet: 'S', row: 3, count: 2 }))).toEqual({ notes: { A1: 'first', C5: 'alone' }, links: { B2: 'https://example.com/', C5: 'S!A1' }, frozen: { rows: 2, columns: 1 } })
  })

  it('deleted rows and columns take their notes and links; the split draws back, one line staying frozen at least', () => {
    expect(parts(after({ op: 'deleteRows', sheet: 'S', row: 1 }))).toEqual({ notes: { C2: 'alone' }, links: { B1: 'https://example.com/', C2: 'S!A1' }, frozen: { rows: 1, columns: 1 } })
    expect(parts(after({ op: 'deleteRows', sheet: 'S', row: 1, count: 2 })).frozen).toEqual({ rows: 1, columns: 1 })
    expect(parts(after({ op: 'deleteColumns', sheet: 'S', column: 'A', count: 2 }))).toEqual({ notes: { A3: 'alone' }, links: { A3: 'S!A1' }, frozen: { rows: 2, columns: 1 } })
    expect(parts(after({ op: 'insertColumns', sheet: 'S', column: 'A', count: 3 })).frozen).toEqual({ rows: 2, columns: 4 })
  })

  it('a move takes them along and clears the cells it lands on; a copy copies them and replaces the target cells', () => {
    expect(parts(after({ op: 'moveRange', from: { sheet: 'S', range: 'A1:B2' }, to: { sheet: 'S', cell: 'B3' } }))).toMatchObject({ notes: { B3: 'first' }, links: { C4: 'https://example.com/' } })
    expect(parts(after({ op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', range: 'C3' } }))).toMatchObject({ notes: { A1: 'first', C3: 'first' }, links: { B2: 'https://example.com/' } })
    expect(parts(after({ op: 'copyRange', from: { sheet: 'S', range: 'A2' }, to: { sheet: 'S', range: 'C3' } }))).toMatchObject({ notes: { A1: 'first' }, links: { B2: 'https://example.com/' } })
    expect(parts(after({ op: 'moveRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'T', cell: 'D4' } }), 'T').notes).toEqual({ D4: 'first' })
  })

  it("a renamed or deleted sheet leaves links' locations as written, as Excel does", () => {
    expect(parts(after({ op: 'renameSheet', sheet: 'S', name: 'Data' }), 'T').links).toEqual({ A1: "'S'!B1" })
    expect(parts(after({ op: 'renameSheet', sheet: 'S', name: 'Data' }), 'Data').frozen).toEqual({ rows: 2, columns: 1 })
    expect(parts(after({ op: 'deleteSheet', sheet: 'S' }), 'T').links).toEqual({ A1: "'S'!B1" })
  })
})

describe.each(['closures', 'eval'] as const)('notes, links and frozen panes in a session, execution %s', (execution: Execution) => {
  it('setCells changes them without recomputing; cells() gives them; the hash stays a full computation of the book', () => {
    const s = createSession(book(), { clock, execution })
    const before = s.getHash()
    // listed with the link and the note they now have, their values as they were
    expect(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 1, link: { url: 'mailto:a@example.com' } } }, { sheet: 'S', cell: 'D9', input: { note: { text: 'new' } } }]).changes).toEqual([
      { sheet: 'S', cell: 'A1', value: 1, text: '1', link: { url: 'mailto:a@example.com' } },
      { sheet: 'S', cell: 'D9', value: null, text: '', note: { text: 'new' } },
    ])
    expect(s.getCells('S', 'A1:D9')).toMatchObject({ A1: { value: 1, link: { url: 'mailto:a@example.com' } }, C3: { value: null, note: { text: 'alone' }, link: { location: 'S!A1' } }, D9: { value: null, note: { text: 'new' } } })
    // the edit keeps the note it does not name
    expect(s.getCells('S', 'A1').A1.note).toEqual({ author: 'Ann', text: 'first' })
    expect(s.getHash()).toBe(before)
    expect(compute(s.getBook(), { clock: s.clock, execution }).hash).toBe(before)
    s.close()
  })

  it('setLayout freezes and unfreezes panes; a editStructure moves them; nothing is recomputed', () => {
    const s = createSession(book(), { clock, execution })
    expect(s.setLayout([{ sheet: 'T', frozen: { columns: 2 } }, { sheet: 'S', frozen: null }]).changes).toEqual([])
    expect(s.getBook().sheets.map((x) => x.frozen)).toEqual([undefined, { columns: 2 }])
    expect('frozen' in s.getBook().sheets[0]).toBe(false)
    const r = s.editStructure({ op: 'insertColumns', sheet: 'T', column: 'A' })
    expect(r.ok && s.getBook().sheets[1].frozen).toEqual({ columns: 3 })
    expect(() => s.setLayout([{ sheet: 'S', frozen: { rows: -1 } }])).toThrow(InputError)
    s.close()
  })
})
