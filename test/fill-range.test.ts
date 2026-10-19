/**
 * `fillRange` beside Excel's answers (`structure-win.test.ts`): in a session it gives the book `editStructure` gives and
 * undoes and redoes it in one step, in every kind of session; dates given as text stay text, in the 1904 system too;
 * formulas, styles and the part of the source at the far end go as copies take them; copies and series land in rows
 * hidden by hand and by a filter, as Excel's fill handle filled them (the set `hidden`, Excel for Mac and for Windows);
 * the operation's input is checked.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { ExcelStructAnswer } from '../bench/coverage/struct.js'
import { buildStructCases } from '../bench/coverage/struct/hidden.js'
import { compute, createSession, editStructure, type BookInput, type StructureOp } from '../src/index.js'
import { clock, hashes, inOtherSessions, inputErrorPath, opCases } from './struct-sessions.js'

const cases = await opCases(['fill', 'fill-held', 'fill-third-held'])

/** The cells of a book's first sheet as their values, formulas with `=`. */
const shown = (book: BookInput) => Object.fromEntries(Object.entries(book.sheets[0].cells).map(([a, c]) => [a, c.formula !== undefined ? `=${c.formula}` : c.value]))

const fill = (range: string, to: string, more: Partial<Extract<StructureOp, { op: 'fillRange' }>> = {}): StructureOp => ({ op: 'fillRange', sheet: 'S', range, to, ...more })

describe('fillRange', () => {
  it('in a session gives the book editStructure gives, and undoes and redoes it in one step', () => {
    for (const c of cases) {
      const s = createSession(c.book, { clock })
      const was = hashes(s.getBook())
      const r = s.editStructure(c.ops[0])
      expect(r.ok, c.id).toBe(true)
      expect(hashes(s.getBook()), c.id).toEqual(hashes(c.after!))
      expect(s.getHash(), c.id).toBe(compute(c.after!, { clock }).hash)
      expect(s.undo()!.call, c.id).toBe('editStructure')
      expect(hashes(s.getBook()), c.id).toEqual(was)
      s.redo()
      expect(hashes(s.getBook()), c.id).toEqual(hashes(c.after!))
    }
  })

  it('keeps a date given as text as text, in the 1904 system too', () => {
    for (const dateSystem of [1900, 1904] as const) {
      const book: BookInput = { dateSystem, sheets: [{ name: 'S', cells: { A1: { value: { date: '2025-01-31' } }, A2: { value: { date: '2025-02-28' } }, B1: { value: { date: '2025-03-01' } } } }] }
      const r = editStructure(book, fill('A1:B2', 'A1:B5'))
      if (!r.ok) throw new Error(r.message)
      expect(shown(r.book), String(dateSystem)).toEqual({
        A1: { date: '2025-01-31' },
        A2: { date: '2025-02-28' },
        A3: { date: '2025-03-31' },
        A4: { date: '2025-04-30' },
        A5: { date: '2025-05-31' },
        B1: { date: '2025-03-01' },
        B3: { date: '2025-03-02' },
        B5: { date: '2025-03-03' },
      })
    }
  })

  it('copies formulas and styles, the nearest part of the source at the far end', () => {
    const book: BookInput = { styles: [{ font: { bold: true } }], sheets: [{ name: 'S', cells: { A1: { value: 'Item1', style: 0 }, A2: { formula: 'B2*2' }, A3: { value: 7, numberFormat: '0.00' } } }] }
    const r = editStructure(book, fill('A1:A3', 'A1:A8'))
    if (!r.ok) throw new Error(r.message)
    expect(r.book.sheets[0].cells).toEqual({
      ...book.sheets[0].cells,
      A4: { value: 'Item2', style: 0 },
      A5: { formula: 'B5*2' },
      A6: { value: 8, numberFormat: '0.00' },
      A7: { value: 'Item3', style: 0 },
      A8: { formula: 'B8*2' },
    })
  })

  it('goes backwards up and left, and copies with type copy', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { C5: { value: 1 }, C6: { value: 2 }, E5: { value: 'Mon' } } }] }
    const up = editStructure(book, fill('C5:C6', 'C2:C6'))
    const left = editStructure(book, fill('E5', 'D5:E5'))
    const copy = editStructure(book, fill('C5:C6', 'C5:C9', { type: 'copy' }))
    if (!up.ok || !left.ok || !copy.ok) throw new Error('refused')
    expect(shown(up.book)).toMatchObject({ C2: -2, C3: -1, C4: 0 })
    expect(shown(left.book)).toMatchObject({ D5: 'Sun' })
    expect(shown(copy.book)).toMatchObject({ C7: 1, C8: 2, C9: 1 })
  })

  it("fills rows hidden by hand and by a filter, text, formulas and copies alike, as Excel's fill handle did", () => {
    const { books } = buildStructCases()
    for (const file of ['excel/struct/hidden.json', 'excel/win/struct/hidden.json']) {
      const answers = JSON.parse(readFileSync(new URL(file, import.meta.url), 'utf8')).answers as Record<string, ExcelStructAnswer>
      for (const book of ['manual', 'filtered'])
        for (const type of ['default', 'copy'] as const) {
          const r = editStructure(books[book], fill('D2:E2', 'D2:E16', { type }))
          const excel = answers[`${book}-fill-handle`].sheets[0].cells
          const rows = Array.from({ length: 15 }, (_, i) => i + 2)
          expect(r.ok && rows.map((n) => [shown(r.book)[`D${n}`], shown(r.book)[`E${n}`]]), `${file} ${book} ${type}`).toEqual(rows.map((n) => [excel[`D${n}`], excel[`E${n}`]]))
        }
    }
  })

  it('a series and copies of texts over a row a filter hid', () => {
    for (const [a, b, then] of [[1, 2, [3, 4, 5, 6]], ['x', 'y', ['x', 'y', 'x', 'y']]] as const) {
      const r = editStructure({ sheets: [{ name: 'S', cells: { A1: { value: a }, A2: { value: b } }, rows: { '3': { hidden: true } }, filtered: true }] }, fill('A1:A2', 'A1:A6'))
      expect(r.ok && [3, 4, 5, 6].map((n) => shown(r.book)[`A${n}`])).toEqual(then)
    }
  })

  it('checks its input', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] }
    const bad: [unknown, string][] = [
      [fill('A1:A2', 'A1:A2'), 'op.to'],
      [fill('A1:A2', 'A2:A5'), 'op.to'],
      [fill('A1:A2', 'A1:B5'), 'op.to'],
      [fill('A1', 'B2:B3'), 'op.to'],
      [fill('A1', 'A1:A'), 'op.to'],
      [fill('A1:', 'A1:A3'), 'op.range'],
      [{ op: 'fillRange', sheet: 'X', range: 'A1', to: 'A1:A3' }, 'op.sheet'],
      [{ op: 'fillRange', sheet: 'S', range: 'A1', to: 'A1:A3', type: 'days' }, 'op.type'],
      [{ op: 'fillRange', sheet: 'S', range: 'A1', to: 'A1:A3', withMeta: true }, 'op'],
    ]
    for (const [op, path] of bad) expect(inputErrorPath(book, op), JSON.stringify(op)).toBe(path)
  })
})

describe('fillRange in the other kinds of session', () => {
  inOtherSessions(cases.filter((x) => x.id.startsWith('down-numbers') || x.id.startsWith('up-')))
})
