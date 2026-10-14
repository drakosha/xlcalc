/**
 * `mergeCells` and `unmergeCells` in sessions: a session gives the book `editStructure` gives, with or without building
 * it again (`session/merged.ts`), in one step of the undo, in every kind of session; the cell moved into an empty
 * top-left cell is where `mapCell` says; the operations' input is checked. Excel's answers are in `structure-win.test.ts`.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, editStructure, registerLocale, type BookInput } from '../src/index.js'
import { clock, hashes, inOtherSessions, inputErrorPath, opCases } from './struct-sessions.js'

const cases = await opCases(['merge', 'merge-held', 'merge-third-held', 'merge-top-held'])

describe('mergeCells and unmergeCells in a session', () => {
  it('give the book editStructure gives, and undo and redo each in one step', () => {
    for (const c of cases) {
      const s = createSession(c.book, { clock })
      const was = hashes(s.getBook())
      for (const op of c.ops) s.editStructure(op)
      if (!c.after) continue
      expect(hashes(s.getBook()), c.id).toEqual(hashes(c.after))
      expect(s.getHash(), c.id).toBe(compute(c.after, { clock }).hash)
      const after = hashes(s.getBook())
      expect(s.undo()!.call, c.id).toBe('editStructure')
      if (c.ops.length === 1) expect(hashes(s.getBook()), c.id).toEqual(was)
      s.redo()
      expect(hashes(s.getBook()), c.id).toEqual(after)
    }
  })

  it('list the cells they clear and move, and map the moved cell to the top-left one', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { B1: { value: 2 }, C1: { value: 3 }, D1: { formula: 'B1*10' } } }] }
    const s = createSession(book, { clock })
    const r = s.editStructure({ op: 'mergeCells', sheet: 'S', range: 'A1:C1' })
    if (!r.ok) throw new Error(r.message)
    expect(r.mapCell({ sheet: 'S', cell: 'B1' })).toEqual({ sheet: 'S', cell: 'A1' })
    expect(s.getBook().sheets[0].cells).toEqual({ A1: { value: 2 }, D1: { formula: 'A1*10' } })
    // by the cells' new places: B1 came to A1 with its value, C1 is cleared, D1 reads A1 now
    expect(r.changes.map((x) => `${x.cell}=${x.value}`).sort()).toEqual(['C1=null', 'D1=20'])
    const plain = s.editStructure({ op: 'mergeCells', sheet: 'S', range: 'E1:F1', center: true })
    if (!plain.ok) throw new Error(plain.message)
    expect(plain.mapCell({ sheet: 'S', cell: 'F1' })).toEqual({ sheet: 'S', cell: 'F1' })
    expect(s.getBook().sheets[0].merges).toEqual(['A1:C1', 'E1:F1'])
    expect(s.getLayout().sheets[0].merges).toEqual(['A1:C1', 'E1:F1'])
  })

  it('check their input', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: {} }] }
    const bad: [unknown, string][] = [
      [{ op: 'mergeCells', sheet: 'S', range: 'A1:B2', across: 'yes' }, 'op.across'],
      [{ op: 'mergeCells', sheet: 'S', range: 'A1:B2', center: 1 }, 'op.center'],
      [{ op: 'mergeCells', sheet: 'S', range: 'A1:B' }, 'op.range'],
      [{ op: 'mergeCells', sheet: 'X', range: 'A1:B2' }, 'op.sheet'],
      [{ op: 'unmergeCells', sheet: 'S', range: 'A1', across: true }, 'op'],
    ]
    for (const [op, path] of bad) expect(inputErrorPath(book, op), JSON.stringify(op)).toBe(path)
  })

  it('refuse in the language of the book', async () => {
    registerLocale((await import('../src/locales/ru-RU.js')).ruRU)
    const book: BookInput = { locale: 'ru-RU', sheets: [{ name: 'S', cells: { A1: { value: 'h' }, A2: { value: 1 } }, tables: [{ name: 'Sales', range: 'A1:A2', columns: [{ name: 'h' }] }] }] }
    const r = editStructure(book, { op: 'mergeCells', sheet: 'S', range: 'A2:B2' })
    expect(r).toMatchObject({ ok: false, code: 'table', message: 'Диапазон A2:B2 пересекается с таблицей «Sales», её ячейки нельзя объединить' })
  })
})

describe('mergeCells and unmergeCells in the other kinds of session', () => {
  inOtherSessions(cases.filter((x) => x.after && x.id.startsWith('merge-A1-B2')))
})
