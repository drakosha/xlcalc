/**
 * Hidden sheets: `SheetInput.hidden` (true, or 'very'), the structure operations hideSheet and unhideSheet, the last
 * visible sheet that neither a delete nor a hide takes, values that do not depend on what is hidden, and the xlsx state
 * and active tab, in the synchronous session, in the same thread and in the pool.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createInlineCalc, InputError, openSession, restructure, type BookInput, type Calc, type StructureOp } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { pool } from './fixtures/pool.js'

const CLOCK = { now: '2026-10-10T12:00:00Z', seed: 1 }

/** Jan, Feb and Mar with B2:B5, a sheet of formulas over them; Feb hidden as given. */
const book = (hidden?: true | 'very'): BookInput => {
  const month = (name: string, k: number) => ({ name, cells: Object.fromEntries([2, 3, 4, 5].map((r) => [`B${r}`, { value: k * r }])) })
  return {
    sheets: [
      month('Jan', 1),
      { ...month('Feb', 10), rows: { 3: { hidden: true } }, ...(hidden ? { hidden } : {}) },
      month('Mar', 100),
      { name: 'Sum', cells: { A1: { formula: "SUM('Jan:Mar'!B2)" }, A2: { formula: 'SUBTOTAL(9,Feb!B2:B5)' }, A3: { formula: 'SUBTOTAL(109,Feb!B2:B5)' }, A4: { formula: 'Feb!B2*2' } } },
    ],
  }
}
const values = (b: BookInput) => Object.values(compute(b, { clock: CLOCK }).sheets[3].cells).map((c) => c.value)

describe('hidden sheets in the book', () => {
  it('compute the same values: a reference across sheets takes in a hidden one, SUBTOTAL skips hidden rows alone', () => {
    expect(values(book())).toEqual([222, 140, 110, 40])
    for (const h of [true, 'very'] as const) {
      expect(values(book(h))).toEqual([222, 140, 110, 40])
      expect(compute(book(h), { clock: CLOCK }).hash).toBe(compute(book(), { clock: CLOCK }).hash)
      expect(compute(book(h), { clock: CLOCK }).inputHash).not.toBe(compute(book(), { clock: CLOCK }).inputHash)
    }
  })

  it('a book keeps one visible sheet; hidden is true, false or "very"', () => {
    const all: BookInput = { sheets: [{ name: 'A', cells: {}, hidden: true }, { name: 'B', cells: {}, hidden: 'very' }] }
    expect(() => compute(all)).toThrow(new InputError('book.sheets', 'a book keeps one visible sheet at least: every sheet is hidden'))
    expect(() => compute({ sheets: [{ name: 'A', cells: {}, hidden: 'yes' as never }] })).toThrow(new InputError('book.sheets[0].hidden', 'expected true, false or "very", got the string "yes"'))
    expect(compute({ sheets: [{ name: 'A', cells: {}, hidden: false }, { name: 'B', cells: {}, hidden: true }] }).sheets.length).toBe(2)
  })
})

describe('hideSheet and unhideSheet', () => {
  const op = (b: BookInput, o: StructureOp) => {
    const r = restructure(b, o)
    return r.ok ? r.book.sheets.map((s) => `${s.name}${s.hidden === 'very' ? '!!' : s.hidden ? '!' : ''}`).join(' ') : r.code
  }

  it('hide, very hide and show a sheet; the last visible sheet is neither hidden nor deleted', () => {
    expect(op(book(), { op: 'hideSheet', sheet: 'feb' })).toBe('Jan Feb! Mar Sum')
    expect(op(book(), { op: 'hideSheet', sheet: 'Feb', very: true })).toBe('Jan Feb!! Mar Sum')
    expect(op(book('very'), { op: 'hideSheet', sheet: 'Feb' })).toBe('Jan Feb! Mar Sum')
    expect(op(book('very'), { op: 'unhideSheet', sheet: 'Feb' })).toBe('Jan Feb Mar Sum')
    expect(op(book(), { op: 'unhideSheet', sheet: 'Feb' })).toBe('Jan Feb Mar Sum')
    const one: BookInput = { sheets: [{ name: 'A', cells: {}, hidden: true }, { name: 'B', cells: {} }, { name: 'C', cells: {}, hidden: 'very' }] }
    expect(restructure(one, { op: 'hideSheet', sheet: 'B' })).toMatchObject({ ok: false, code: 'last_sheet' })
    expect(restructure(one, { op: 'deleteSheet', sheet: 'B' })).toMatchObject({ ok: false, code: 'last_sheet', message: 'a book keeps at least one visible sheet' })
    expect(op(one, { op: 'hideSheet', sheet: 'A', very: true })).toBe('A!! B C!!')
    expect(op(one, { op: 'deleteSheet', sheet: 'A' })).toBe('B C!!')
    expect(() => restructure(book(), { op: 'hideSheet', sheet: 'Feb', very: 1 } as never)).toThrow(new InputError('op.very', 'expected a boolean, got 1'))
  })

  it('a hidden sheet is renamed, moved and deleted as any other and keeps its state; a very hidden one is renamed alone', () => {
    expect(op(book(true), { op: 'renameSheet', sheet: 'Feb', name: 'February' })).toBe('Jan February! Mar Sum')
    expect(op(book(true), { op: 'moveSheet', sheet: 'Feb' })).toBe('Jan Mar Sum Feb!')
    expect(op(book(true), { op: 'insertSheet', name: 'New', before: 'Feb' })).toBe('Jan New Feb! Mar Sum')
    expect(op(book(true), { op: 'deleteSheet', sheet: 'Feb' })).toBe('Jan Mar Sum')
    expect(op(book('very'), { op: 'renameSheet', sheet: 'Feb', name: 'February' })).toBe('Jan February!! Mar Sum')
    expect(restructure(book('very'), { op: 'moveSheet', sheet: 'Feb' })).toMatchObject({ ok: false, code: 'very_hidden' })
    expect(restructure(book('very'), { op: 'deleteSheet', sheet: 'Feb' })).toMatchObject({ ok: false, code: 'very_hidden' })
    const r = restructure(book(true), { op: 'renameSheet', sheet: 'Feb', name: 'February' })
    expect(r.ok && r.book.sheets[3].cells.A4.formula).toBe("'February'!B2*2")
  })

  it('a sheet moved before a hidden one goes before the next visible sheet, or to the end', () => {
    expect(op(book(true), { op: 'moveSheet', sheet: 'Sum', before: 'Feb' })).toBe('Jan Feb! Sum Mar')
    const two: BookInput = { sheets: ['A', 'B', 'C', 'D', 'E'].map((name) => ({ name, cells: {}, ...('BC'.includes(name) ? { hidden: true } : {}) })) }
    expect(op(two, { op: 'moveSheet', sheet: 'E', before: 'B' })).toBe('A B! C! E D')
    expect(op(two, { op: 'moveSheet', sheet: 'D', before: 'B' })).toBe('A B! C! D E')
    expect(op({ sheets: [...two.sheets.slice(0, 3), { name: 'D', cells: {}, hidden: true }] }, { op: 'moveSheet', sheet: 'A', before: 'B' })).toBe('B! C! D! A')
    // a 3D reference covers the sheets between its ends where the move put them
    const r = restructure(book(true), { op: 'moveSheet', sheet: 'Mar', before: 'Feb' })
    expect(r.ok && r.book.sheets.map((s) => s.name)).toEqual(['Jan', 'Feb', 'Mar', 'Sum'])
  })
})

describe.each(['closures', 'eval'] as const)('a session hides and shows sheets (%s)', (execution) => {
  it('recomputes nothing, keeps the hash and refuses the last visible sheet', () => {
    const s = openSession(book(), { clock: CLOCK, execution })
    const hash = s.hash()
    const r = s.restructure({ op: 'hideSheet', sheet: 'Feb', very: true })
    expect(r.ok && r.changes).toEqual([])
    expect(r.ok && r.mapCell({ sheet: 'Feb', cell: 'B2' })).toEqual({ sheet: 'Feb', cell: 'B2' })
    expect(s.book().sheets[1].hidden).toBe('very')
    expect(s.hash()).toBe(hash)
    expect(s.result().inputHash).toBe(compute(s.book(), { clock: CLOCK }).inputHash)
    // later edits reach the hidden sheet and the formulas over it
    expect(s.setCells([{ sheet: 'Feb', cell: 'B2', input: { value: 1 } }]).map((c) => `${c.sheet}!${c.cell} ${c.value}`).sort()).toEqual(['Feb!B2 1', 'Sum!A1 203', 'Sum!A2 121', 'Sum!A3 91', 'Sum!A4 2'])
    expect(s.restructure({ op: 'unhideSheet', sheet: 'Feb' }).ok).toBe(true)
    expect('hidden' in s.book().sheets[1]).toBe(false)
    for (const sheet of ['Jan', 'Mar', 'Sum']) expect(s.restructure({ op: 'hideSheet', sheet }).ok).toBe(true)
    expect(s.restructure({ op: 'hideSheet', sheet: 'Feb' })).toMatchObject({ ok: false, code: 'last_sheet' })
    expect(s.restructure({ op: 'deleteSheet', sheet: 'Feb' })).toMatchObject({ ok: false, code: 'last_sheet' })
    expect(s.result().hash).toBe(compute(s.book(), { clock: CLOCK }).hash)
    s.close()
  })
})

describe('in the same thread and in the pool', () => {
  const open: Calc[] = []
  afterAll(() => Promise.all(open.map((p) => p.close())))
  const steps: StructureOp[] = [
    { op: 'hideSheet', sheet: 'Jan' },
    { op: 'hideSheet', sheet: 'Mar', very: true },
    { op: 'renameSheet', sheet: 'Mar', name: 'March' },
    { op: 'hideSheet', sheet: 'Sum' },
    { op: 'hideSheet', sheet: 'Feb' },
    { op: 'deleteSheet', sheet: 'Feb' },
    { op: 'unhideSheet', sheet: 'Jan' },
    { op: 'deleteSheet', sheet: 'Feb' },
  ]

  it.each([['inline', () => createInlineCalc()], ['pool', () => pool({ threads: 1 })]] as const)('%s: the synchronous session\'s answers and book', async (_, make) => {
    const calc = await make()
    open.push(calc)
    const a = await calc.openSession(book(), { clock: CLOCK })
    const y = openSession(book(), { clock: CLOCK })
    for (const o of steps) {
      const [r, w] = [await a.restructure(o), y.restructure(o)]
      expect(r.ok ? { ok: true, changes: r.changes } : r, JSON.stringify(o)).toEqual(w.ok ? { ok: true, changes: w.changes } : w)
      if (r.ok) expect(r.mapCell({ sheet: 'Sum', cell: 'A1' })).toEqual({ sheet: 'Sum', cell: 'A1' })
    }
    expect(await a.book()).toEqual(y.book())
    expect(y.book().sheets.map((s) => [s.name, s.hidden])).toEqual([['Jan', undefined], ['March', 'very'], ['Sum', true]])
    expect(await a.hash()).toBe(y.hash())
  }, 60_000)
})

describe('xlsx', () => {
  it('writes the state of hidden sheets and makes the first visible sheet active; reads them back', async () => {
    const b: BookInput = { sheets: [{ name: 'A', cells: {}, hidden: true }, { name: 'B', cells: {}, hidden: 'very' }, { name: 'C', cells: {} }, { name: 'D', cells: {}, hidden: true }] }
    const { data, issues } = await writeXlsx(b)
    expect(issues).toEqual([])
    const part = (name: string) => openArchive(data).text(name)
    expect(await part('xl/workbook.xml')).toContain('<bookViews><workbookView activeTab="2"/></bookViews><sheets><sheet name="A" sheetId="1" state="hidden" r:id="rId1"/><sheet name="B" sheetId="2" state="veryHidden" r:id="rId2"/><sheet name="C" sheetId="3" r:id="rId3"/><sheet name="D" sheetId="4" state="hidden" r:id="rId4"/></sheets>')
    expect(await Promise.all([1, 2, 3, 4].map(async (k) => (await part(`xl/worksheets/sheet${k}.xml`))!.includes('tabSelected="1"')))).toEqual([false, false, true, false])
    const back = await readXlsx(data)
    expect(back.book).toEqual(b)
    expect(back.issues).toEqual([])
    // a book whose first sheet is visible keeps Excel's default active tab
    expect(await openArchive((await writeXlsx(book(true))).data).text('xl/workbook.xml')).toContain('<workbookView/>')
  })
})
