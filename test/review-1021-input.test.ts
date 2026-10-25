/**
 * Input the contract types but did not check: the values of external functions, a cell's meta, a value that is not a
 * finite number and a clock that is not a moment the same in every time zone, refused at the boundary as InputError.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createInlineCalc, createSession, displayValue, editStructure, entryText, InputError, parseEntry, type BookInput, type CellInput } from '../src/index.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { mergedArea, rangeAddress, rectContains, rectsMeet } from '../src/a1.js'
import { ASYNC_CALCS, closeCalcs, KINDS, openerFor } from './session-kinds.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const BALANCE = 'BALANCE("x")'
/** The path and message of the InputError `f` throws. */
function refusal(f: () => unknown): { path: string; message: string } | string {
  try {
    f()
    return 'no refusal'
  } catch (e) {
    return e instanceof InputError ? { path: e.path, message: e.message } : String(e)
  }
}

const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: BALANCE }, A2: { formula: 'A1+1' } } }] }

describe.each(['closures', 'eval'] as const)('%s', (execution) => {
  const withValue = (v: unknown) => ({ clock, execution, externals: { functions: ['BALANCE'], values: { [BALANCE]: v } } }) as never

  it.each([
    ['an unknown error code', { error: 'BOGUS' }],
    ['a prototype key as an error code', { error: 'toString' }],
    ['a date input', { date: '2026-01-01' }],
    ['an array', [1, 2]],
    ['a Date', new Date(0)],
    ['undefined', undefined],
    ['a function', () => 1],
    ['a BigInt', 5n],
  ])('a value of an external function: %s is refused', (_, v) => {
    const path = `opts.externals.values[${JSON.stringify(BALANCE)}]`
    expect(refusal(() => compute(book, withValue(v)))).toEqual({ path, message: expect.stringMatching(/expected a number, text, a boolean, null or \{ error: code \}/) })
    expect(() => createSession(book, withValue(v))).toThrow(InputError)
  })

  it('a value of an external function: a CellValue computes, a non-finite number still gives #NUM!', () => {
    const a2 = (v: unknown) => compute(book, withValue(v)).sheets[0].cells.A2.value
    expect([a2(3), a2('7'), a2(true), a2(null), a2({ error: 'DIV0' })]).toEqual([4, 8, 2, 1, { error: 'DIV0' }])
    expect([a2(NaN), a2(Infinity)]).toEqual([{ error: 'NUM' }, { error: 'NUM' }])
  })

  it("a cell's meta that is cyclic or holds a BigInt is refused, in compute, a session's edits and on the same thread", async () => {
    const cyclic: Record<string, unknown> = { row: 1 }
    cyclic.self = { back: cyclic }
    for (const [meta, why] of [[cyclic, /cycle/], [{ id: 10n }, /BigInt/], [[1, { n: 2n }], /BigInt/]] as const) {
      const b: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1, meta: meta as never } } }] }
      expect(refusal(() => compute(b, { clock, execution }))).toEqual({ path: expect.stringMatching(/^book\.sheets\[0\]\.cells\["A1"\]\.meta/), message: expect.stringMatching(why) })
      expect(() => createSession(b, { clock, execution })).toThrow(InputError)
      const s = createSession(book, { clock, execution })
      expect(() => s.setCells([{ sheet: 'S', cell: 'B1', input: { value: 1, meta: meta as never } }])).toThrow(InputError)
      s.close()
      const calc = createInlineCalc()
      await expect(calc.compute(b, { clock, execution })).rejects.toThrow(InputError)
      calc.close()
    }
    // JSON meta, shared but not cyclic, stays as it was
    const shared = { tag: 'x' }
    const r = compute({ sheets: [{ name: 'S', cells: { A1: { value: 1, meta: { a: shared, b: shared } }, A2: { value: 2, meta: 'm' } } }] }, { clock, execution })
    expect(r.sheets[0].cells.A1.meta).toEqual({ a: shared, b: shared })
  })
})

describe.each(['closures', 'eval'] as const)('a value that is not a finite number is refused (fix-u373), %s', (execution) => {
  afterAll(closeCalcs)
  const cell = (v: number): CellInput => ({ value: v })
  const at = (v: number): BookInput => ({ sheets: [{ name: 'S', cells: { A1: cell(v), A2: { formula: 'A1+1' } } }] })
  const BAD = [NaN, Infinity, -Infinity]
  const why = { path: 'book.sheets[0].cells["A1"].value', message: expect.stringMatching(/expected a finite number/) }

  it('compute and createSession, for each value', () => {
    for (const v of BAD) {
      expect(refusal(() => compute(at(v), { clock, execution })), String(v)).toEqual(why)
      expect(refusal(() => createSession(at(v), { clock, execution })), String(v)).toEqual(why)
    }
    expect(compute(at(-0), { clock, execution }).sheets[0].cells.A2.value).toBe(1)
  })

  it.each(ASYNC_CALCS)('compute and createSession %s', async (_, make) => {
    const calc = await make()
    for (const v of BAD) {
      await expect(calc.compute(at(v), { clock, execution })).rejects.toMatchObject(why)
      await expect(calc.createSession(at(v), { clock, execution })).rejects.toMatchObject(why)
    }
    await calc.close()
  })

  describe.each(KINDS)('the edits of %s', (kind) => {
    const open = openerFor(kind)
    it('setCells, replaceCells and applyChanges refuse it and change nothing', async () => {
      const s = await open(at(5), { clock, execution })
      const before = (await s.getResult()).hash
      for (const v of BAD) {
        const edit = [{ sheet: 'S', cell: 'A1', input: cell(v) }]
        await expect(s.setCells(edit), String(v)).rejects.toMatchObject({ path: 'edits[0].input.value' })
        await expect(s.replaceCells(edit), String(v)).rejects.toMatchObject({ path: 'edits[0].input.value' })
        await expect(s.applyChanges([{ op: 'setCells', edits: edit }]), String(v)).rejects.toMatchObject({ path: 'operations[0].edits[0].input.value' })
        await expect(s.applyChanges([{ op: 'replaceCells', edits: edit }]), String(v)).rejects.toMatchObject({ path: 'operations[0].edits[0].input.value' })
        await expect(s.validateEntry('S', 'A1', cell(v)), String(v)).rejects.toMatchObject({ path: 'input.value' })
      }
      expect((await s.getResult()).hash).toBe(before)
      await s.close()
    })
  })
})

describe.each(['closures', 'eval'] as const)("a book editStructure returned is checked whole, the caller's cells it shares included (speed-u95), %s", (execution) => {
  afterAll(closeCalcs)
  const given = (): BookInput => ({ styles: [{}, { font: { bold: true } }], sheets: [{ name: 'S', cells: { A1: { value: 1, style: 1 }, A2: { formula: 'A1+1' } } }] })
  // the caller changes a cell of its own book after the call: the returned book holds the same cell
  const changed = (cell: CellInput): BookInput => {
    const book = given()
    const r = editStructure(book, { op: 'insertRows', sheet: 'S', row: 3 })
    if (!r.ok) throw new Error(r.message)
    Object.assign(book.sheets[0].cells.A1, cell)
    return r.book
  }
  const BAD: [string, string, CellInput][] = [
    ['a value of NaN', 'value', { value: Number.NaN }],
    ['a style the book does not have', 'style', { style: 7 }],
    ['a value that is an object', 'value', { value: { foo: 1 } as never }],
  ]

  it.each(BAD)('%s: editStructure, compute and createSession refuse it', (_, field, cell) => {
    const why = { path: `book.sheets[0].cells["A1"].${field}` }
    expect(refusal(() => editStructure(changed(cell), { op: 'insertRows', sheet: 'S', row: 1 }))).toMatchObject(why)
    expect(refusal(() => compute(changed(cell), { clock, execution }))).toMatchObject(why)
    expect(refusal(() => createSession(changed(cell), { clock, execution }))).toMatchObject(why)
  })

  it.each(ASYNC_CALCS)('compute and createSession %s refuse it', async (_, make) => {
    const calc = await make()
    for (const [, field, cell] of BAD) {
      const why = { path: `book.sheets[0].cells["A1"].${field}` }
      await expect(calc.compute(changed(cell), { clock, execution })).rejects.toMatchObject(why)
      await expect(calc.createSession(changed(cell), { clock, execution })).rejects.toMatchObject(why)
    }
    await calc.close()
  })
})

describe.each(['closures', 'eval'] as const)('a clock is ISO 8601 with Z or an offset, or a date alone (fix-u374), %s', (execution) => {
  afterAll(closeCalcs)
  const now = { sheets: [{ name: 'S', cells: { A1: { formula: 'NOW()' } } }] }
  const at = (moment: string) => compute(now, { clock: { now: moment, seed: 7 }, execution }).sheets[0].cells.A1.value
  // read in the host's time zone, by the engine's own rules, or a day the calendar does not have
  const REFUSED = ['2026-10-04T12:00:00', '2026-10-04T12:00', '10/4/2026', '1', 'Sun Oct 04 2026', '2026-02-30T00:00:00Z', '2026-02-30', '2026-13-01', '2026-10-04T24:00:00Z', '2026-10-04 12:00:00Z', '2026-10']
  const why = (path: string) => ({ path, message: expect.stringMatching(/invalid moment .*expected ISO 8601 with Z or an offset/) })

  it('compute, createSession, setClock and the options of typed text refuse the others', () => {
    const s = createSession(now, { clock, execution })
    for (const moment of REFUSED) {
      expect(refusal(() => compute(now, { clock: { now: moment, seed: 7 }, execution })), moment).toEqual(why('opts.clock.now'))
      expect(refusal(() => createSession(now, { clock: { now: moment, seed: 7 }, execution })), moment).toEqual(why('opts.clock.now'))
      expect(refusal(() => s.setClock({ now: moment, seed: 7 })), moment).toEqual(why('clock.now'))
      expect(refusal(() => parseEntry('3/15', { now: moment })), moment).toEqual(why('opts.now'))
      expect(refusal(() => entryText(1, { now: moment })), moment).toEqual(why('opts.now'))
    }
    expect(s.getResult().clock).toEqual(clock)
    s.close()
  })

  it.each(ASYNC_CALCS)('compute and createSession %s refuse them', async (_, make) => {
    const calc = await make()
    await expect(calc.compute(now, { clock: { now: '2026-10-04T12:00:00', seed: 7 }, execution })).rejects.toMatchObject(why('opts.clock.now'))
    await expect(calc.createSession(now, { clock: { now: '2026-02-30T00:00:00Z', seed: 7 }, execution })).rejects.toMatchObject(why('opts.clock.now'))
    await calc.close()
  })

  it('Z, an offset, fractions of a second and a date alone name one moment wherever the host is', () => {
    expect(at('2026-10-04T12:00:00Z')).toBe(46299.5)
    expect([at('2026-10-04T15:00:00+03:00'), at('2026-10-04T05:00-07:00'), at('2026-10-04T12:00:00.000Z'), at('2026-10-04T12:00:00.123456Z')]).toEqual([46299.5, 46299.5, 46299.5, at('2026-10-04T12:00:00.123Z')])
    expect([at('2026-10-04'), at('2024-02-29'), at('2026-10-05T02:00:00+14:00')]).toEqual([46299, 45351, 46299.5])
    expect(parseEntry('3/15', { now: '2026-12-31T23:30:00-08:00' })).toEqual(parseEntry('3/15', { now: '2027-01-01T07:30:00Z' }))
  })
})

describe("a locale object not loaded: its rules checked where it is taken", () => {
  const base = { ...ruRU, tag: 'xx-XX' }
  it('refuses rules with a field missing or not an object as an InputError naming the field', () => {
    expect(refusal(() => displayValue(1234.5, { numberFormat: '#,##0.00', locale: { ...base, regional: { decimal: ',' } } as never }))).toEqual({ path: 'opts.locale.regional.group', message: 'opts.locale.regional.group: expected a value' })
    expect(refusal(() => parseEntry('1,5', { locale: { ...base, regional: { ...ruRU.regional, extra: 1 } } as never }))).toMatchObject({ path: 'opts.locale.regional' })
    expect(refusal(() => entryText(1.5, { locale: { ...base, regional: 'x' } as never }))).toMatchObject({ path: 'opts.locale.regional' })
  })

  it('takes a locale without rules, calculated by en-US\'s, and one with all of them', () => {
    expect(displayValue(1234.5, { numberFormat: '#,##0.00', locale: { ...base, regional: undefined } })).toEqual({ text: '1,234.50' })
    expect(displayValue(1234.5, { numberFormat: '#,##0.00', locale: base }).text).toBe(displayValue(1234.5, { numberFormat: '#,##0.00', locale: ruRU }).text)
  })
})

describe('the rectangles of xlcalc/a1', () => {
  const reversed = { top: 5, left: 3, bottom: 1, right: 1 }
  const ordered = { top: 1, left: 1, bottom: 5, right: 3 }
  const merge = { top: 3, left: 2, bottom: 4, right: 4 }

  it('are taken with their corners in either order, as rangeAddress takes them', () => {
    expect(rangeAddress(reversed)).toBe('A1:C5')
    expect(mergedArea(reversed, [merge])).toEqual({ top: 1, left: 1, bottom: 5, right: 4 })
    expect(mergedArea(ordered, [{ top: 4, left: 4, bottom: 3, right: 2 }])).toEqual({ top: 1, left: 1, bottom: 5, right: 4 })
    expect(mergedArea({ top: 3, left: 3, bottom: 3, right: 3 }, [{ top: 4, left: 4, bottom: 2, right: 2 }])).toEqual({ top: 2, left: 2, bottom: 4, right: 4 })
    expect([rectsMeet(reversed, merge), rectContains(reversed, { top: 2, left: 2, bottom: 2, right: 2 })]).toEqual([true, true])
  })

  it('are refused off the sheet or not a rectangle, as an InputError naming the argument', () => {
    const zeros = { top: 0, left: 0, bottom: 0, right: 0 }
    expect(refusal(() => rectsMeet(zeros, zeros))).toMatchObject({ path: 'a.top' })
    expect(refusal(() => rectContains(ordered, { top: '1', left: 1, bottom: 1, right: 1 } as never))).toMatchObject({ path: 'inner.top' })
    expect(refusal(() => rectsMeet(undefined as never, ordered))).toMatchObject({ path: 'a.top' })
    expect(refusal(() => rangeAddress(undefined as never))).toMatchObject({ path: 'rect.top' })
    expect(refusal(() => mergedArea(ordered, [null as never]))).toMatchObject({ path: 'merges[0].top' })
  })
})
