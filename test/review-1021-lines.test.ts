/**
 * A sheet's rows and columns, from the review of 21.10 (speed): a map of many entries is read once per call, not once
 * per cell, edit or line. Each case took seconds before, on a size where it now takes a few milliseconds or tens of them.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { editStructure, type BookInput, type Execution, type RowInput } from '../src/index.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const EXECUTIONS: Execution[] = ['closures', 'eval']

afterAll(closeCalcs)

/** Milliseconds a call takes. */
async function took(f: () => unknown): Promise<number> {
  const t0 = performance.now()
  await f()
  return performance.now() - t0
}

/** Milliseconds a synchronous call takes. */
function tookSync(f: () => unknown): number {
  const t0 = performance.now()
  f()
  return performance.now() - t0
}

/** Rows 1 to `n`, each its own entry: heights that differ from the next row's, as an xlsx with fitted rows keeps them. */
const manyRows = (n: number, entry: (r: number) => RowInput = () => ({})): Record<string, RowInput> =>
  Object.fromEntries(Array.from({ length: n }, (_, i) => [String(i + 1), { height: 15 + (i % 3), ...entry(i + 1) }]))

describe.each(EXECUTIONS)('review 21.10 rows and columns, %s', (execution) => {
  it('a series fill and a merge across read a sheet of 34,000 row entries once (speed-u94)', () => {
    const cells = Object.fromEntries(Array.from({ length: 300 }, (_, i) => [`A${i + 1}`, { value: i + 1 }]))
    const book: BookInput = { sheets: [{ name: 'S', rows: manyRows(34_000), cells }] }
    // the rows read once per cell of the series' source, once per row merged: 2 to 3 s each
    let fill: ReturnType<typeof editStructure> | undefined, merge: ReturnType<typeof editStructure> | undefined
    expect(tookSync(() => (fill = editStructure(book, { op: 'fillRange', sheet: 'S', range: 'A1:A300', to: 'A1:A310' })))).toBeLessThan(500)
    expect(fill!.ok && fill!.book.sheets[0].cells.A310).toEqual({ value: 310 })
    expect(tookSync(() => (merge = editStructure(book, { op: 'mergeCells', sheet: 'S', range: 'C1:D600', across: true })))).toBeLessThan(500)
    expect(merge!.ok && merge!.book.sheets[0].merges!.length).toBe(600)
  })

  describe.each(KINDS)('in %s', (kind) => {
    const open = openerFor(kind)

    it('a list of data validation reads the rows once, not once per item (speed-u94)', async () => {
      const rows = manyRows(34_000, (r) => (r === 7 ? { numberFormat: '0.0%' } : {}))
      const cells = Object.fromEntries(Array.from({ length: 1000 }, (_, i) => [`B${i + 1}`, { value: i }]))
      const s = await open({ sheets: [{ name: 'S', rows, cells, validations: [{ range: 'D1', type: 'list', formula1: '$B$1:$B$1000' }] }] }, { execution, clock })
      let list: unknown
      // the rows read for each item: 5 s
      expect(await took(async () => (list = (await s.getValidation('S', 'D1'))!.list))).toBeLessThan(500)
      expect((list as { text: string }[]).map((x) => x.text).slice(5, 8)).toEqual(['5', '600.0%', '7'])
      expect((list as unknown[]).length).toBe(1000)
      await s.close()
    })

    it('typing text with a line break into 20,000 rows patches the rows map once (speed-u116)', async () => {
      const s = await open({ sheets: [{ name: 'S', cells: {} }] }, { execution, clock, display: false, undo: { depth: 0 } })
      const edits = Array.from({ length: 20_000 }, (_, i) => ({ sheet: 'S', cell: `A${i + 1}`, text: 'a\nb' }))
      // one patch of the whole map per row: 22 s
      expect(await took(() => s.setCellTexts(edits))).toBeLessThan(4000)
      const rows = (await s.getLayout()).sheets[0].rows!
      expect(Object.keys(rows).length).toBe(20_000)
      expect(rows['20000']).toEqual({ height: 32, bestFit: true })
      await s.close()
    })

    it('typing into a wrapped cell and AutoFit of a row do not walk every cell of the sheet (speed-u96)', async () => {
      const cells = Object.fromEntries(Array.from({ length: 200_000 }, (_, i) => [`${'ABCDEFGHIJ'[i % 10]}${Math.floor(i / 10) + 1}`, { value: i }]))
      const s = await open({ styles: [{}, { alignment: { wrap: true } }], sheets: [{ name: 'S', cells: { ...cells, L5: { value: null, style: 1 } } }] }, { execution, clock, display: false, undo: { depth: 0 } })
      // each call walked the 200,000 cells with a regular expression: about 40 ms each, 8 s for 200
      expect(
        await took(async () => {
          for (let k = 0; k < 100; k++) {
            await s.setCellText('S', 'L5', k % 2 ? 'x\ny' : 'x\ny\nz')
            await s.setLayout([{ sheet: 'S', rows: '5', fit: true }])
          }
        }),
      ).toBeLessThan(1500)
      expect((await s.getLayout()).sheets[0].rows).toEqual({ 5: { height: 32, bestFit: true } })
      await s.close()
    })

    it('applyChanges and setLayout on a sheet of 34,000 row entries compare the rows by identity (speed-u111)', async () => {
      const s = await open({ sheets: [{ name: 'S', rows: manyRows(34_000, (r) => (r % 5 ? {} : { hidden: true })), cells: { A1: { value: 1 }, B1: { formula: 'SUBTOTAL(109,A1:A10)' } } }] }, { execution, clock, undo: { depth: 0 } })
      // each call wrote out and read the rows map: about 13 ms, 6 s for 500
      expect(
        await took(async () => {
          for (let k = 0; k < 250; k++) {
            await s.applyChanges([{ op: 'setCells', edits: [{ sheet: 'S', cell: 'A2', input: { value: k } }] }])
            await s.setLayout([{ sheet: 'S', columns: 'C', set: { width: 10 + (k % 5) } }])
          }
        }),
      ).toBeLessThan(1500)
      expect((await s.getCell('S', 'B1'))!.value).toBe(250)
      await s.close()
    })
  })

  it.each([
    ['8,000 rows of their own entries in one group', manyRows(8000, () => ({ outlineLevel: 1 }))],
    ['3,000 groups of three rows', Object.fromEntries(Array.from({ length: 3000 }, (_, i) => [`${4 * i + 1}:${4 * i + 3}`, { outlineLevel: 1 }]))],
  ])('outline commands over %s patch the rows once (speed-u92)', (_, rows) => {
    const book: BookInput = { sheets: [{ name: 'S', rows, cells: {} }] }
    let hidden: BookInput | undefined
    // a patch of the whole map per entry and per mark: 7 to 12 s for each
    expect(
      tookSync(() => {
        const r = editStructure(book, { op: 'showLevels', sheet: 'S', rows: 1 })
        if (!r.ok) throw new Error(r.code)
        hidden = r.book
        const shown = editStructure(hidden, { op: 'showLevels', sheet: 'S', rows: 2 })
        if (!shown.ok) throw new Error(shown.code)
      }),
    ).toBeLessThan(1000)
    expect(Object.values(hidden!.sheets[0].rows!).every((e) => e.outlineLevel !== 1 || e.hidden === true)).toBe(true)
  })
})
