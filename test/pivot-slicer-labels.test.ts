/**
 * A slicer's items labelled by `getLayout` in the workbook's locale, worked out from the pivot cache on each call: the
 * same book in en-US and ru-RU, a label for an item a refresh brings, in a session of each execution, the inline
 * calculator and a pool's threads; and the labels Excel for Windows showed on the slicers of the set slicer-labels.
 */
import { describe, expect, it } from 'vitest'
import { createInlineCalc, createSession, editStructure, registerLocale, registerPivot, type BookInput, type StructureOp } from '../src/index.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { pivotEngine } from '../src/pivot/index.js'
import { closedAfter } from './fixtures/pool.js'

const day = (y: number, m: number, d: number) => Date.UTC(y, m - 1, d) / 86_400_000 + 25_569
/** Region with a blank, Price in a currency, Day dates, Flag booleans, Amount to sum. */
const ROWS: [string | null, number, number, boolean, number][] = [
  ['North', 1234.5, day(2024, 1, 15), true, 1],
  [null, 0.5, day(2025, 3, 5), false, 2],
  ['East', -20, day(2024, 1, 15), true, 3],
]
const cells: BookInput['sheets'][number]['cells'] = Object.fromEntries(['Region', 'Price', 'Day', 'Flag', 'Amount'].map((v, c) => [`${'ABCDE'[c]}1`, { value: v }]))
ROWS.forEach(([region, price, d, flag, amount], i) => {
  if (region) cells[`A${i + 2}`] = { value: region }
  cells[`B${i + 2}`] = { value: price, numberFormat: '$#,##0.00' }
  cells[`C${i + 2}`] = { value: d, numberFormat: 14 }
  cells[`D${i + 2}`] = { value: flag }
  cells[`E${i + 2}`] = { value: amount }
})
const BOOK: BookInput = {
  sheets: [
    { name: 'D', cells },
    { name: 'P', cells: {}, pivotTables: [{ name: 'PT1', cache: 1, location: { range: 'A3:B4', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 }, fields: [{}, {}, {}, {}, {}], rows: [0], values: [{ field: 4 }] }] },
  ],
  pivotCaches: [{ id: 1, source: { type: 'range', sheet: 'D', range: 'A1:E4' }, fields: ['Region', 'Price', 'Day', 'Flag', 'Amount'].map((name) => ({ name })) }],
}
const PT1 = { sheet: 'P', name: 'PT1' }
const slicer = (field: string): StructureOp => ({ op: 'addSlicer', sheet: 'P', slicer: {}, cache: { field, pivots: [PT1] } })
const FIELDS = ['Region', 'Price', 'Day', 'Flag']

/** The book refreshed with a slicer over each field. */
function sliced(locale?: string): BookInput {
  let r = editStructure(locale ? { ...BOOK, locale } : BOOK, { op: 'refreshPivots' })
  for (const f of FIELDS) if (r.ok) r = editStructure(r.book, slicer(f))
  if (!r.ok) throw new Error(r.message)
  return r.book
}

type Layout = { slicerCaches?: { field: string; items?: { label: string }[] }[] }
const labels = (layout: Layout) => Object.fromEntries((layout.slicerCaches ?? []).map((c) => [c.field, c.items!.map((x) => x.label)]))

const EN = { Region: ['East', 'North', '(blank)'], Price: ['-$20.00', '$0.50', '$1,234.50'], Day: ['1/15/2024', '3/5/2025'], Flag: ['FALSE', 'TRUE'] }
const RU = { Region: ['East', 'North', '(пусто)'], Price: ['-$20,00', '$0,50', '$1 234,50'], Day: ['15.01.2024', '05.03.2025'], Flag: ['ЛОЖЬ', 'ИСТИНА'] }
/** After Region gets South in the source and a refresh: the new item comes last, unchosen. */
const REFRESHED = { ...EN, Region: ['East', 'North', 'South', '(blank)'] }

describe("a slicer's items labelled in getLayout", () => {
  it('loads', () => {
    registerPivot(pivotEngine)
    registerLocale(ruRU)
  })

  for (const execution of ['closures', 'eval'] as const)
    it(`in the workbook's locale, again after a refresh brings an item, in a session (${execution})`, () => {
      expect(labels(createSession(sliced('ru-RU'), { execution }).getLayout())).toEqual(RU)
      const s = createSession(sliced(), { execution })
      expect(labels(s.getLayout())).toEqual(EN)
      // the book's own caches hold no label
      expect(s.getBook().slicerCaches!.every((c) => c.items!.every((x) => !('label' in x)))).toBe(true)
      s.setCells([{ sheet: 'D', cell: 'A3', input: { value: 'South' } }])
      expect(s.editStructure({ op: 'refreshPivots' }).ok).toBe(true)
      expect(labels(s.getLayout())).toEqual(REFRESHED)
      s.undo()
      expect(labels(s.getLayout())).toEqual(EN)
      s.close()
    })

  const { made } = closedAfter()
  for (const [kind, calc] of [['inline', async () => createInlineCalc()], ['pool', () => made({ threads: 1 })]] as const)
    it(`in the ${kind} calculator, in each locale and after a refresh`, async () => {
      const c = await calc()
      const ru = await c.createSession(sliced('ru-RU'))
      expect(labels(await ru.getLayout())).toEqual(RU)
      await ru.close()
      const s = await c.createSession(sliced())
      await s.setCells([{ sheet: 'D', cell: 'A3', input: { value: 'South' } }])
      expect((await s.editStructure({ op: 'refreshPivots' })).ok).toBe(true)
      expect(labels(await s.getLayout())).toEqual(REFRESHED)
      await s.close()
    })
})
