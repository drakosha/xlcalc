/**
 * A slicer's items labelled by `getLayout` in the workbook's locale, worked out from the pivot cache on each call: the
 * same book in en-US and ru-RU, a label for an item a refresh brings, in a session of each execution, the inline
 * calculator and a pool's threads; and the labels Excel for Windows showed on the slicers of the set slicer-labels.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { buildPivotCases } from '../bench/coverage/pivot/slicer-labels.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
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

describe('the labels of the slicers Excel for Windows made (the set slicer-labels)', () => {
  for (const tag of ['', 'ru-RU'])
    it(`are the captions Excel showed (${tag || 'en-US'}): in its own number format, a PivotTable's, the short date and time, a number written out`, async () => {
      const dir = new URL(`./excel/win/pivot/${tag ? `${tag}/` : ''}`, import.meta.url)
      const answers = (JSON.parse(readFileSync(new URL('slicer-labels.json', dir), 'utf8')) as { answers: Record<string, { phases: { results?: Record<string, unknown> }[] }> }).answers
      let n = 0
      for (const c of buildPivotCases()) {
        const results = answers[c.id].phases.at(-1)!.results ?? {}
        const layout = createSession({ ...(await readXlsx(readFileSync(new URL(`slicer-labels/${c.id}.xlsx`, dir)))).book, locale: tag || 'en-US' }).getLayout()
        c.ops.forEach((o, k) => {
          // a slicer Excel's object model did not make (the grouped fields' names of ru-RU) has no answer
          if (!('slicerGet' in o) || !results[`op ${k}`]) return
          const shown = (results[`op ${k}`] as [unknown, string[][]])[1].map((x) => x[1])
          const own = layout.sheets.flatMap((s) => s.slicers ?? []).find((s) => s.name === o.slicerGet)!
          expect([c.id, layout.slicerCaches!.find((x) => x.name === own.cache)!.items!.map((x) => x.label)]).toEqual([c.id, shown])
          n++
        })
      }
      expect(n).toBeGreaterThan(19)
    })
})

describe('a field of times of day without a date', () => {
  it("is read from Excel's 1899-12-30, laid out, labelled and written back as Excel writes it", async () => {
    const r = await readXlsx(readFileSync(new URL('./excel/win/pivot/slicer-labels/time.xlsx', import.meta.url)))
    expect(r.book.pivotCaches![0].fields.find((f) => f.name === 'Time')!.items![0]).toEqual({ date: '1900-01-00T06:00:00' })
    const s = createSession(r.book)
    expect(s.editStructure({ op: 'refreshPivots' }).ok).toBe(true)
    expect(s.getValues('P', 'A4:A5')).toEqual([[0.0625], [0.25]])
    expect(s.getLayout().slicerCaches![0].items!.map((x) => x.label)).toEqual(['1:30', '6:00', '12:00', '13:30', '18:00', '23:58'])
    const xml = (await openArchive((await writeXlsx({ book: s.getBook() })).data).text('xl/pivotCache/pivotCacheDefinition1.xml'))!
    expect(/<cacheField name="Time".*?<\/cacheField>/.exec(xml)![0]).toContain('minDate="1899-12-30T01:30:00" maxDate="1899-12-30T23:58:34" count="6"><d v="1899-12-30T06:00:00"/>')
    s.close()
  })
})

describe("a slicer's items after its PivotTables are refreshed, over Excel's books (the set slicer-labels)", () => {
  const dir = new URL('./excel/win/pivot/', import.meta.url)
  const answers = (JSON.parse(readFileSync(new URL('slicer-labels.json', dir), 'utf8')) as { answers: Record<string, { phases: { slicers?: Record<string, [string, boolean, boolean][]> }[] }> }).answers
  /** Each slicer's items as Excel's slicer lists them: label, chosen, with data. */
  type Items = Record<string, [string, boolean, boolean][]>
  const shown = (layout: Awaited<ReturnType<ReturnType<typeof createSession>['getLayout']>>): Items =>
    Object.fromEntries(layout.sheets.flatMap((s) => s.slicers ?? []).map((sl) => [sl.name, layout.slicerCaches!.find((c) => c.name === sl.cache)!.items!.map((x) => [x.label, x.selected !== false, !x.noData] as [string, boolean, boolean])]))
  // numbers grouped: the groups, those with data first; errors as the PivotTable lists them; months and years grouped
  const IDS = ['n-numbers', 'err', 'day-months-years']
  const book = async (id: string) => (await readXlsx(readFileSync(new URL(`slicer-labels/${id}.xlsx`, dir)))).book

  for (const execution of ['closures', 'eval'] as const)
    it(`are the groups and the order Excel showed, in a session (${execution})`, async () => {
      for (const id of IDS) {
        const s = createSession(await book(id), { execution })
        expect(s.editStructure({ op: 'refreshPivots' }).ok).toBe(true)
        expect([id, shown(s.getLayout())]).toEqual([id, answers[id].phases.at(-1)!.slicers])
        s.close()
      }
    })

  const { made } = closedAfter()
  for (const [kind, calc] of [['inline', async () => createInlineCalc()], ['pool', () => made({ threads: 1 })]] as const)
    it(`are the groups and the order Excel showed, in the ${kind} calculator`, async () => {
      const c = await calc()
      for (const id of IDS) {
        const s = await c.createSession(await book(id))
        expect((await s.editStructure({ op: 'refreshPivots' })).ok).toBe(true)
        expect([id, shown(await s.getLayout())]).toEqual([id, answers[id].phases.at(-1)!.slicers])
        await s.close()
      }
    })
})
