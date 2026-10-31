/**
 * A pivot cache's times of day without a date as an older reader left them in a workbook a host saved (`getBook`, a
 * snapshot): on Excel's 1899-12-30 (`{ date: '1899-12-30T08:30:00' }`) in items, records, facts and a group's range.
 * Taken in as the model holds them (`1900-01-00T08:30:00`), so a session over such a book labels, lays out, refreshes
 * and groups as over the book read now, in a session of each execution, the inline calculator and a pool's threads;
 * a date before day 0 even so is refused on input with its path.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { createInlineCalc, createSession, InputError, registerPivot, type BookInput, type StructureOp } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'
import { readXlsx } from '../src/xlsx.js'
import { bookTaken } from '../src/taken.js'
import { closedAfter } from './fixtures/pool.js'

registerPivot(pivotEngine)

/** The book Excel saved with a field of times (Time, on the rows, with a slicer) and one of times among dates (Clock, records alone). */
const now = async (): Promise<BookInput> => (await readXlsx(readFileSync(new URL('./excel/win/pivot/slicer-labels/time.xlsx', import.meta.url)))).book
/** The same book as 0.71 kept it: every time of day on 1899-12-30, a group's range too. */
const old = (b: BookInput): BookInput => JSON.parse(JSON.stringify(b).replaceAll('"1900-01-00T', '"1899-12-30T'))
/** Time grouped by hours from 06:00, its range's start a time of day. */
const HOURS: StructureOp = { op: 'groupPivotField', sheet: 'P', pivot: 'PT1', field: 8, group: { by: 'dates', periods: ['hours'], start: '1900-01-00T06:00:00' } }

/** What a session over the book shows: the slicer's labels, the PivotTable after a refresh, and after Time grouped by hours. */
async function seen(s: { getLayout(): unknown; editStructure(op: StructureOp): unknown; getValues(sheet: string, range: string): unknown; getBook(): unknown }) {
  const layout = (await s.getLayout()) as { slicerCaches: { items: { label: string }[] }[] }
  const labels = layout.slicerCaches[0].items.map((x) => x.label)
  expect(((await s.editStructure({ op: 'refreshPivots' })) as { ok: boolean }).ok).toBe(true)
  const refreshed = await s.getValues('P', 'A3:B11')
  const grouped = (await s.editStructure(HOURS)) as { ok: boolean; message?: string }
  const book = (await s.getBook()) as BookInput
  return { labels, refreshed, grouped: grouped.ok ? await s.getValues('P', 'A3:B11') : grouped.message, cache: book.pivotCaches![0].fields[8] }
}

describe('a cache with times of day on 1899-12-30, as a host saved it with 0.71', () => {
  it('is the book read now: the field 8 is Time, its items, facts and records on 1899-12-30 in the old form', async () => {
    const b = old(await now())
    const pc = b.pivotCaches![0]
    expect([pc.fields[8].name, pc.fields[8].items![0], pc.fields[8].facts!.min, pc.records![0][10]]).toEqual(['Time', { date: '1899-12-30T06:00:00' }, { date: '1899-12-30T01:30:00' }, { date: '1899-12-30T08:29:46' }])
  })

  for (const execution of ['closures', 'eval'] as const)
    it(`labels, lays out, refreshes and groups as the book read now, in a session (${execution})`, async () => {
      const want = await seen(createSession(await now(), { execution }))
      expect(want.labels).toEqual(['1:30', '6:00', '12:00', '13:30', '18:00', '23:58'])
      expect(typeof want.grouped).not.toBe('string')
      const s = createSession(old(await now()), { execution })
      // taken in as the model holds them: the book given back has no 1899-12-30
      expect(JSON.stringify(s.getBook())).not.toContain('1899-12-30')
      expect(await seen(s)).toEqual(want)
    })

  const { made } = closedAfter()
  for (const [kind, calc] of [['inline', async () => createInlineCalc()], ['pool', () => made({ threads: 1 })]] as const)
    it(`labels, lays out, refreshes and groups as the book read now, in the ${kind} calculator`, async () => {
      const c = await calc()
      const want = await seen(await c.createSession(await now()))
      expect(await seen(await c.createSession(old(await now())))).toEqual(want)
    })

  it('takes a book with none as it is, its records not copied', async () => {
    const b = await now()
    expect(bookTaken(b, [])).toBe(b)
    const o = old(b)
    const t = bookTaken(o, [])
    // the records of the cache hold one in Clock (the field without items): those of other rows stay the same arrays
    const [r0, r1] = [o.pivotCaches![0].records!, t.pivotCaches![0].records!]
    expect([r1[0] === r0[0], r1[0][10], t.pivotCaches![0].fields[0] === o.pivotCaches![0].fields[0]]).toEqual([false, { date: '1900-01-00T08:29:46' }, true])
    expect(r1.filter((r, i) => r === r0[i]).length).toBe(r0.filter((r) => !JSON.stringify(r).includes('1899-12-30')).length)
  })

  it('refuses on input a date before day 0 even so, with its path', async () => {
    const b = old(await now())
    const pc = b.pivotCaches![0]
    const bad: BookInput = { ...b, pivotCaches: [{ ...pc, fields: pc.fields.map((f, i) => (i === 8 ? { ...f, items: [{ date: '1899-12-29T06:00:00' }, ...f.items!.slice(1)] } : f)) }] }
    expect(() => createSession(bad)).toThrow(new InputError('book.pivotCaches[0].fields[8].items[0].date', `"1899-12-29T06:00:00" is before the 1900 date system's day 0 (1900-01-00): Excel has no such date`))
    const later = { ...b, dateSystem: 1904 as const }
    expect(() => createSession(later)).toThrow(InputError)
  })
})
