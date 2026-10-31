/**
 * Σ Values on a PivotTable's axes when a workbook gives none (a host's workbook, one saved before 0.72): with
 * `xlcalc/pivot` loaded, a PivotTable of two value fields or more is taken in with it where adding the second value
 * field puts it (the end of the columns, or of the rows with `valuesOnRows`, at `valuesPosition` when it has one), in a
 * session of each execution, the inline calculator and a pool's threads; `getBook` gives it back and `writeXlsx` writes
 * it on its axis (`<field x="-2"/>`), as Excel requires with two value fields. A workbook with none such stays itself.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { createInlineCalc, createSession, registerPivot, type BookInput, type PivotTable } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'
import { bookTaken } from '../src/taken.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { closedAfter } from './fixtures/pool.js'

registerPivot(pivotEngine)

/** Books Excel saved with two value fields: Σ Values at the columns' end, first on the columns, in the middle of the rows. */
const BOOKS = ['second-after-columns', 'second-no-columns', 'columns-first-drop-back', 'rows-middle-field-off']
const read = async (id: string): Promise<BookInput> => (await readXlsx(readFileSync(new URL(`./excel/win/pivot/values-field/${id}.xlsx`, import.meta.url)))).book
/** The book with Σ Values on no axis, its place kept (`valuesOnRows`, `valuesPosition`). */
const without = (b: BookInput): BookInput => ({ ...b, sheets: b.sheets.map((s) => (s.pivotTables ? { ...s, pivotTables: s.pivotTables.map((t) => ({ ...t, rows: t.rows?.filter((f) => f !== 'values'), columns: t.columns?.filter((f) => f !== 'values') })) } : s)) })
const axes = (b: BookInput) => b.sheets.flatMap((s) => (s.pivotTables ?? []).map((t: PivotTable) => [t.rows, t.columns]))

/** The `<field x>` of the row and column fields of each PivotTable part a file holds. */
async function written(bytes: Uint8Array): Promise<string[]> {
  const a = openArchive(bytes)
  const out: string[] = []
  for (const n of a.names.filter((x: string) => /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(x)).sort()) {
    const xml = (await a.text(n)) ?? ''
    for (const tag of ['rowFields', 'colFields']) out.push(`${tag}: ${[...(new RegExp(`<${tag}[^>]*>([\\s\\S]*?)</${tag}>`).exec(xml)?.[1] ?? '').matchAll(/x="(-?\d+)"/g)].map((m) => m[1]).join(' ')}`)
  }
  return out
}

describe('Σ Values taken in where a workbook gives none', () => {
  it.each(BOOKS)('a session over %s gives it back where Excel kept it, and writes it', async (id) => {
    const excel = await read(id)
    for (const execution of ['closures', 'eval'] as const) {
      const s = createSession(without(excel), { execution })
      const book = s.getBook()
      expect(axes(book)).toEqual(axes(excel))
      expect(await written((await writeXlsx({ book })).data)).toEqual(await written(new Uint8Array(readFileSync(new URL(`./excel/win/pivot/values-field/${id}.xlsx`, import.meta.url)))))
      expect(s.getValues('P', 'A1:F20')).toEqual(createSession(excel, { execution }).getValues('P', 'A1:F20'))
    }
  })

  const { made } = closedAfter()
  for (const [kind, calc] of [['inline', async () => createInlineCalc()], ['pool', () => made({ threads: 1 })]] as const)
    it(`in the ${kind} calculator`, async () => {
      const c = await calc()
      for (const id of BOOKS) {
        const excel = await read(id)
        const s = await c.createSession(without(excel))
        expect(axes((await s.getBook()) as BookInput)).toEqual(axes(excel))
      }
    })

  it('takes a workbook with none such as it is', async () => {
    const one = await read('second-no-columns-1')
    const marked = await read('second-no-columns')
    expect([bookTaken(one, []) === one, bookTaken(marked, []) === marked]).toEqual([true, true])
  })
})
