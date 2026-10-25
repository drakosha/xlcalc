/**
 * `xlcalc/xlsx/keep` against Excel, without it: the files Excel for Mac made with charts, a chart sheet, pictures, pivot
 * tables, formatted runs, theme colours, the print area, page breaks and notes (`test/excel/keep/made/`), and the files
 * `writeKept` wrote over them for five edits, which Excel opened and saved (`test/excel/keep/open/`, `open.json`, by
 * `bench/coverage/run.ts keep`). Each written file is the one Excel saw, Excel opened it without its repair window as
 * predicted (`bench/coverage/keep/open.predicted.json`), and the file it saved holds the same parts and the edit's values.
 * For the operations of the structure Excel did over the made files (`test/excel/keep/struct.json`, `struct/`), the
 * references `writeKept` writes are the ones Excel saved.
 */
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { readXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { createSession } from '../src/index.js'
import { MADE, partsOf, STRUCT } from '../bench/coverage/keep/books.js'
import { structWritten, written, type OpenRun, type StructRun } from '../bench/coverage/keep.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'

const dir = new URL('./excel/keep/', import.meta.url)
const run = JSON.parse(readFileSync(new URL('open.json', dir), 'utf8')) as OpenRun
const struct = JSON.parse(readFileSync(new URL('struct.json', dir), 'utf8')) as StructRun
/**
 * Where `writeKept` writes an operation's references otherwise than Excel, by case: what it writes. A series' category
 * or name that lost all its cells keeps `#REF!` in its place; Excel moves it to an extension of the series
 * (`c15:filteredCategoryTitle`, `c15:filteredSeriesTitle`), which no `<c:f>` holds.
 */
const KNOWN: Record<string, Record<string, unknown>> = {
  'chart-delete-series': { series: ["'Data'!#REF!", "'Data'!#REF!", "'Data'!$B$1"] },
  'chartsheet-source-deleted': { series: ['#REF!', '#REF!'] },
}
const predicted = JSON.parse(readFileSync(new URL('../bench/coverage/keep/open.predicted.json', import.meta.url), 'utf8')) as { cases: Record<string, { opens: boolean; parts: Record<string, number> }> }

/** A book's typed values by sheet and cell, the formulas left out: what Excel's file and the written one share whatever their spelling. */
const values = (book: Awaited<ReturnType<typeof readXlsx>>['book']) =>
  Object.fromEntries(book.sheets.map((s) => [s.name, Object.fromEntries(Object.entries(s.cells).flatMap(([a, c]) => (c.value !== undefined ? [[a, c.value]] : [])))]))

describe('xlsx/keep in Excel', () => {
  it('a made file opened and written without an edit is its own bytes', async () => {
    for (const m of MADE) {
      const data = readFileSync(new URL(`made/${m.id}.xlsx`, dir))
      const r = await openKept(data)
      const w = await writeKept(data, createSession(r.book, { cached: r.cached }).getSnapshot())
      expect(Buffer.compare(Buffer.from(w.data), data), m.id).toBe(0)
    }
  })

  it('Excel opened every file written over a made one without its repair window, and saved its charts, pictures and pivot tables', async () => {
    const files = await written()
    expect(files.map((f) => f.id).sort()).toEqual(Object.keys(predicted.cases).sort())
    for (const f of files) {
      const c = run.cases[f.id]
      expect(c, `${f.id}: no Excel answer: run the keep open step`).toBeDefined()
      expect(c.sent, `${f.id}: writeKept changed since Excel saw the file: run the keep open step`).toBe(await partsHash(f.data))
      expect(c.rejected, f.id).toBeUndefined()
      expect(c.written, f.id).toEqual(await partsOf(f.data))
      expect(c.saved, f.id).toEqual(predicted.cases[f.id].parts)
      // what the edit wrote is what Excel read and saved
      const saved = readFileSync(new URL(`open/${f.id}.xlsx`, dir))
      expect(c.saved, f.id).toEqual(await partsOf(saved))
      expect(values((await readXlsx(saved)).book), f.id).toEqual(values((await readXlsx(f.data)).book))
    }
  }, 60_000)

  it("writeKept moves the parts' references as Excel moved them in the operations of the structure", async () => {
    const ours = await structWritten()
    expect(Object.keys(ours).sort()).toEqual(Object.keys(struct.cases).sort())
    // each answer is Excel's for the case as it is now: its operation over its made file
    for (const c of STRUCT) {
      const data = readFileSync(new URL(`made/${c.made}.xlsx`, dir))
      expect(struct.cases[c.id].sent, `${c.id}: the case changed since Excel did it: run the keep struct step`).toBe(createHash('sha256').update(JSON.stringify(c.excel)).update(await partsHash(data)).digest('hex'))
    }
    for (const [id, o] of Object.entries(ours)) {
      expect('refused' in o ? o.refused : undefined, id).toBeUndefined()
      expect('refs' in o && o.refs, id).toEqual({ ...struct.cases[id].refs, ...KNOWN[id] })
    }
  }, 60_000)
})
