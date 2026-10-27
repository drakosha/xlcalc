/**
 * `xlcalc/xlsx/keep` against Excel, without it: the files Excel for Mac made with charts, a chart sheet, pictures, pivot
 * tables, formatted runs, theme colours, the print area, page breaks and notes (`test/excel/keep/made/`), and the files
 * `writeKept` wrote over them for five edits, which Excel opened and saved (`test/excel/keep/open/`, `open.json`, by
 * `bench/coverage/run.ts keep`). Each written file is the one Excel saw, Excel opened it without its repair window as
 * predicted (`bench/coverage/keep/open.predicted.json`), and the file it saved holds the same parts and the edit's values.
 * For the operations of the structure Excel did over the made files (`test/excel/keep/struct.json`, `struct/`), the
 * references `writeKept` writes are the ones Excel saved; Excel opened the files `writeKept` wrote for them without its
 * repair window and kept their parts and references, as predicted (`struct-open.json`, `struct-open/`). Excel answered
 * an AutoFilter over a table and a table over an AutoFilter as the engine does (`made.json`, `PROBES`), and the
 * workbook's AutoFilter after each operation is the one Excel saved.
 */
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { readXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { createSession, InputError } from '../src/index.js'
import { baseBook, MADE, MADE_PROBES, partsOf, PROBES, refsOf, STRUCT } from '../bench/coverage/keep/books.js'
import { structFiles, structWritten, written, type MadeRun, type OpenRun, type StructOpenRun, type StructRun } from '../bench/coverage/keep.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import { partsOf as textsOf } from './keep-files.js'

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
/**
 * The operations `writeKept` refuses, by case: the refusal. A pivot cache whose source sheet is deleted: Excel keeps
 * `ref` and `sheet` and links the source to the file by its path (`r:id`, an external link path to the file's name),
 * which a writer of bytes does not know; the source kept without the link, Excel does not open the file.
 */
const REFUSED: Record<string, string> = {
  'pivot-source-deleted': 'xlsx/keep: the source of the pivot cache xl/pivotCache/pivotCacheDefinition1.xml is on a sheet the change deleted, which Excel makes a link to the file by its path',
}
const structOpen = JSON.parse(readFileSync(new URL('struct-open.json', dir), 'utf8')) as StructOpenRun
const structPredicted = JSON.parse(readFileSync(new URL('../bench/coverage/keep/struct-open.predicted.json', import.meta.url), 'utf8')) as { cases: Record<string, { parts: Record<string, number>; refs: Record<string, unknown> }> }
/**
 * Where Excel saved a file `writeKept` wrote for an operation of the structure otherwise than predicted, by case: what
 * it saved. Excel drops the formats and strings no cell uses, as with its own operation, and moves a series' category
 * or name that lost its cells to an extension of the series.
 */
const RESAVED: Record<string, { parts?: Record<string, number>; refs?: Record<string, unknown> }> = {
  'chart-delete-in': { parts: { themeColors: 0 } },
  'chart-delete-series': { parts: { themeColors: 0 }, refs: { series: ["'Data'!#REF!", "'Data'!$B$1"] } },
  'chart-delete-chart': { parts: { themeColors: 0 } },
  'chartsheet-source-deleted': { parts: { richStrings: 0 }, refs: { series: ['#REF!'] } },
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
      expect('refused' in o ? o.refused : undefined, id).toBe(REFUSED[id])
      if ('refs' in o) expect(o.refs, id).toEqual({ ...struct.cases[id].refs, ...KNOWN[id] })
    }
    // Excel's file for the refused one: the source kept, linked to the file by its name in Excel's sandbox
    const saved = await textsOf(readFileSync(new URL('struct/pivot-source-deleted.xlsx', dir)))
    expect(/<worksheetSource ref="A1:C21" sheet="Data" r:id="rId1"\/>/.test(saved.get('xl/pivotCache/pivotCacheDefinition1.xml')!)).toBe(true)
    expect(saved.get('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels')).toContain('relationships/externalLinkPath" Target="book.xlsx" TargetMode="External"')
  }, 60_000)

  it('Excel answers an AutoFilter over a table and a table over an AutoFilter as the engine does: it refuses the first, and the second takes its place', () => {
    const made = JSON.parse(readFileSync(new URL('made.json', dir), 'utf8')) as MadeRun
    for (const m of MADE_PROBES) {
      const answer = made.books[m.id]?.made
      expect(answer, `${m.id}: no Excel answer: run the keep made step`).toBeDefined()
      expect(answer.map((line, i) => line.slice(0, PROBES[m.id][i]?.length)), m.id).toEqual(PROBES[m.id])
    }
    // the engine's answers: the AutoFilter over the table an input error, a table over the AutoFilter or part of it takes its place
    const book = baseBook()
    book.sheets[2].tables = [{ name: 'Plain', range: 'A1:A2', columns: [{ name: 'plain' }] }]
    book.sheets[0].autoFilter = { range: 'A1:C21' }
    expect(() => createSession(book).setLayout([{ sheet: 'Plain', autoFilter: { range: 'A1:B5' } }])).toThrow(InputError)
    for (const [range, columns] of [['A1:C21', ['Region', 'Amount', 'Month']], ['B10:D12', ['70', 'M2', 'Column1']]] as const) {
      const s = createSession(book)
      expect(s.editStructure({ op: 'addTable', sheet: 'Data', table: { name: 'T', range, columns: columns.map((name) => ({ name })) } }).ok, range).toBe(true)
      expect(s.getBook().sheets[0].autoFilter, range).toBeUndefined()
      s.close()
    }
  })

  it("the workbook's AutoFilter after each operation of the structure is the one Excel saved", async () => {
    for (const c of STRUCT.filter((x) => x.made === 'filter')) {
      const read = await openKept(readFileSync(new URL(`made/${c.made}.xlsx`, dir)))
      const s = createSession(read.book, { cached: read.cached })
      expect(s.editStructure(c.op).ok, c.id).toBe(true)
      const ours = Object.fromEntries(s.getBook().sheets.flatMap((x) => (x.autoFilter ? [[x.name, x.autoFilter.range]] : [])))
      expect(ours, c.id).toEqual(struct.cases[c.id].refs?.filters ?? {})
      s.close()
    }
  })

  it('Excel opened every file writeKept wrote for the operations of the structure, and kept its parts and their references', async () => {
    const files = await structFiles()
    expect(files.map((f) => f.id).sort()).toEqual(Object.keys(structPredicted.cases).sort())
    for (const f of files) {
      if (!('data' in f.written)) {
        expect(f.written.refused, f.id).toBe(REFUSED[f.id])
        continue
      }
      const c = structOpen.cases[f.id]
      expect(c, `${f.id}: no Excel answer: run the keep struct-open step`).toBeDefined()
      expect(c.sent, `${f.id}: writeKept changed since Excel saw the file: run the keep struct-open step`).toBe(await partsHash(f.written.data))
      expect(c.rejected, f.id).toBeUndefined()
      const p = structPredicted.cases[f.id]
      expect(c.saved, f.id).toEqual({ ...p.parts, ...RESAVED[f.id]?.parts })
      expect(c.refs, f.id).toEqual({ ...p.refs, ...RESAVED[f.id]?.refs })
      const saved = readFileSync(new URL(`struct-open/${f.id}.xlsx`, dir))
      expect(await partsOf(saved), f.id).toEqual(c.saved)
      expect(await refsOf(saved), f.id).toEqual(c.refs)
    }
  }, 60_000)
})
