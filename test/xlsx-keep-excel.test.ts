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
 * workbook's AutoFilter after each operation is the one Excel saved. After the operations of `FORMATS` over a file with
 * formats the workbook does not tell apart (`formats.json`), each cell, row and column that moved has the cell format
 * Excel gave it, and Excel opened the files `writeKept` wrote and kept their formats.
 */
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { readXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { createSession, InputError } from '../src/index.js'
import { baseBook, MADE, MADE_PROBES, partsOf, PROBES, refsOf, STRUCT } from '../bench/coverage/keep/books.js'
import { formatsFiles, structFiles, structWritten, written, type FormatsRun, type MadeRun, type OpenRun, type StructOpenRun, type StructRun } from '../bench/coverage/keep.js'
import { FORMATS, formatsOf } from '../bench/coverage/keep/formats.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import { partsOf as textsOf } from './keep-files.js'

const dir = new URL('./excel/keep/', import.meta.url)
const run = JSON.parse(readFileSync(new URL('open.json', dir), 'utf8')) as OpenRun
const struct = JSON.parse(readFileSync(new URL('struct.json', dir), 'utf8')) as StructRun
/**
 * Where `writeKept` writes an operation's references otherwise than Excel, by case: what it writes. A series' category
 * or name that lost all its cells keeps `#REF!` in its place; Excel moves it to an extension of the series
 * (`c15:filteredCategoryTitle`, `c15:filteredSeriesTitle`), which no `<c:f>` holds. The filter database goes with the
 * workbook's AutoFilter where that moves; Excel moves it as a name, so it keeps it over cells moved to another sheet
 * without the AutoFilter.
 */
const KNOWN: Record<string, Record<string, unknown>> = {
  'chart-delete-series': { series: ["'Data'!#REF!", "'Data'!#REF!", "'Data'!$B$1"] },
  'chartsheet-source-deleted': { series: ['#REF!', '#REF!'] },
  'filter-move-sheet': { names: [] },
  'pivot-delete-source-cells': { series: ["'Data'!#REF!", "'Data'!#REF!", "'Data'!#REF!"] },
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

  it('Excel answers an AutoFilter over a table and a table over an AutoFilter as the engine does: it refuses the first, and the second but over all of the AutoFilter, where it takes its place', () => {
    const made = JSON.parse(readFileSync(new URL('made.json', dir), 'utf8')) as MadeRun
    for (const m of MADE_PROBES) {
      const answer = made.books[m.id]?.made
      expect(answer, `${m.id}: no Excel answer: run the keep made step`).toBeDefined()
      expect(answer.map((line, i) => line.slice(0, PROBES[m.id][i]?.length)), m.id).toEqual(PROBES[m.id])
    }
    // the engine's answers: the AutoFilter over the table an input error, a table over all of the AutoFilter in its place, over part of it refused
    const book = baseBook()
    book.sheets[2].tables = [{ name: 'Plain', range: 'A1:A2', columns: [{ name: 'plain' }] }]
    book.sheets[0].autoFilter = { range: 'A1:C21' }
    expect(() => createSession(book).setLayout([{ sheet: 'Plain', autoFilter: { range: 'A1:B5' } }])).toThrow(InputError)
    for (const [range, width, made] of [['A1:C21', 3, true], ['B10:D12', 3, false], ['A1:B10', 2, false], ['A1:D22', 4, true], ['A1:C25', 3, true], ['A2:C21', 3, false]] as const) {
      const s = createSession(book)
      const columns = Array.from({ length: width }, (_, i) => ({ name: `c${i}` }))
      expect(s.editStructure({ op: 'addTable', sheet: 'Data', table: { name: 'T', range, columns } }).ok, range).toBe(made)
      expect(s.getBook().sheets[0].autoFilter === undefined, range).toBe(made)
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

  it('the cells, rows and columns that moved have the cell formats Excel gave them, and Excel kept them in the files writeKept wrote', async () => {
    const answers = JSON.parse(readFileSync(new URL('formats.json', dir), 'utf8')) as FormatsRun
    const predicted = JSON.parse(readFileSync(new URL('../bench/coverage/keep/formats.predicted.json', import.meta.url), 'utf8')) as { cases: Record<string, { formats: Awaited<ReturnType<typeof formatsOf>> }> }
    const made = new Uint8Array(readFileSync(new URL('made/formats.xlsx', dir)))
    const files = await formatsFiles()
    expect(files.map((f) => f.id)).toEqual(FORMATS.map((c) => c.id))
    for (const f of files) {
      const c = answers.cases[f.id]
      const id = FORMATS.find((x) => x.id === f.id)!
      expect(c?.sent, `${f.id}: the case changed since Excel did it: run the keep formats step`).toBe(createHash('sha256').update(JSON.stringify(id.excel)).update(await partsHash(made)).digest('hex'))
      expect(c.open.sent, `${f.id}: writeKept changed since Excel saw its file: run the keep formats step`).toBe(await partsHash(f.data))
      const ours = await formatsOf(f.data)
      expect(ours, f.id).toEqual(predicted.cases[f.id].formats)
      const differ = [...new Set([...Object.keys(ours.at), ...Object.keys(c.formats.at)])].filter((a) => ours.at[a] !== c.formats.at[a]).sort()
      expect(differ, f.id).toEqual([...(FORMATS_KNOWN[f.id] ?? [])].sort())
      // each format as content, the same in both files
      for (const a of Object.keys(ours.at)) if (c.formats.at[a] === ours.at[a]) expect(ours.formats[ours.at[a]]).toBe(c.formats.formats[ours.at[a]])
      expect(c.open.rejected, f.id).toBeUndefined()
      expect(c.open.formats, f.id).toEqual(ours)
      expect(await formatsOf(readFileSync(new URL(`formats-open/${f.id}.xlsx`, dir))), f.id).toEqual(ours)
      expect(await formatsOf(readFileSync(new URL(`formats/${f.id}.xlsx`, dir))), f.id).toEqual(c.formats)
    }
  }, 60_000)
})

/**
 * Where `writeKept` gives cells, rows and columns other formats than Excel after an operation of `FORMATS`, by case.
 * Excel gives a cell, row or column it inserts the format of the line before it, the same cell format; `writeKept` gives
 * a new cell inside a block of a format that format, but a cell of a plain look (a quote prefix, a cell style, a pattern
 * fill) is not written (A8, B6:B9, C17:C18), and a row or column inserted after a block takes the first format of its
 * look (row 26, column 9, B14:C14). A cell that looks as its row and its column (G25, H23 of the made file) is left out
 * where rows or columns moved, kept as the file has it where the workbook's rows are the same.
 */
const FORMATS_KNOWN: Record<string, string[]> = {
  'formats-insert-above': ['G27', 'H25'],
  'formats-insert-in': ['A8', 'G26', 'H24'],
  'formats-delete-in': ['G23', 'H21'],
  'formats-insert-column': ['B6', 'B7', 'B8', 'B9', 'H25', 'I23'],
  'formats-delete-column': ['F25', 'G23'],
  'formats-insert-at-row': ['G26', 'H23'],
  'formats-insert-below-row': ['G25', 'G26', 'H23', 'row 26'],
  'formats-insert-after-column': ['I23', 'column 9'],
  'formats-insert-in-colour': ['G26', 'H24'],
  'formats-insert-in-blocks': ['C17', 'C18', 'H25', 'I23'],
  'formats-delete-across': ['G23', 'H21'],
  'formats-insert-below-block': ['B14', 'C14', 'G26', 'H24'],
  'formats-delete-colour-column': ['H23'],
}
