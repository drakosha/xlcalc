/**
 * `xlcalc/xlsx/keep` against Excel, without it: the files Excel for Mac made with charts, a chart sheet, pictures, pivot
 * tables, formatted runs, theme colours, the print area, page breaks and notes (`test/excel/keep/made/`), and the files
 * `writeKept` wrote over them for five edits, which Excel opened and saved (`test/excel/keep/open/`, `open.json`, by
 * `bench/coverage/run.ts keep`). Each written file is the one Excel saw, Excel opened it without its repair window as
 * predicted (`bench/coverage/keep/open.predicted.json`), and the file it saved holds the same parts and the edit's values.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { readXlsx } from '../src/xlsx.js'
import { KeptRefusal, openKept, writeKept } from '../src/xlsx-keep.js'
import { createSession } from '../src/index.js'
import { MADE, partsOf } from '../bench/coverage/keep/books.js'
import { written, type OpenRun } from '../bench/coverage/keep.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'

const dir = new URL('./excel/keep/', import.meta.url)
const run = JSON.parse(readFileSync(new URL('open.json', dir), 'utf8')) as OpenRun
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

  it("a change that would move the references of Excel's parts is refused, with the part named", async () => {
    const refusal = async (id: string, op: Parameters<ReturnType<typeof createSession>['editStructure']>[0]) => {
      const data = readFileSync(new URL(`made/${id}.xlsx`, dir))
      const r = await openKept(data)
      const s = createSession(r.book, { cached: r.cached })
      expect(s.editStructure(op).ok).toBe(true)
      return writeKept(data, s.getSnapshot()).then(
        () => 'written',
        (e: Error) => (e instanceof KeptRefusal ? e.message.replace(/ \(xl\/[^)]*\)|xl\/[\w/]+\.xml/g, '').slice(0, 60) : `${e.name}: ${e.message}`),
      )
    }
    // a row inserted above what the chart reads moves its anchor first: the chart sits over the data's rows
    expect(await refusal('chart', { op: 'insertRows', sheet: 'Data', row: 3 })).toMatch(/^xlsx\/keep: a drawing on 'Data' refers to 'Data'!\$F\$1:\$L\$13, /)
    expect(await refusal('pivot', { op: 'deleteRows', sheet: 'Data', row: 21 })).toMatch(/^xlsx\/keep: (the chart|the source of the pivot cache)/)
    expect(await refusal('print', { op: 'insertColumns', sheet: 'Data', column: 'B' })).toMatch(/^xlsx\/keep: /)
    expect(await refusal('chartsheet', { op: 'renameSheet', sheet: 'Data', name: 'Amounts' })).toMatch(/^xlsx\/keep: the chart  refers to 'Data'!/)
    expect(await refusal('all', { op: 'moveRange', from: { sheet: 'Data', range: 'A1:C21' }, to: { sheet: 'Plain', cell: 'D1' } })).toMatch(/^xlsx\/keep: /)
    expect(await refusal('all', { op: 'deleteSheet', sheet: 'Pivot' })).toMatch(/^xlsx\/keep: /)
  })
})
