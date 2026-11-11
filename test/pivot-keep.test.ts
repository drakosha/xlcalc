/**
 * PivotTables over the file they were read from (`writeKept`): a book of Excel's written back without a change is the
 * file byte for byte; a PivotTable the session changed is written again from the model, its cache kept as the file
 * has it while unchanged; a new PivotTable gets parts of its own and its cache an entry of the workbook; a deleted one
 * goes with its parts and its cache. Each file read back gives the session's PivotTables.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { createSession, registerPivot, type BookInput } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'
import { readXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { partsOf } from './keep-files.js'

registerPivot(pivotEngine)

const file = new Uint8Array(readFileSync('test/excel/win/pivot/basic/columns.xlsx'))
const pivotsOf = (b: BookInput) => b.sheets.map((s) => s.pivotTables ?? [])

describe('writeKept over a file with a PivotTable', () => {
  it('writes the file back byte for byte without a change, and after a refresh that changes nothing laid out', async () => {
    const s = createSession((await openKept(file)).book)
    expect((await writeKept(file, s.getSnapshot())).data).toEqual(file)
  })

  it('writes a changed PivotTable again from the model, its cache as the file has it', async () => {
    const s = createSession((await openKept(file)).book)
    expect(s.editStructure({ op: 'editPivotTable', sheet: 'P', pivot: 'PT1', set: { rows: [1], columns: [0] } })).toMatchObject({ ok: true })
    const out = await writeKept(file, s.getSnapshot())
    const [before, after] = [await partsOf(file), await partsOf(out.data)]
    expect(after.get('xl/pivotTables/pivotTable1.xml')).not.toBe(before.get('xl/pivotTables/pivotTable1.xml'))
    expect(after.get('xl/pivotCache/pivotCacheRecords1.xml')).toBe(before.get('xl/pivotCache/pivotCacheRecords1.xml'))
    const back = await readXlsx(out.data)
    expect(pivotsOf(back.book)).toEqual(pivotsOf(s.getBook()))
    expect(back.book.sheets[1].cells).toEqual(s.getBook().sheets[1].cells)
  })

  it('adds a new PivotTable with parts of its own, and deletes one with its parts and cache', async () => {
    const s = createSession((await openKept(file)).book)
    expect(s.editStructure({ op: 'addPivotTable', sheet: 'G', cell: 'A3', source: { type: 'range', sheet: 'D', range: 'A1:J49' }, definition: { rows: [7], values: [{ field: 5 }] } })).toMatchObject({ ok: true })
    const added = await writeKept(file, s.getSnapshot())
    const parts = await partsOf(added.data)
    expect([...parts.keys()].filter((k) => /pivot/i.test(k)).sort()).toEqual([
      'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels', 'xl/pivotCache/_rels/pivotCacheDefinition2.xml.rels', 'xl/pivotCache/pivotCacheDefinition1.xml', 'xl/pivotCache/pivotCacheDefinition2.xml',
      'xl/pivotCache/pivotCacheRecords1.xml', 'xl/pivotCache/pivotCacheRecords2.xml', 'xl/pivotTables/_rels/pivotTable1.xml.rels', 'xl/pivotTables/_rels/pivotTable2.xml.rels', 'xl/pivotTables/pivotTable1.xml', 'xl/pivotTables/pivotTable2.xml',
    ])
    expect(parts.get('xl/workbook.xml')).toMatch(/<pivotCaches><pivotCache cacheId="104" r:id="rId4"\/><pivotCache cacheId="105" r:id="rId\d+"\/><\/pivotCaches>/)
    const back = await readXlsx(added.data)
    expect(pivotsOf(back.book)).toEqual(pivotsOf(s.getBook()))
    expect(back.book.pivotCaches?.map((c) => c.id)).toEqual([104, 105])
    s.editStructure({ op: 'deletePivotTable', sheet: 'P', pivot: 'PT1' })
    const gone = await writeKept(file, s.getSnapshot())
    const left = await partsOf(gone.data)
    expect([...left.keys()].filter((k) => /pivot/i.test(k) && !/_rels/.test(k)).sort()).toEqual(['xl/pivotCache/pivotCacheDefinition2.xml', 'xl/pivotCache/pivotCacheRecords2.xml', 'xl/pivotTables/pivotTable1.xml'])
    const again = await readXlsx(gone.data)
    expect([pivotsOf(again.book), again.book.pivotCaches?.map((c) => c.id)]).toEqual([pivotsOf(s.getBook()), [105]])
  })

  it('deletes the sheet of a PivotTable with its parts and cache; refuses to delete its source sheet, which Excel makes a link to the file', async () => {
    const s = createSession((await openKept(file)).book)
    s.editStructure({ op: 'deleteSheet', sheet: 'P' })
    const out = await writeKept(file, s.getSnapshot())
    const parts = await partsOf(out.data)
    expect([...parts.keys()].filter((k) => /pivot/i.test(k))).toEqual([])
    expect(parts.get('xl/workbook.xml')).not.toContain('pivotCache')
    const back = await readXlsx(out.data)
    expect([back.book.pivotCaches, back.issues.filter((i) => i.code === 'pivotTable')]).toEqual([undefined, []])
    const d = createSession((await openKept(file)).book)
    d.editStructure({ op: 'deleteSheet', sheet: 'D' })
    await expect(writeKept(file, d.getSnapshot())).rejects.toThrow('which Excel makes a link to the file by its path')
  })

  it("writes a slicer's cache whose choice changed through its part, and another slicer of that cache", async () => {
    const data = new Uint8Array(readFileSync('test/excel/win/pivot/slicer/row-two.xlsx'))
    const { book } = await openKept(data)
    const c = book.slicerCaches![0]
    const chosen = { ...book, slicerCaches: [{ ...c, items: c.items!.map(({ selected: _, ...x }) => x) }] }
    const out = await writeKept(data, { book: chosen })
    expect((await readXlsx(out.data)).book.slicerCaches).toEqual(chosen.slicerCaches)
    const [before, after] = [await partsOf(data), await partsOf(out.data)]
    // the cache's part, and the calculation chain gone as with any change
    expect([...after.keys()].filter((k) => after.get(k) !== before.get(k)).sort()).toEqual(['[Content_Types].xml', 'xl/_rels/workbook.xml.rels', 'xl/slicerCaches/slicerCache1.xml', 'xl/workbook.xml'])
    const s = book.sheets.findIndex((x) => x.slicers)
    const more = { ...book, sheets: book.sheets.map((x, i) => (i === s ? { ...x, slicers: [...x.slicers!, { ...x.slicers![0], name: 'Another' }] } : x)) }
    expect((await readXlsx((await writeKept(data, { book: more })).data)).book.sheets[s].slicers?.map((x) => x.name)).toEqual(more.sheets[s].slicers!.map((x) => x.name))
  })

  it("writes a slicer's changed choice and look on a renamed sheet through the sheet's parts", async () => {
    const data = new Uint8Array(readFileSync('test/excel/win/pivot/slicer/row-two.xlsx'))
    const s = createSession((await openKept(data)).book)
    const sheet = s.getBook().sheets.find((x) => x.slicers)!.name
    s.editStructure({ op: 'renameSheet', sheet, name: 'Renamed' })
    const b = s.getBook()
    const c = b.slicerCaches![0]
    const book = {
      ...b,
      slicerCaches: [{ ...c, items: c.items!.map(({ selected: _, ...x }) => x) }],
      sheets: b.sheets.map((x) => (x.slicers ? { ...x, slicers: x.slicers.map((y) => ({ ...y, caption: 'Regions' })) } : x)),
    }
    const out = await writeKept(data, { book })
    const back = (await readXlsx(out.data)).book
    expect([back.slicerCaches, back.sheets.find((x) => x.name === 'Renamed')?.slicers?.[0].caption]).toEqual([book.slicerCaches, 'Regions'])
  })

  it("writes a slicer's and a timeline's Locked and a slicer's Disable resizing and moving as Excel saved them, a renamed one's anchor again, the others byte for byte", async () => {
    // Excel set Locked off on both (locked-off) and Disable resizing and moving on the slicer (move-off), the set slicer-objects
    const unlocked = new Uint8Array(readFileSync('test/excel/win/pivot/slicer-objects/locked-off.xlsx'))
    const fixed = new Uint8Array(readFileSync('test/excel/win/pivot/slicer-objects/move-off.xlsx'))
    const own = (b: BookInput) => b.sheets.flatMap((x) => [...(x.slicers ?? []), ...(x.timelines ?? [])]).map((x) => [x.name, x.locked, 'positionLocked' in x ? x.positionLocked : 'none'])
    expect(own((await readXlsx(unlocked)).book)).toEqual([['Slicer_Region', false, 'none'], ['Timeline_Date', false, 'none']])
    expect(own((await readXlsx(fixed)).book)).toEqual([['Slicer_Region', undefined, true], ['Timeline_Date', undefined, 'none']])
    for (const data of [unlocked, fixed]) expect((await writeKept(data, { book: (await openKept(data)).book })).data).toEqual(data)
    // the slicer with both, Locked off and Disable resizing and moving on; the timeline renamed with neither
    const s = createSession((await openKept(fixed)).book)
    expect(s.editStructure({ op: 'editSlicer', slicer: 'Slicer_Region', set: { locked: false } }).ok).toBe(true)
    expect(s.editStructure({ op: 'editTimeline', timeline: 'Timeline_Date', set: { name: 'Dates' } }).ok).toBe(true)
    const out = await writeKept(fixed, s.getSnapshot())
    expect(own((await readXlsx(out.data)).book)).toEqual([['Slicer_Region', false, true], ['Dates', undefined, 'none']])
    const drawing = (await partsOf(out.data)).get('xl/drawings/drawing1.xml')!
    expect(drawing.match(/<xdr:clientData[^>]*>/g)).toEqual(['<xdr:clientData fLocksWithSheet="0"/>', '<xdr:clientData/>'])
    expect(drawing).toContain('<a:graphicFrameLocks noMove="1" noResize="1"/>')
    expect(drawing).toContain('<tsle:timeslicer name="Dates"/>')
    // a change of the timeline alone leaves the slicer's anchor as the file has it
    const t = createSession((await openKept(fixed)).book)
    t.editStructure({ op: 'editTimeline', timeline: 'Timeline_Date', set: { locked: false } })
    const one = (await partsOf((await writeKept(fixed, t.getSnapshot())).data)).get('xl/drawings/drawing1.xml')!
    const slicerAnchor = (x: string) => /<xdr:twoCellAnchor\b(?:(?!<\/xdr:twoCellAnchor>)[\s\S])*?sle:slicer[\s\S]*?<\/xdr:twoCellAnchor>/.exec(x)![0]
    expect(slicerAnchor(one)).toBe(slicerAnchor((await partsOf(fixed)).get('xl/drawings/drawing1.xml')!))
  })

  it('Excel for Windows opened the files writeXlsx and writeKept wrote with these locks, read them as written and saved them so (the set slicer-locks)', async () => {
    const answers = (JSON.parse(readFileSync('test/excel/win/pivot/slicer-locks.json', 'utf8')) as { answers: Record<string, { phases: { results?: Record<string, [unknown[], unknown]> }[] }> }).answers
    const WRITTEN = [['Slicer_Region', false, true], ['Products', true, false], ['Timeline_Date', false, null], ['Timeline_When', true, null]]
    for (const how of ['xlsx', 'kept']) {
      expect([how, Object.values(answers[how].phases.at(-1)!.results!).map((x) => x[0])]).toEqual([how, WRITTEN])
      const saved = (await readXlsx(readFileSync(`test/excel/win/pivot/slicer-locks/${how}.xlsx`))).book
      const own = saved.sheets.flatMap((x) => [...(x.slicers ?? []).map((y) => [y.name, y.locked !== false, !!y.positionLocked]), ...(x.timelines ?? []).map((y) => [y.name, y.locked !== false, null])])
      expect([how, own.sort()]).toEqual([how, [...WRITTEN].sort()])
    }
  })
})
