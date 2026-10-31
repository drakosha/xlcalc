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
})
