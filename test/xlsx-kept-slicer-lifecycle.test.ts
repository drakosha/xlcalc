/** Slicers and timelines added to or removed from an original file, with its other drawing objects kept. */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { createSession, registerPivot, type BookInput } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'
import { readXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { zip } from '../src/xlsx/zip/write.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { withParts, partsOf } from './keep-files.js'

registerPivot(pivotEngine)
const load = (path: string) => new Uint8Array(readFileSync(`test/excel/win/${path}.xlsx`))
const objects = (b: BookInput) => b.sheets.flatMap((s) => [...(s.slicers ?? []), ...(s.timelines ?? [])].map((x) => [s.name, x.name, x.cache]))

describe('writeKept slicer and timeline lifecycle', () => {
  it('adds a second slicer of one cache and removes its first without losing the second', async () => {
    const original = load('table-slicer/plain')
    const s = createSession((await openKept(original)).book)
    const sheet = s.getBook().sheets.find((x) => x.slicers)!
    expect(
      s.editStructure({ op: 'addSlicer', sheet: sheet.name, cache: s.getBook().slicerCaches![0].name, slicer: { name: 'Another', place: { from: 'J2', to: 'L10' } } }).ok,
    ).toBe(true)
    const first = await writeKept(original, s.getSnapshot())
    expect(objects((await readXlsx(first.data)).book)).toEqual(objects(s.getBook()))
    expect(s.editStructure({ op: 'deleteSlicer', slicer: sheet.slicers![0].name }).ok).toBe(true)
    const second = await writeKept(original, s.getSnapshot())
    expect(objects((await readXlsx(second.data)).book)).toEqual(objects(s.getBook()))
    const parts = await partsOf(second.data)
    expect(parts.get('xl/drawings/drawing1.xml')).toContain('name="Another"')
    expect(parts.get('xl/drawings/drawing1.xml')).not.toContain(`name="${sheet.slicers![0].name}"`)
    s.close()
  })

  it.each(['table-slicer/plain', 'pivot/slicer/row-two', 'pivot/slicer-objects/locked-off'])(
    'removes all slicers and timelines with their parts, relationships and names: %s',
    async (path) => {
      const original = load(path)
      const s = createSession((await openKept(original)).book)
      for (const sheet of s.getBook().sheets) {
        for (const x of sheet.slicers ?? []) expect(s.editStructure({ op: 'deleteSlicer', slicer: x.name }).ok).toBe(true)
        for (const x of sheet.timelines ?? []) expect(s.editStructure({ op: 'deleteTimeline', timeline: x.name }).ok).toBe(true)
      }
      const out = await writeKept(original, s.getSnapshot())
      const back = (await readXlsx(out.data)).book
      expect(objects(back)).toEqual([])
      expect([back.slicerCaches ?? [], back.timelineCaches ?? []]).toEqual([[], []])
      const parts = await partsOf(out.data)
      expect([...parts.keys()].filter((p) => /(?:slicer|timeline)s?\//i.test(p))).toEqual([])
      for (const [p, text] of parts)
        if (p.endsWith('.rels') || p === 'xl/workbook.xml' || p === '[Content_Types].xml')
          expect(text).not.toMatch(/(?:slicer|timeline)(?:Cache|Ref|List)|name="Slicer_|name="Timeline_/)
      s.close()
    },
  )

  it('adds a new table cache using the original column ids, even after inserting a column', async () => {
    const original = load('table-slicer/plain')
    const s = createSession((await openKept(original)).book)
    const table = s.getBook().sheets.find((x) => x.tables)?.tables![0]
    expect(s.editStructure({ op: 'insertColumns', sheet: 'Data', column: 'B', count: 1 }).ok).toBe(true)
    expect(s.editStructure({ op: 'addSlicer', sheet: 'Data', cache: { table: table!.name, column: 'Amount' }, slicer: { name: 'Amounts' } }).ok).toBe(true)
    const out = await writeKept(original, s.getSnapshot())
    const back = (await readXlsx(out.data)).book
    expect(back.slicerCaches?.find((c) => c.field === 'Amount')?.table).toEqual({ name: table!.name, column: 'Amount' })
    expect(objects(back)).toEqual(objects(s.getBook()))
    s.close()
  })

  it('adds a slicer and timeline to a sheet with no drawing or slicer parts', async () => {
    const original = load('pivot/slicer-objects/locked-off')
    const s = createSession((await openKept(original)).book)
    expect(s.editStructure({ op: 'insertSheet', name: 'New' }).ok).toBe(true)
    expect(s.editStructure({ op: 'addSlicer', sheet: 'New', cache: s.getBook().slicerCaches![0].name, slicer: { name: 'OnNew', place: { from: 'A1', to: 'D9' } } }).ok).toBe(true)
    expect(
      s.editStructure({ op: 'addTimeline', sheet: 'New', cache: s.getBook().timelineCaches![0].name, timeline: { name: 'NewDates', place: { from: 'E1', to: 'J9' } } }).ok,
    ).toBe(true)
    const out = await writeKept(original, s.getSnapshot())
    expect(objects((await readXlsx(out.data)).book)).toEqual(objects(s.getBook()))
    s.close()
  })
  it('keeps chart and picture parts and their shifted anchors while adding a table slicer', async () => {
    const original = await withParts(
      { sheets: [{ name: 'S', cells: { A1: { value: 'Item' }, A2: { value: 'Apple' } }, tables: [{ name: 'Sales', range: 'A1:A2', columns: [{ name: 'Item' }] }] }] },
      { chart: { sheet: 0, series: ['S!A2'], anchor: [2, 5, 10, 8] }, picture: { sheet: 0, anchor: [12, 5, 20, 8] } },
    )
    const s = createSession((await openKept(original)).book)
    expect(s.editStructure({ op: 'insertRows', sheet: 'S', row: 1, count: 2 }).ok).toBe(true)
    const moved = await writeKept(original, s.getSnapshot())
    expect(s.editStructure({ op: 'addSlicer', sheet: 'S', cache: { table: 'Sales', column: 'Item' }, slicer: { name: 'Items' } }).ok).toBe(true)
    const out = await writeKept(original, s.getSnapshot())
    const [before, after] = [await partsOf(moved.data), await partsOf(out.data)]
    for (const part of ['xl/charts/chart1.xml', 'xl/media/image1.png', 'xl/drawings/_rels/drawing1.xml.rels']) expect(after.get(part)).toBe(before.get(part))
    const anchors = (xml: string) => xml.match(/<xdr:twoCellAnchor\b[\s\S]*?<\/xdr:twoCellAnchor>/g) ?? []
    expect(anchors(after.get('xl/drawings/drawing1.xml')!).slice(0, 2)).toEqual(anchors(before.get('xl/drawings/drawing1.xml')!))
    expect(objects((await readXlsx(out.data)).book)).toEqual(objects(s.getBook()))
    s.close()
  })

  it('creates new pivot slicer and timeline caches, with the source cache id the original file uses', async () => {
    const original = load('pivot/slicer-objects/locked-off')
    const s = createSession((await openKept(original)).book)
    const b = s.getBook()
    for (const [op, c, name] of [
      ['addSlicer', b.slicerCaches![0], 'NewRegion'],
      ['addTimeline', b.timelineCaches![0], 'NewDate'],
    ] as const) {
      const where = { field: c.field, pivots: c.pivots }
      const object = { name, place: { from: 'M2', to: 'R10' } }
      expect(
        s.editStructure(op === 'addSlicer' ? { op, sheet: c.pivots[0].sheet, cache: where, slicer: object } : { op, sheet: c.pivots[0].sheet, cache: where, timeline: object }).ok,
      ).toBe(true)
    }
    const out = await writeKept(original, s.getSnapshot())
    const back = (await readXlsx(out.data)).book
    expect(objects(back)).toEqual(objects(s.getBook()))
    expect(back.slicerCaches?.map((c) => [c.name, c.cache])).toEqual(s.getBook().slicerCaches?.map((c) => [c.name, c.cache]))
    expect(back.timelineCaches?.map((c) => [c.name, c.cache])).toEqual(s.getBook().timelineCaches?.map((c) => [c.name, c.cache]))
    s.close()
  })

  it('reads the six synthetic lifecycle books Windows Excel opened and saved, including the corrected shared timeline', async () => {
    const expected = JSON.parse(readFileSync(new URL('./excel/win/keep/slicer-lifecycle.json', import.meta.url), 'utf8')).expected
    for (const [id, want] of Object.entries(expected)) {
      const b = (await readXlsx(readFileSync(new URL(`./excel/win/keep/slicer-lifecycle/${id}.xlsx`, import.meta.url)))).book
      const caches = [...(b.slicerCaches ?? []), ...(b.timelineCaches ?? [])]
        .map((c) => ({
          name: c.name,
          source: c.field,
          slicers: b.sheets
            .flatMap((s) => [...(s.slicers ?? []), ...(s.timelines ?? [])].filter((x) => x.cache === c.name).map((x) => `${s.name}!${x.name}:${x.caption ?? x.name}`))
            .sort(),
        }))
        .sort((a, b) => a.name.localeCompare(b.name))
      expect({ opened: true, caches }, id).toEqual(want)
    }
  })
  it('keeps an unread table cache in the table extensions when another slicer is added', async () => {
    const archive = openArchive(load('table-slicer/plain'))
    const original = await zip(
      await Promise.all(
        archive.names.map(async (name) => {
          const bytes = (await archive.bytes(name))!
          return {
            name,
            data: name === 'xl/slicerCaches/slicerCache1.xml' ? new TextEncoder().encode(new TextDecoder().decode(bytes).replace('tableId="1"', 'tableId="777"')) : bytes,
          }
        }),
      ),
    )
    const s = createSession((await openKept(original)).book)
    expect(s.getBook().slicerCaches![0].table).toBeUndefined()
    expect(s.editStructure({ op: 'addSlicer', sheet: 'Data', cache: { table: 'Sales', column: 'Amount' }, slicer: { name: 'Amounts' } }).ok).toBe(true)
    const out = await writeKept(original, s.getSnapshot())
    const parts = await partsOf(out.data)
    expect(parts.get('xl/slicerCaches/slicerCache1.xml')).toBe((await archive.text('xl/slicerCaches/slicerCache1.xml'))!.replace('tableId="1"', 'tableId="777"'))
    const ext = /<ext uri="\{46BE6895[^}]*\}"[^>]*>[\s\S]*?<\/ext>/i.exec(parts.get('xl/workbook.xml')!)![0]
    expect(ext.match(/<x14:slicerCache /g)).toHaveLength(2)
    s.close()
  })
  it('reads the four lifecycle books Mac Excel saved, including new parts and an empty drawing removed', async () => {
    const cases = JSON.parse(readFileSync(new URL('./excel/keep/slicer-lifecycle.json', import.meta.url), 'utf8')).cases
    for (const [id, want] of Object.entries(cases)) {
      const b = (await readXlsx(readFileSync(new URL(`./excel/keep/slicer-lifecycle/${id}.xlsx`, import.meta.url)))).book
      expect(
        {
          caches: [...(b.slicerCaches ?? []), ...(b.timelineCaches ?? [])].map((c) => c.name).sort(),
          objects: b.sheets.flatMap((s) => [...(s.slicers ?? []), ...(s.timelines ?? [])].map((x) => `${s.name}!${x.name}`)).sort(),
        },
        id,
      ).toEqual(want)
    }
  })
})
