/**
 * A table's slicer against the books Excel for Windows made and saved (`bench/windows/table-slicer.ps1`): its cache names the
 * table and the column by their ids in the file, which `readXlsx` turns into their names and `writeKept` keeps.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { createSession } from '../src/index.js'
import { readXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { openArchive } from '../src/xlsx/zip/read.js'

const file = (id: string) => new Uint8Array(readFileSync(new URL(`./excel/win/table-slicer/${id}.xlsx`, import.meta.url)))
const CACHE = 'xl/slicerCaches/slicerCache1.xml'
const text = async (data: Uint8Array, part: string) => (await openArchive(data).text(part)) ?? ''
const cacheOf = async (id: string) => (await readXlsx(file(id))).book.slicerCaches![0]
const tableIds = (xml: string) => [...xml.matchAll(/<tableColumn id="(\d+)"[^>]*? name="([^"]*)"/g)].map((m) => [Number(m[1]), m[2]])

/** The file with the book's change through a session, written over it. */
async function kept(id: string, change: (s: ReturnType<typeof createSession>) => { ok: boolean } | void) {
  const original = file(id)
  const { book } = await openKept(original)
  const s = createSession(book)
  const r = change(s)
  if (r) expect(r.ok).toBe(true)
  const { data } = await writeKept(original, s.getSnapshot())
  s.close()
  return data
}

describe("reading a table's slicer", () => {
  it('names the column, whose id in the file is not its place', async () => {
    expect(await cacheOf('plain')).toEqual({ name: 'Slicer_Product', field: 'Product', pivots: [], table: { name: 'Sales', column: 'Product' } })
  })

  it.each([
    ['sort-desc', { sort: 'descending' }],
    ['no-custom', { customListSort: false }],
    ['cross-none', { noData: { mark: false } }],
    ['cross-last', { noData: { last: false } }],
  ])('%s: the order and display the table cache holds', async (id, fields) => {
    expect(await cacheOf(id)).toMatchObject(fields)
  })

  it('a column renamed in Excel: the field and the column by its new name', async () => {
    expect(await cacheOf('column-renamed')).toMatchObject({ field: 'Goods', table: { name: 'Sales', column: 'Goods' } })
  })
})

describe("a table's slicer in a session", () => {
  const session = async () => createSession((await readXlsx(file('plain'))).book)

  it('follows the table renamed', async () => {
    const s = await session()
    expect(s.editStructure({ op: 'renameTable', table: 'Sales', name: 'Goods' }).ok).toBe(true)
    expect(s.getSnapshot().book.slicerCaches![0].table).toEqual({ name: 'Goods', column: 'Product' })
    s.close()
  })

  it('follows its column renamed by its header, as Excel does', async () => {
    const s = await session()
    s.setCells([{ sheet: 'Data', cell: 'B1', input: { value: 'Goods' } }])
    expect(s.getSnapshot().book.slicerCaches![0]).toMatchObject({ name: 'Slicer_Product', field: 'Goods', table: { name: 'Sales', column: 'Goods' } })
    s.close()
  })
})

describe("writeKept and a table's slicer", () => {
  it('its order changed: still the cache of the table and column, in the form Excel writes', async () => {
    const xml = await text(await kept('plain', (s) => s.editStructure({ op: 'editSlicer', slicer: 'Product', cache: { sort: 'descending', customListSort: false, noData: { mark: false } } })), CACHE)
    expect(xml).toContain('<x15:tableSlicerCache tableId="1" column="3" sortOrder="descending" customListSort="0" crossFilter="none"/>')
    expect(xml).toContain('uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}"')
    expect(xml).not.toContain('tabular')
    expect(xml).not.toContain('pivotTables')
  })

  it('the table renamed: the ids of its columns stay, and the part of the cache', async () => {
    const data = await kept('plain', (s) => s.editStructure({ op: 'renameTable', table: 'Sales', name: 'Goods' }))
    expect(tableIds(await text(data, 'xl/tables/table1.xml'))).toEqual([[2, 'Region'], [3, 'Product'], [4, 'Amount']])
    expect(await text(data, CACHE)).toBe(await text(file('plain'), CACHE))
    expect((await readXlsx(data)).book.slicerCaches![0].table).toEqual({ name: 'Goods', column: 'Product' })
  })

  it('a column added to the table takes an id of its own', async () => {
    const data = await kept('plain', (s) => s.editStructure({ op: 'insertColumns', sheet: 'Data', column: 'B', count: 1 }))
    const ids = tableIds(await text(data, 'xl/tables/table1.xml'))
    expect(ids.map((x) => x[0])).toEqual([2, 5, 3, 4])
    expect((await readXlsx(data)).book.slicerCaches![0].table).toEqual({ name: 'Sales', column: 'Product' })
  })

  it('its column renamed: the cache takes the new source name, the id stays', async () => {
    const data = await kept('plain', (s) => s.setCells([{ sheet: 'Data', cell: 'B1', input: { value: 'Goods' } }]) as never)
    const xml = await text(data, CACHE)
    expect(xml).toContain('sourceName="Goods"')
    expect(xml).toContain('column="3"')
    expect(tableIds(await text(data, 'xl/tables/table1.xml'))[1]).toEqual([3, 'Goods'])
  })

  it('the slicer renamed: its frame in the form Excel gives a table slicer', async () => {
    const xml = await text(await kept('plain', (s) => s.editStructure({ op: 'editSlicer', slicer: 'Product', set: { name: 'Goods' } })), 'xl/drawings/drawing1.xml')
    expect(xml).toContain('Requires="sle15"')
    expect(xml).toContain('xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer"')
    expect(xml).toContain('<sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="Goods"/>')
  })

  it('nothing changed: the file as it was', async () => {
    const original = file('chosen')
    const { book, cached } = await openKept(original)
    const { data } = await writeKept(original, { book, result: cached })
    expect(data).toEqual(original)
  })
})
