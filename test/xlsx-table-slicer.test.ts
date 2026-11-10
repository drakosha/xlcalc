/**
 * A table's slicer against the books Excel for Windows made and saved (`bench/windows/table-slicer.ps1`): its cache names the
 * table and the column by their ids in the file, which `readXlsx` turns into their names and `writeKept` keeps.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredLinks } from '../bench/coverage/links.js'
import { buildLinksCases, SAVED, type Held } from '../bench/coverage/links/table-slicer-kept.js'
import { createSession } from '../src/index.js'
import { readXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { zip } from '../src/xlsx/zip/write.js'

const file = (id: string) => new Uint8Array(readFileSync(new URL(`./excel/win/table-slicer/${id}.xlsx`, import.meta.url)))
const CACHE = 'xl/slicerCaches/slicerCache1.xml'
const text = async (data: Uint8Array, part: string) => (await openArchive(data).text(part)) ?? ''
const cacheOf = async (id: string) => (await readXlsx(file(id))).book.slicerCaches![0]
const tableIds = (xml: string) => [...xml.matchAll(/<tableColumn id="(\d+)"[^>]*? name="([^"]*)"/g)].map((m) => [Number(m[1]), m[2]])

/** The file with one part's text changed. */
async function rezipped(data: Uint8Array, part: string, change: (xml: string) => string) {
  const a = openArchive(data)
  return zip(await Promise.all(a.names.map(async (name) => ({ name, data: await a.bytes(name) ?? new Uint8Array() }))).then((all) => all.map((e) => (e.name === part ? { name: part, data: new TextEncoder().encode(change(new TextDecoder().decode(e.data))) } : e))))
}

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
    const data = await kept('plain', (s) => void s.setCells([{ sheet: 'Data', cell: 'B1', input: { value: 'Goods' } }]))
    const xml = await text(data, CACHE)
    expect(xml).toContain('sourceName="Goods"')
    expect(xml).toContain('column="3"')
    expect(tableIds(await text(data, 'xl/tables/table1.xml'))[1]).toEqual([3, 'Goods'])
  })

  it('two columns renamed, one to the name the other had: no id is taken twice', async () => {
    const data = await kept('plain', (s) => void s.setCells([{ sheet: 'Data', cell: 'A1', input: { value: 'Area' } }, { sheet: 'Data', cell: 'B1', input: { value: 'Region' } }]))
    const ids = tableIds(await text(data, 'xl/tables/table1.xml'))
    expect(ids).toEqual([[5, 'Area'], [2, 'Region'], [4, 'Amount']])
    expect(new Set(ids.map((x) => x[0])).size).toBe(ids.length)
  })

  it('a column deleted and another added inside: no id is taken twice', async () => {
    const data = await kept('plain', (s) => {
      s.editStructure({ op: 'deleteColumns', sheet: 'Data', column: 'A', count: 1 })
      s.editStructure({ op: 'insertColumns', sheet: 'Data', column: 'B', count: 1 })
    })
    const ids = tableIds(await text(data, 'xl/tables/table1.xml')).map((x) => x[0])
    expect(new Set(ids).size).toBe(ids.length)
  })

  it('the slicer renamed: its frame in the form Excel gives a table slicer', async () => {
    const xml = await text(await kept('plain', (s) => s.editStructure({ op: 'editSlicer', slicer: 'Product', set: { name: 'Goods' } })), 'xl/drawings/drawing1.xml')
    expect(xml).toContain('Requires="sle15"')
    expect(xml).toContain('xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer"')
    expect(xml).toContain('<sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="Goods"/>')
  })

  it('a cache of an OLAP source changed is refused, as the model does not hold its part', async () => {
    const original = await rezipped(file('plain'), CACHE, (x) => x.replace(/<extLst>.*<\/extLst>/s, '<data><olap pivotCacheId="1"><levels count="0"/><selections count="0"/></olap></data>'))
    const { book } = await openKept(original)
    const s = createSession(book)
    s.editStructure({ op: 'editSlicer', slicer: 'Product', cache: { sort: 'descending' } })
    await expect(writeKept(original, s.getSnapshot())).rejects.toMatchObject({ code: 'slicers' })
    s.close()
  })

  it('nothing changed: the file as it was', async () => {
    const original = file('chosen')
    const { book, cached } = await openKept(original)
    const { data } = await writeKept(original, { book, result: cached })
    expect(data).toEqual(original)
  })
})

/** What Excel's saved file holds, as `SAVED` predicts it. */
async function heldBy(data: Uint8Array): Promise<Held> {
  const table = await text(data, 'xl/tables/table1.xml')
  const sheet = await text(data, 'xl/worksheets/sheet1.xml')
  const cache = await text(data, CACHE)
  const filters = [...table.matchAll(/<filter val="([^"]*)"/g)].map((m) => m[1])
  return {
    table: /displayName="([^"]*)"/.exec(table)![1],
    columns: tableIds(table) as [number, string][],
    source: /sourceName="([^"]*)"/.exec(cache)![1],
    cache: /<x15:tableSlicerCache ([^/]*)\/>/.exec(cache)![1],
    slicers: [...(await text(data, 'xl/slicers/slicer1.xml')).matchAll(/<slicer name="([^"]*)"/g)].map((m) => m[1]),
    hidden: [...sheet.matchAll(/<row [^>]*hidden="1"/g)].length,
    ...(filters.length ? { filters } : {}),
  }
}

describe.each([['Windows', 'win/']])('Excel for %s holds what writeKept wrote', (_, dir) => {
  const url = new URL(`./excel/${dir}links/table-slicer-kept.json`, import.meta.url)
  const set = existsSync(url) ? (JSON.parse(readFileSync(url, 'utf8')) as StoredLinks) : null
  it.skipIf(!set)('the cases the prediction missed', async () => {
    const missed: string[] = []
    for (const c of buildLinksCases().cases) {
      if (!!c.refused !== !!set!.rejected?.[c.id]) missed.push(c.id)
      const want = SAVED[c.id]
      if (!want || !set!.answers[c.id]) continue
      const got = await heldBy(new Uint8Array(readFileSync(new URL(`./excel/${dir}links/table-slicer-kept/${c.id}.xlsx`, import.meta.url))))
      if (JSON.stringify(got) !== JSON.stringify(want)) missed.push(c.id)
    }
    expect(missed).toEqual(MISSED[dir])
  })
})

/** The cases the committed prediction missed, by Excel: its folder of answers. */
const MISSED: Record<string, string[]> = { 'win/': [] }
