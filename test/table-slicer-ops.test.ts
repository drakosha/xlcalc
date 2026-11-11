/**
 * A table's slicer made by `addSlicer` over its column, and gone with its table or its column as Excel saves the book
 * after the same edit (`test/excel/win/table-slicer/answers.json`, from the book `plain` Excel made).
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { createSession, editStructure, InputError, type BookInput, type StructureOp } from '../src/index.js'
import { readXlsx } from '../src/xlsx.js'

const book = (): BookInput => ({
  sheets: [
    {
      name: 'Data',
      cells: { A1: { value: 'Region' }, B1: { value: 'Product' }, C1: { value: 'Amount' }, A2: { value: 'North' }, B2: { value: 'Apple' }, C2: { value: 5 } },
      tables: [{ name: 'Sales', range: 'A1:C2', columns: [{ name: 'Region' }, { name: 'Product' }, { name: 'Amount' }] }],
    },
  ],
})

const done = (b: BookInput, op: StructureOp) => {
  const r = editStructure(b, op)
  if (!r.ok) throw new Error(r.message)
  return r.book
}

describe("addSlicer over a table's column", () => {
  it("makes the cache Excel makes, named by the column, and the slicer beside the table", () => {
    const b = done(book(), { op: 'addSlicer', sheet: 'Data', slicer: {}, cache: { table: 'sales', column: 'product' } })
    expect(b.slicerCaches).toEqual([{ name: 'Slicer_Product', field: 'Product', pivots: [], table: { name: 'Sales', column: 'Product' } }])
    const [s] = b.sheets[0].slicers!
    expect(s).toMatchObject({ name: 'Product', cache: 'Slicer_Product', caption: 'Product' })
    expect(s.place?.from).toBe('E1')
    const again = done(b, { op: 'addSlicer', sheet: 'Data', slicer: {}, cache: { table: 'Sales', column: 'Product' } })
    expect(again.slicerCaches?.map((c) => c.name)).toEqual(['Slicer_Product', 'Slicer_Product1'])
    expect(again.sheets[0].slicers?.map((x) => x.name)).toEqual(['Product', 'Product 1'])
  })

  it('a table or a column that is not there is an input error', () => {
    expect(() => editStructure(book(), { op: 'addSlicer', sheet: 'Data', slicer: {}, cache: { table: 'Nope', column: 'Product' } })).toThrow(InputError)
    expect(() => editStructure(book(), { op: 'addSlicer', sheet: 'Data', slicer: {}, cache: { table: 'Sales', column: 'Price' } })).toThrow(/has no column "Price"/)
    expect(() => editStructure(book(), { op: 'addTimeline', sheet: 'Data', timeline: {}, cache: { table: 'Sales', column: 'Product' } as never })).toThrow(InputError)
  })
})

describe("a table's slicer goes with its table or its column, as in Excel", () => {
  it.each(['#N/A', '=#n/a', '42'])('removes only a lost cache-owned global #N/A name, given %s', (ref) => {
    let b = done(book(), { op: 'addSlicer', sheet: 'Data', slicer: {}, cache: { table: 'Sales', column: 'Product' } })
    b = done(b, { op: 'addSlicer', sheet: 'Data', slicer: {}, cache: { table: 'Sales', column: 'Amount' } })
    const names = [
      { name: 'Slicer_Product', ref: '#N/A' },
      { name: 'sLiCeR_AmOuNt', ref },
      { name: 'Slicer_Amount', ref: '#N/A', scope: 'Data' },
      { name: 'Slicer_Unrelated', ref: '#N/A' },
      { name: 'UserName', ref: '42' },
    ]
    const s = createSession({ ...b, names, sheets: [...b.sheets, { name: 'Checks', cells: { A1: { formula: 'Slicer_Amount' }, B1: { formula: 'Slicer_Product' } } }] })
    expect(s.editStructure({ op: 'deleteColumns', sheet: 'Data', column: 'C' }).ok).toBe(true)
    const snap = s.getSnapshot()
    expect(snap.book.slicerCaches?.map((c) => c.name)).toEqual(['Slicer_Product'])
    expect(snap.book.sheets[0].slicers?.map((x) => x.cache)).toEqual(['Slicer_Product'])
    expect(snap.book.names).toEqual(ref === '42' ? names : names.filter((n) => n.name !== 'sLiCeR_AmOuNt'))
    expect(snap.result.sheets[1].cells.A1.value).toEqual(ref === '42' ? 42 : { error: 'NAME' })
    expect(snap.result.sheets[1].cells.B1.value).toEqual({ error: 'NA' })
    s.close()
  })

  const answers = new Map((JSON.parse(readFileSync(new URL('./excel/win/table-slicer/answers.json', import.meta.url), 'utf8')).cases as { id: string; caches: { name: string }[] }[]).map((c) => [c.id, c.caches.map((x) => x.name)]))
  const plain = async () => (await readXlsx(readFileSync(new URL('./excel/win/table-slicer/plain.xlsx', import.meta.url)))).book
  const column = async (name: string) => {
    const b = await plain()
    const t = b.sheets[0].tables![0]
    // the table starts in a single-letter column
    return String.fromCharCode(t.range.charCodeAt(0) + t.columns.findIndex((c) => c.name === name))
  }
  it.each([
    ['column-deleted', async (): Promise<StructureOp> => ({ op: 'deleteColumns', sheet: 'Data', column: await column('Product') })],
    ['other-column-deleted', async (): Promise<StructureOp> => ({ op: 'deleteColumns', sheet: 'Data', column: await column('Region') })],
    ['table-deleted', async (): Promise<StructureOp> => ({ op: 'deleteTable', table: 'Sales' })],
    ['table-unlisted', async (): Promise<StructureOp> => ({ op: 'convertTable', table: 'Sales' })],
  ] as const)('%s', async (id, op) => {
    const b = done(await plain(), await op())
    expect((b.slicerCaches ?? []).map((c) => c.name)).toEqual(answers.get(id))
    expect(b.sheets.flatMap((s) => s.slicers ?? []).map((x) => x.cache)).toEqual(answers.get(id))
  })
})
