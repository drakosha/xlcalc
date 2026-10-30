/**
 * `writeKept` over parts that start with a UTF-8 byte order mark, as some writers leave them before `<?xml`: without
 * edits every part stays byte for byte, a part written again keeps its mark, and one whose text comes out the same is
 * the file's own.
 */
import { describe, expect, it } from 'vitest'
import { writeXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { zip } from '../src/xlsx/zip/write.js'

const BOM = [0xef, 0xbb, 0xbf]
const MARKED = ['xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/sharedStrings.xml']

async function marked(): Promise<Uint8Array> {
  const { data } = await writeXlsx({ book: { sheets: [{ name: 'S', cells: { A1: { value: 'x' }, B1: { formula: 'A1&"y"' } } }], names: [{ name: 'N', ref: 'S!$A$1' }] } })
  const a = openArchive(data)
  return zip(await Promise.all(a.names.map(async (name) => ({ name, data: MARKED.includes(name) ? new Uint8Array([...BOM, ...(await a.bytes(name))!]) : (await a.bytes(name))! }))))
}

const parts = async (data: Uint8Array) => {
  const a = openArchive(data)
  return new Map(await Promise.all(a.names.map(async (n) => [n, (await a.bytes(n))!] as const)))
}

describe('writeKept over parts with a byte order mark', () => {
  it('leaves every part as it was without edits', async () => {
    const data = await marked()
    const out = await writeKept(data, { book: (await openKept(data)).book })
    expect(await parts(out.data)).toEqual(await parts(data))
  })

  it('keeps the mark on the parts it writes again, and the file\'s own part where the text is the same', async () => {
    const data = await marked()
    const { book } = await openKept(data)
    const out = await writeKept(data, { book: { ...book, sheets: book.sheets.map((s) => ({ ...s, cells: { ...s.cells, A2: { value: 'new' } } })) } })
    const [before, after] = [await parts(data), await parts(out.data)]
    for (const name of MARKED) expect([name, [...after.get(name)!.slice(0, 4)]]).toEqual([name, [...BOM, 0x3c]])
    expect(after.get('xl/worksheets/sheet1.xml')).not.toEqual(before.get('xl/worksheets/sheet1.xml'))
    // the workbook's text comes out the same (it already asks for a full calculation on loading)
    expect(after.get('xl/workbook.xml')).toEqual(before.get('xl/workbook.xml'))
  })
})
