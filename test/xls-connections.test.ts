import { readFile } from 'node:fs/promises'
import { describe, expect, it } from 'vitest'
import { openCompound } from '../src/xls/compound.js'
import { records } from '../src/xls/records.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { attributes, elements } from '../src/xlsx/xml/scan.js'
import { readText } from '../src/xlsx/xml/text.js'
import { libUrl } from './fixtures/target.js'

const { readXlsWithConnections } = await import(libUrl('xls-connections.js')) as typeof import('../src/xls-connections.js')
const fixture = (name: string, ext: string) => readFile(new URL(`./excel/xls-read/${name}.${ext}`, import.meta.url))

async function excelConnections(name: string) {
  const xlsx = await fixture(name, 'xlsx')
  const xml = await openArchive(xlsx).text('xl/connections.xml')
  return elements(xml ?? '', 'connection').map((node) => {
    const attrs = attributes(node.attrs)
    const database = attributes(elements(node.body ?? '', 'dbPr')[0]?.attrs ?? '')
    const text = attributes(elements(node.body ?? '', 'textPr')[0]?.attrs ?? '')
    return { id: Number(attrs.id), type: Number(attrs.type), name: readText(attrs.name),
      database: { connectionString: readText(database.connection ?? ''), command: readText(database.command ?? '') },
      text: { sourceFile: readText(text.sourceFile ?? ''), comma: text.comma === '1' } }
  })
}

describe('readXlsWithConnections', () => {
  it.each(['pivot-external-text', 'pivot-external-many', 'source-ole-pivot'])('reads Excel’s connection definitions in %s', async (name) => {
    const read = await fixture(name, 'xls').then(readXlsWithConnections)
    const expected = await excelConnections(name)
    expect(expected.length).toBeGreaterThan(0)
    expect(read.connectionIssues).toEqual([])
    expect(read.connections).toHaveLength(expected.length)
    expect(read.connections).toMatchObject(expected.map(({ text: _text, ...connection }) => connection))
    for (const cache of read.book.pivotCaches ?? []) {
      const { source } = cache
      if (source.type !== 'external') continue
      expect(read.connections.find((connection) => connection.id === source.connection)?.pivotCache).toBe(cache.id)
    }
  })

  it('reads Excel’s saved text query settings without changing its cells', async () => {
    const read = await fixture('source-query', 'xls').then(readXlsWithConnections)
    const [expected] = await excelConnections('source-query')
    expect(read.connectionIssues).toEqual([])
    expect(read.connections).toMatchObject([{ id: expected.id, type: 6, name: expected.name, saveData: true, text: expected.text }])
    expect(read.book.sheets[0].cells.B5).toEqual({ value: 20 })
  })

  it('reads the OLAP hierarchies, measures, dimensions and field links stored by Excel', async () => {
    const [xls, xlsx] = await Promise.all([fixture('olap-offline-cube', 'xls'), fixture('olap-offline-cube', 'xlsx')])
    const read = await readXlsWithConnections(xls)
    const archive = openArchive(xlsx)
    const cache = await archive.text('xl/pivotCache/pivotCacheDefinition1.xml') ?? ''
    const expected = elements(cache, 'cacheHierarchy').map((hierarchy) => {
      const attrs = attributes(hierarchy.attrs)
      return { uniqueName: readText(attrs.uniqueName), caption: readText(attrs.caption),
        kind: attrs.measure === '1' ? 'measure' : 'hierarchy',
        ...(attrs.dimensionUniqueName ? { dimension: readText(attrs.dimensionUniqueName) } : {}) }
    })
    expect(expected).toHaveLength(13)
    expect(read.connectionIssues).toEqual([])
    expect(read.connections).toMatchObject([
      { id: 1, name: 'Local cube', type: 5, pivotCache: 1, olap: true, database: { commandType: 1, command: 'Cube Test EN' } },
    ])
    expect(read.olapPivots).toHaveLength(1)
    const [pivot] = read.olapPivots
    expect(pivot).toMatchObject({ sheet: 'Report', name: 'OfflineCube', range: 'A3:B6', cache: 1, fields: [{ hierarchy: 0, level: 1 }, { hierarchy: 8, level: 0 }] })
    expect(pivot.hierarchies.slice(0, expected.length)).toMatchObject(expected)
    expect(pivot.hierarchies[0]).toMatchObject({ axis: 'row', defaultMember: '[Products].[Product].[All]' })
    expect(pivot.hierarchies[8]).toMatchObject({ axis: 'values' })
    expect(pivot.hierarchies.at(-1)).toMatchObject({ kind: 'kpi', uniqueName: 'Average' })
  })

  it('uses the worksheet OLAP flag even when the connection catalogue is absent', async () => {
    const bytes = await fixture('olap-offline-cube', 'xls')
    const workbook = openCompound(bytes, {}).workbook
    const connection = Array.from(records(workbook, 0, workbook.length)).find((record) => record.type === 0x0876)!
    const marker = workbook.subarray(connection.offset, connection.offset + 16)
    const position = bytes.indexOf(marker)
    expect(position).toBeGreaterThan(0)
    expect(bytes.lastIndexOf(marker)).toBe(position)
    bytes.writeUInt16LE(0, position)
    const read = await readXlsWithConnections(bytes)
    expect(read.connections).toEqual([])
    expect(read.olapPivots[0]?.hierarchies[8]).toMatchObject({ uniqueName: '[Measures].[Sales]' })
    expect(read.book.pivotCaches?.[0].fields[1]).toEqual({ name: '[Measures].[Sales]' })
    expect(read.issues).toContainEqual(expect.objectContaining({ code: 'pivotTable', message: expect.stringContaining('OLAP') }))
  })

  it('reports a malformed definition while leaving the worksheet available', async () => {
    const file = await fixture('source-query', 'xls')
    const workbook = Uint8Array.from(openCompound(file, {}).workbook)
    const connection = Array.from(records(workbook, 0, workbook.length)).find((record) => record.type === 0x0876)!
    // The first segmented string now declares more characters than remain in DConn.
    new DataView(workbook.buffer).setUint32(connection.offset + 4 + 29, 0xffffffff, true)
    const read = await readXlsWithConnections(workbook)
    expect(read.connectionIssues).toEqual([{ id: 1, reason: 'invalid' }])
    expect(read.book.sheets[0].cells.B5).toEqual({ value: 20 })
  })
})
