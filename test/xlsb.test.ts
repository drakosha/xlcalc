import { describe, expect, it } from 'vitest'
import { readXlsb } from '../src/xlsb.js'
import { records } from '../src/xlsb/records.js'
import { Xml } from '../src/xlsb/xml.js'
import { zip } from '../src/xlsx/zip/write.js'
import { rels } from './xlsx-fixtures.js'

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const workbook = 'xl/workbook.bin'
const sheet = 'xl/worksheets/sheet1.bin'
const bytes = (value: string) => new TextEncoder().encode(value)
const relationship = (type: string, target: string) => bytes(rels(`<Relationship Id="rId1" Type="${REL}/${type}" Target="${target}"/>`))

const fixture = (data: Uint8Array) => zip([
  { name: '[Content_Types].xml', data: bytes(`<Types><Override PartName="/${sheet}" ContentType="application/vnd.ms-excel.worksheet"/></Types>`) },
  { name: '_rels/.rels', data: relationship('officeDocument', workbook) },
  { name: 'xl/_rels/workbook.bin.rels', data: relationship('worksheet', 'worksheets/sheet1.bin') },
  {
    name: workbook,
    data: new Uint8Array([
      0x83, 1, 0, // BrtBeginBook
      0x9c, 1, 26, // BrtBundleSh, with rId1 and S as XLWideStrings
      0, 0, 0, 0, 1, 0, 0, 0,
      4, 0, 0, 0, 114, 0, 73, 0, 100, 0, 49, 0,
      1, 0, 0, 0, 83, 0,
      0x84, 1, 0, // BrtEndBook
    ]),
  },
  { name: sheet, data },
])

describe('XLSB archive', () => {
  it('resolves the workbook and worksheets from their original package paths', async () => {
    const data = await fixture(Uint8Array.of(0x81, 1, 0, 0x82, 1, 0))
    const read = await readXlsb(data)
    expect(read.book.sheets).toEqual([{ name: 'S', cells: {} }])
  })

  it('refuses incomplete headers and record sizes past the part with a precise location', async () => {
    const header = await fixture(Uint8Array.of(0x81))
    await expect(readXlsb(header)).rejects.toThrow(/sheet1\.bin at offset 1: incomplete record header/)
    const size = await fixture(Uint8Array.of(0x81, 1, 0xff, 0xff, 0xff, 0xff))
    await expect(readXlsb(size)).rejects.toThrow(/sheet1\.bin at offset 0: record 129 declares 268435455 bytes/)
  })

  it('refuses a string field beyond the end of its record', () => {
    const record = records(Uint8Array.of(156, 1, 4, 8, 0, 0, 0), workbook).next().value
    expect(() => record!.payload.wide()).toThrow(/workbook\.bin at offset 7: record 156: expected 16 bytes/)
  })

  it('charges expanded UTF-8 XML before joining it', () => {
    const entry = new Xml(sheet, { entryBytes: 5 }, { spent: 0 })
    expect(() => entry.add('é🙂')).toThrow(/sheet1\.bin.*limits\.entryBytes/)
    const total = new Xml(sheet, { entryBytes: 6, totalBytes: 8 }, { spent: 3 })
    // Six expanded bytes plus the three bytes already read from the ZIP.
    expect(() => total.add('é🙂')).toThrow(/sheet1\.bin.*limits\.totalBytes/)
  })
})
