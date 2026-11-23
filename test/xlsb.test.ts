import { readFile } from 'node:fs/promises'
import { describe, expect, it } from 'vitest'
import { readXlsb } from '../src/xlsb.js'
import { readXlsx } from '../src/xlsx.js'
import { records } from '../src/xlsb/records.js'
import { Xml } from '../src/xlsb/xml.js'
import { zip } from '../src/xlsx/zip/write.js'
import { rels } from './xlsx-fixtures.js'

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const workbook = 'xl/workbook.bin'
const sheet = 'xl/worksheets/sheet1.bin'
const link = 'xl/externalLinks/externalLink1.bin'
const bytes = (value: string) => new TextEncoder().encode(value)
const related = (type: string, target: string, id = 'rId1') => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`

const fixture = (data: Uint8Array, external?: Uint8Array) => zip([
  {
    name: '[Content_Types].xml',
    data: bytes(`<Types><Override PartName="/${sheet}" ContentType="application/vnd.ms-excel.worksheet"/>`
      + (external ? `<Override PartName="/${link}" ContentType="application/vnd.ms-excel.externalLink"/>` : '') + '</Types>'),
  },
  { name: '_rels/.rels', data: bytes(rels(related('officeDocument', workbook))) },
  {
    name: 'xl/_rels/workbook.bin.rels',
    data: bytes(rels(related('worksheet', 'worksheets/sheet1.bin') + (external ? related('externalLink', 'externalLinks/externalLink1.bin', 'rId2') : ''))),
  },
  {
    name: workbook,
    data: new Uint8Array([
      0x83, 1, 0, // BrtBeginBook
      0x9c, 1, 26, // BrtBundleSh, with rId1 and S as XLWideStrings
      0, 0, 0, 0, 1, 0, 0, 0,
      4, 0, 0, 0, 114, 0, 73, 0, 100, 0, 49, 0,
      1, 0, 0, 0, 83, 0,
      // BrtSupBookSrc with rId2
      ...(external ? [0xe3, 2, 12, 4, 0, 0, 0, 114, 0, 73, 0, 100, 0, 50, 0] : []),
      0x84, 1, 0, // BrtEndBook
    ]),
  },
  { name: sheet, data },
  ...(external ? [{ name: link, data: external }] : []),
])

describe('XLSB archive', () => {
  it.each(['probe', 'formulas', 'pivot-calculated', 'slicers-and-timeline', 'icon-filter', 'array-exact', 'array-convert'])('reads both of Excel’s %s saves identically', async (name) => {
    const path = `./excel/win/xlsb-read/${name}`
    const [binary, xml] = await Promise.all(['xlsb', 'xlsx'].map((ext) => readFile(new URL(`${path}.${ext}`, import.meta.url))))
    await expect(readXlsb(binary)).resolves.toStrictEqual(await readXlsx(xml))
  })

  it('resolves the workbook and worksheets from their original package paths', async () => {
    const data = await fixture(Uint8Array.of(0x81, 1, 0, 0x82, 1, 0))
    const read = await readXlsb(data)
    expect(read.book.sheets).toEqual([{ name: 'S', cells: {} }])
  })

  it.each([
    [0x3a, [1, 0, 3, 0], "'S'!\u0024D\u00242", false],
    [0x3c, [0x9f, 0x86, 0, 0], '#REF!', true],
    [0x3d, [0x9f, 0x86, 0, 0, 2, 0, 4, 0], '#REF!', true],
  ] as const)('reads external definition token %i and reports a stored invalid reference', async (token, coordinates, ref, invalid) => {
    const definition = [token, 0, 0, 0, 0, ...coordinates]
    const data = await fixture(Uint8Array.of(0x81, 1, 0, 0x82, 1, 0), Uint8Array.of(
      0xe8, 2, 10, 0, 0, 255, 255, 255, 255, 255, 255, 255, 255, // BrtSupBook, no path
      0xe7, 2, 10, 1, 0, 0, 0, 1, 0, 0, 0, 83, 0, // BrtSupTabs, sheet S
      0xc1, 4, 6, 1, 0, 0, 0, 78, 0, // BrtSupNameStart, name N
      0xc9, 4, 4 + definition.length, definition.length, 0, 0, 0, ...definition, // BrtSupNameFmla
      0xcb, 4, 0, // BrtSupNameEnd
    ))
    const read = await readXlsb(data)
    expect(read.book.externalBooks?.[0].names).toEqual([{ name: 'N', ref }])
    expect(read.issues.filter((issue) => issue.code === 'externalLink')).toEqual(invalid
      ? [{ code: 'externalLink', name: 'N', message: expect.stringContaining('stores an invalid reference') }]
      : [])
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
