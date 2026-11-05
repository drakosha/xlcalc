/**
 * The binary workbooks the reader does not read: an `.xlsb` is a zip like an xlsx whose parts are binary, and read as
 * text it would be a workbook with no sheets, without an error. An `.xls` and a password-protected xlsx are OLE files.
 */
import { describe, expect, it } from 'vitest'
import { InputError } from '../src/input-error.js'
import { readXlsx } from '../src/xlsx.js'
import { openKept } from '../src/xlsx-keep.js'
import { pack } from './xlsx-fixtures.js'

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const xlsb = (workbook: string, types = '') =>
  pack({
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${types}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="${workbook}"/></Relationships>`,
    [workbook]: '\u0083\u0001\u0000',
  })
const OLE = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, ...new Array(504).fill(0)])

describe('a binary workbook', () => {
  it('an .xlsb is refused by name, not read as a workbook with no sheets', async () => {
    const file = await xlsb('xl/workbook.bin')
    await expect(readXlsx(file)).rejects.toThrow(InputError)
    await expect(readXlsx(file)).rejects.toThrow(/xlsb/)
    await expect(openKept(file)).rejects.toThrow(/xlsb/)
  })

  it('so is a part of another name of the binary content type', async () => {
    const type = '<Override PartName="/xl/book.dat" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>'
    await expect(readXlsx(await xlsb('xl/book.dat', type))).rejects.toThrow(/xlsb/)
  })

  it('an .xls or a password-protected file says what it may be', async () => {
    await expect(readXlsx(OLE)).rejects.toThrow(InputError)
    await expect(readXlsx(OLE)).rejects.toThrow(/\.xls.*password/)
  })
})
