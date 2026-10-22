/**
 * Reproductions from the review of 21.10 on reading and writing xlsx files that are damaged or written by hand: a zip archive whose
 * entries are cut short, named twice or do not match their CRC-32 is refused with the reader's error; XML comments are
 * no content, and a character reference past U+10FFFF is refused as XML refuses it; a row number, a style index, a
 * shared string's index or a value the reader cannot take is reported, and leaves a valid workbook. Writing reports
 * what it changes: a formula left out, #ERROR! as #NULL!, text cut without splitting a character. Reading a file does
 * not calculate; what a write takes from a result is checked in both execution modes.
 */
import { deflateRawSync } from 'node:zlib'
import { describe, expect, it } from 'vitest'
import { openArchive } from '../src/xlsx/zip/read.js'
import { rawZip, type RawEntry } from './fixtures/raw-zip.js'
import type { BookInput } from '../src/contract.js'
import { compute } from '../src/index.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { pack, parts, read, sheet } from './xlsx-fixtures.js'

const enc = new TextEncoder()
const entry = (name: string, text: string): RawEntry => ({ name, data: enc.encode(text) })

/** The archive with its first central directory record's field at `at` (4 bytes) set to `value`. */
function patched(zip: Uint8Array, at: number, value: (zip: Uint8Array) => number): Uint8Array {
  const out = zip.slice()
  const v = new DataView(out.buffer)
  const dir = v.getUint32(out.length - 22 + 16, true)
  v.setUint32(dir + at, value(out), true)
  return out
}

const refusal = (zip: Uint8Array, name = 'a.xml') => openArchive(zip).bytes(name).then(() => 'read', (e: Error) => `${e.constructor.name}: ${e.message}`)

describe('xlsx: a damaged zip archive (x-xlsx-3, x-xlsx-12, x-xlsx-13)', () => {
  it("an entry whose local header or data lies past the end of the file is refused with the reader's error", async () => {
    const zip = rawZip([entry('a.xml', '<a/>'.repeat(10))])
    const offset = 42
    for (const at of [(z: Uint8Array) => z.length + 100, () => 0x7fffffff, (z: Uint8Array) => z.length - 2, (z: Uint8Array) => z.length - 29]) {
      expect(await refusal(patched(zip, offset, at))).toBe('Error: xlsx: not a readable zip archive: no local header for a.xml')
    }
    // the data cut short: the entry says it has more bytes than the file
    expect(await refusal(patched(zip, 20, () => 1000))).toBe('Error: xlsx: not a readable zip archive: entry a.xml is cut short')
  })

  it('two entries of one name, in any case, are refused', () => {
    const a = entry('xl/worksheets/sheet1.xml', '<a/>')
    for (const b of [{ ...a, data: enc.encode('<b/>') }, { ...a, name: 'XL/WORKSHEETS/SHEET1.XML' }]) {
      expect(() => openArchive(rawZip([a, b]))).toThrow(`xlsx: not a readable zip archive: entries xl/worksheets/sheet1.xml and ${b.name} have the same name`)
    }
  })

  it('an entry whose bytes do not match its CRC-32, or a stored one whose sizes differ, is refused; a sound one reads', async () => {
    const text = '<worksheet/>'
    const deflated: RawEntry = { name: 'a.xml', data: deflateRawSync(enc.encode(text)), method: 8, usize: text.length }
    const damaged = 'Error: xlsx: not a readable zip archive: entry a.xml is damaged (its CRC-32 does not match)'
    expect(await refusal(rawZip([{ ...entry('a.xml', text), crc: 0xdeadbeef }]))).toBe(damaged)
    expect(await refusal(rawZip([{ ...deflated, crc: 0xdeadbeef }]))).toBe(damaged)
    expect(await refusal(rawZip([{ ...entry('a.xml', text), usize: 5 }]))).toBe('Error: xlsx: not a readable zip archive: entry a.xml is stored in 12 bytes but says it holds 5')
    expect(await openArchive(rawZip([deflated])).text('a.xml')).toBe(text)
    expect(await openArchive(rawZip([entry('a.xml', text)])).text('a.xml')).toBe(text)
  })
})

describe('xlsx: XML comments (x-xlsx-15)', () => {
  it('a cell, a row or a sheet commented out is not read; a CDATA section keeps what looks like a comment', async () => {
    const r = await sheet('<row r="1"><c r="A1"><v>1</v></c><!-- <c r="B1"><v>2</v></c> --><c r="C1" t="str"><f>"x"</f><v><![CDATA[<!-- y -->]]></v></c></row><!-- <row r="2"><c r="A2"><v>9</v></c></row> -->')
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 1 }, C1: { formula: '"x"' } })
    expect(r.cached.sheets[0].cells.C1).toEqual({ value: '<!-- y -->' })
    const wb = await read({ sheets: [{ name: 'S', body: '<sheetData/>' }, { name: 'T', body: '<sheetData/>' }], workbook: { before: '<!-- <sheets><sheet name="X" sheetId="9" r:id="rId1"/></sheets> -->' } })
    expect(wb.book.sheets.map((s) => s.name)).toEqual(['S', 'T'])
    expect(wb.issues).toEqual([])
  })
})

describe('xlsx: a character reference past U+10FFFF (x-xlsx-2)', () => {
  it("refuses the part with the reader's error, in any text of the file; within U+10FFFF and in a CDATA section it reads", async () => {
    const shared = (t: string) => sheet('<row r="1"><c r="A1" t="s"><v>0</v></c></row>', { strings: [t] })
    for (const ref of ['&#1114112;', '&#x110000;', `&#${'9'.repeat(400)};`]) {
      await expect(shared(`a${ref}`)).rejects.toThrow(`xlsx: the part xl/sharedStrings.xml is not valid XML: the character reference ${ref} is past U+10FFFF`)
      await expect(read({ sheets: [{ name: `S${ref}`, body: '<sheetData/>' }] })).rejects.toThrow(/^xlsx: the part xl\/workbook\.xml is not valid XML/)
    }
    expect((await shared('a&#x10FFFF;&#128512;')).book.sheets[0].cells.A1).toEqual({ value: 'a\u{10FFFF}😀' })
    expect((await shared('<t><![CDATA[a&#1114112;]]></t>')).book.sheets[0].cells.A1).toEqual({ value: 'a&#1114112;' })
  })
})

/** The issues of a read as [code, cells, message]. */
const issuesOf = (r: { issues: { code: string; cells?: string[]; message: string }[] }) => r.issues.map((i) => [i.code, i.cells, i.message])

describe('xlsx: a row number that is no whole number (x-xlsx-5)', () => {
  it('places the row after the previous one with an issue; the workbook computes', async () => {
    const rows = ['1.5', 'abc', '1e3'].map((r, i) => `<row r="${r}" ht="20" customHeight="1"><c><v>${i + 1}</v></c></row>`).join('')
    const r = await sheet(`<row r="2"><c r="A2"><v>0</v></c></row>${rows}<row r=" 7 "><c><v>7</v></c></row>`)
    const s = r.book.sheets[0]
    expect(s.cells).toEqual({ A2: { value: 0 }, A3: { value: 1 }, A4: { value: 2 }, A5: { value: 3 }, A7: { value: 7 } })
    expect(s.rows).toEqual({ '3:5': { height: 20 } })
    expect(issuesOf(r)).toEqual([['part', ['3:3', '4:4', '5:5'], 'a row whose number is not a whole number is read as the row after the previous one']])
    expect(() => compute(r.book)).not.toThrow()
  })

  it('a row outside the sheet with a height or a format is reported, and the file reads', async () => {
    const r = await sheet('<row r="0" ht="20" customHeight="1"><c><v>1</v></c></row><row r="1048577" hidden="1"/><row r="1"><c r="A1"><v>2</v></c></row>')
    expect(r.book.sheets[0]).toMatchObject({ cells: { A1: { value: 2 } } })
    expect(r.book.sheets[0].rows).toBeUndefined()
    expect(issuesOf(r)).toEqual([['part', ['0:0', '0:1', '1048577:1048577'], 'a cell outside the sheet (past its last row or column) is left out']])
  })
})

describe('xlsx: a shared string the file does not have (x-xlsx-9)', () => {
  it('leaves the value out with an issue; an empty <v/> is no first string', async () => {
    const cells = ['999', '-1', 'abc', '1.5', '', '1e400', ' 1 '].map((v, i) => `<c r="${String.fromCharCode(65 + i)}1" t="s"><v>${v}</v></c>`).join('')
    const r = await sheet(`<row r="1">${cells}<c r="H1" s="1" t="s"><v>9</v></c></row>`, { strings: ['hello', 'x'], styles: '<cellXfs count="2"><xf/><xf numFmtId="10" applyNumberFormat="1"/></cellXfs>' })
    expect(r.book.sheets[0].cells).toEqual({ G1: { value: 'x' } })
    expect(r.book.sheets[0].looks).toEqual([{ range: 'H1', numberFormat: 10 }])
    expect(issuesOf(r)).toEqual([['part', ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'H1'], 'a value that names a shared string the file does not have is left out']])
  })
})

describe('xlsx: a style index past the cell formats (x-xlsx-10)', () => {
  it("is read as no style with an issue, for a cell, a row and a column; index 0 of a file without styles is the default", async () => {
    const styles = '<cellXfs count="2"><xf/><xf numFmtId="10" applyNumberFormat="1"/></cellXfs>'
    const cells = ['999', '-1', 'abc', '1e3', '1.5', '1'].map((s, i) => `<c r="${String.fromCharCode(65 + i)}1" s="${s}"><v>${i}</v></c>`).join('')
    const r = await read({ sheets: [{ name: 'S', body: `<cols><col min="8" max="8" style="7" width="9"/></cols><sheetData><row r="1" s="999" customFormat="1">${cells}</row></sheetData>` }], styles })
    expect(r.book.sheets[0].cells.F1).toEqual({ value: 5, numberFormat: 10 })
    expect(issuesOf(r)).toEqual([['part', ['H:H', '1:1', 'A1', 'B1', 'C1', 'D1', 'E1'], 'a style index that names no cell format of the file is read as no style']])
    expect((await sheet('<row r="1"><c r="A1" s="0"><v>1</v></c></row>')).issues).toEqual([])
  })
})

describe('xlsx: a value that is not a number, or a boolean that is not 0 or 1 (x-xlsx-11)', () => {
  it('is left out with an issue, not coerced', async () => {
    const values: [string, string][] = [['n', 'abc'], ['n', 'Infinity'], ['n', '1e999'], ['n', '1_000'], ['n', 'NaN'], ['n', '0x10'], ['b', 'maybe'], ['b', ''], ['n', ' 12 '], ['n', '-1.5e3'], ['n', '.5'], ['b', '1'], ['b', 'false'], ['n', '']]
    const cells = values.map(([t, v], i) => `<c r="${String.fromCharCode(65 + i)}1" t="${t}"><v>${v}</v></c>`).join('')
    const r = await sheet(`<row r="1">${cells}<c r="O1"><f>1/0</f><v>zzz</v></c></row>`)
    expect(r.book.sheets[0].cells).toEqual({ I1: { value: 12 }, J1: { value: -1500 }, K1: { value: 0.5 }, L1: { value: true }, M1: { value: false }, O1: { formula: '1/0' } })
    expect(r.cached.sheets[0].cells).toEqual({})
    expect(issuesOf(r)).toEqual([['part', ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'O1'], 'a value that is not a number, or a boolean that is not 0 or 1, is left out']])
  })
})

describe('xlsx: a name scoped to a sheet whose part is missing (fix-u331)', () => {
  it('is left out with the sheet, and reported', async () => {
    const f = parts({ sheets: [{ name: 'S', body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' }, { name: 'T', body: '<sheetData/>' }], workbook: { after: '<definedNames><definedName name="loc" localSheetId="1">\'T\'!$A$1</definedName><definedName name="own" localSheetId="0">S!$A$1</definedName><definedName name="all">S!$A$1</definedName></definedNames>' } })
    delete f['xl/worksheets/sheet2.xml']
    const r = await readXlsx(await pack(f))
    expect(r.book.sheets.map((s) => s.name)).toEqual(['S'])
    expect(r.book.names).toEqual([{ name: 'own', ref: 'S!$A$1', scope: 'S' }, { name: 'all', ref: 'S!$A$1' }])
    expect(r.issues.map((i) => [i.code, i.name ?? i.sheet, i.message])).toEqual([
      ['name', 'loc', 'a name scoped to a sheet that was left out is left out too'],
      ['part', 'T', 'a sheet whose part is missing is left out'],
    ])
  })
})

/** The `<sheetData>` of a written file's first sheet. */
async function sheetData(data: Uint8Array): Promise<string> {
  return /<sheetData>.*<\/sheetData>/.exec((await openArchive(data).text('xl/worksheets/sheet1.xml'))!)![0]
}

const clock = { now: '2026-10-22T12:00:00Z', seed: 1 }

describe.each(['closures', 'eval'] as const)('xlsx: what a write changes is reported as it happens, %s', (execution) => {
  it('a formula the file cannot hold, without a value, is reported as left out; with a value as written as it (fix-u328)', async () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: "'Nope'!A1*2", style: 0 }, A2: { formula: 'SUM(' }, A3: { formula: 'SUM(', held: 4 } } }], styles: [{ font: { bold: true } }] }
    const bare = await writeXlsx({ book })
    expect(issuesOf(bare)).toEqual([
      ['formula', ['A1', 'A2'], 'a formula the file cannot hold (it does not parse, or names a sheet, a table, a column or a link to another workbook the workbook does not have) and without a value to write instead is left out: its cell is written empty'],
      ['formula', ['A3'], 'a formula the file cannot hold (it does not parse, or names a sheet, a table, a column or a link to another workbook the workbook does not have) is written as its value'],
    ])
    expect(await sheetData(bare.data)).toBe('<sheetData><row r="1"><c r="A1" s="1"/></row><row r="2"><c r="A2"/></row><row r="3"><c r="A3"><v>4</v></c></row></sheetData>')
    const valued = await writeXlsx({ book, result: compute(book, { clock, execution }) })
    expect(valued.issues.find((i) => i.code === 'formula')?.cells).toEqual(['A1', 'A2', 'A3'])
  })

  it('#ERROR! written as #NULL! is reported, in values mode and from a file\'s stored values (fix-u329)', async () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: 'SUM(' }, A2: { formula: 'A1' }, A3: { formula: 'ERROR.TYPE(A1)' } } }] }
    const result = compute(book, { clock, execution })
    expect(result.sheets[0].cells.A2.value).toEqual({ error: 'ERROR' })
    const NULL = '#ERROR!, the error of a formula that does not compute, is written as #NULL!: Excel has no #ERROR!'
    const values = await writeXlsx({ book, result }, { mode: 'values' })
    expect(issuesOf(values)).toEqual([['errorValue', ['A1', 'A2'], NULL]])
    expect(await sheetData(values.data)).toBe('<sheetData><row r="1"><c r="A1" t="e"><v>#NULL!</v></c></row><row r="2"><c r="A2" t="e"><v>#NULL!</v></c></row><row r="3"><c r="A3"><v>1</v></c></row></sheetData>')
    // a stored #GETTING_DATA, read as #ERROR!, written back
    const read = await sheet('<row r="1"><c r="A1" t="e"><f>NOW()</f><v>#GETTING_DATA</v></c></row>')
    const again = await writeXlsx({ book: read.book, result: read.cached })
    expect(issuesOf(again)).toEqual([['errorValue', ['A1'], NULL]])
    expect((await writeXlsx({ book, result }, { locale: ruRU })).issues.find((i) => i.code === 'errorValue')?.message).toBe('#ERROR!, ошибка формулы, которая не вычисляется, записана как #ПУСТО!: в Excel нет #ERROR!')
  })
})
