/**
 * Reproductions from the review of 21.10 on reading xlsx: a damaged deflate stream fails the read alone, without a
 * rejection left unhandled; a part whose elements are left open reads in time linear in its length. Reading a file does
 * not calculate, so the execution modes do not apply here.
 */
import { deflateRawSync } from 'node:zlib'
import { describe, expect, it } from 'vitest'
import type { BookInput } from '../src/contract.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { rawZip, type RawEntry } from './fixtures/raw-zip.js'
import { pack, read, sheet } from './xlsx-fixtures.js'

/** The rejections nobody handled while `f` ran, and a moment after, when Node reports them. */
async function unhandledDuring(f: () => Promise<void>): Promise<unknown[]> {
  const seen: unknown[] = []
  const on = (e: unknown) => seen.push(e)
  process.on('unhandledRejection', on)
  try {
    await f()
    await new Promise((resolve) => setTimeout(resolve, 50))
  } finally {
    process.off('unhandledRejection', on)
  }
  return seen
}

describe('xlsx: a damaged deflate stream', () => {
  it('fails the read with the entry named, and leaves no rejection unhandled', async () => {
    const text = new TextEncoder().encode('<a>'.repeat(5000))
    const garbage: RawEntry = { name: 'a.xml', data: new Uint8Array(64).fill(0xff), method: 8, usize: 100 }
    // a valid stream cut in half: it ends before its last block
    const cut: RawEntry = { name: 'a.xml', data: deflateRawSync(text).subarray(0, 20), method: 8, usize: text.length }
    const { data } = await writeXlsx({ book: { sheets: [{ name: 'S', cells: { A1: { value: 'hello' } } }] } })
    const archive = openArchive(data)
    const entries = await Promise.all(archive.names.map(async (name): Promise<RawEntry> => ({ name, data: (await archive.bytes(name))! })))
    const damaged = rawZip(entries.map((e) => (e.name === 'xl/sharedStrings.xml' ? { ...e, data: new Uint8Array(32).fill(0xff), method: 8, usize: e.data.length } : e)))
    const unhandled = await unhandledDuring(async () => {
      for (const e of [garbage, cut]) await expect(openArchive(rawZip([e])).bytes('a.xml')).rejects.toThrow(/^xlsx: not a readable zip archive: entry a\.xml is not valid deflate data/)
      await expect(readXlsx(damaged)).rejects.toThrow(/^xlsx: not a readable zip archive: entry xl\/sharedStrings\.xml is not valid deflate data/)
      // the same file whole reads as before
      expect((await readXlsx(rawZip(entries))).book.sheets[0].cells.A1).toEqual({ value: 'hello' })
    })
    expect(unhandled).toEqual([])
  })
})

/** How long `f` takes, in milliseconds, and what it gives. */
async function timed<T>(f: () => Promise<T>): Promise<[number, T]> {
  const t = performance.now()
  const v = await f()
  return [performance.now() - t, v]
}

const strings = '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>'

describe('xlsx: elements left open', () => {
  // each case took seconds when each open tag searched the rest of its part for its closing tag (40,000 unclosed
  // <si>: 4.7 s); read in one pass, each takes tens of milliseconds
  const cases: [string, () => ReturnType<typeof read>, (r: Awaited<ReturnType<typeof read>>) => void][] = [
    ['shared strings', () => read({ sheets: [{ name: 'S', body: strings }], strings: [], parts: { 'xl/sharedStrings.xml': `<sst><si><t>a</t></si>${'<si><t>x</t>'.repeat(40_000)}</sst>` } }), (r) => expect(r.book.sheets[0].cells.A1).toEqual({ value: 'a' })],
    ['phonetic runs of a string', () => read({ sheets: [{ name: 'S', body: strings }], strings: [], parts: { 'xl/sharedStrings.xml': `<sst><si><t>a</t>${'<rPh sb="0" eb="1"><t>x</t>'.repeat(40_000)}</si></sst>` } }), (r) => expect(r.book.sheets[0].cells.A1?.value).toMatch(/^ax{40000}$/)],
    ['data validations', () => read({ sheets: [{ name: 'S', body: `<sheetData/><dataValidations count="1"><dataValidation type="whole" operator="equal" sqref="B1"><formula1>2</formula1></dataValidation>${'<dataValidation type="whole" operator="equal" sqref="A1"><formula1>1</formula1>'.repeat(20_000)}</dataValidations>` }] }), (r) => expect(r.book.sheets[0].validations).toHaveLength(1)],
    ['conditional formats', () => read({ sheets: [{ name: 'S', body: `<sheetData/><conditionalFormatting sqref="B1"><cfRule type="expression" priority="1"><formula>1</formula></cfRule></conditionalFormatting>${'<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>1</formula></cfRule>'.repeat(20_000)}` }] }), (r) => expect(r.book.sheets[0].conditionalFormats).toHaveLength(1)],
    ['inline strings of a cell', () => read({ sheets: [{ name: 'S', body: `<sheetData><row r="1"><c r="B1" t="inlineStr"><is><t>z</t></is></c><c r="A1" t="inlineStr">${'<is><t>x</t>'.repeat(40_000)}</row></sheetData>` }] }), (r) => expect(r.book.sheets[0].cells.B1).toEqual({ value: 'z' })],
    ['values and formulas of a cell', () => read({ sheets: [{ name: 'S', body: `<sheetData><row r="1"><c r="B1"><f>1+1</f><v>2</v></c><c r="A1">${'<f>1</x><v>1</x>'.repeat(30_000)}</row></sheetData>` }] }), (r) => expect(r.book.sheets[0].cells.B1).toEqual({ formula: '1+1' })],
    ['opening tags without their end', () => read({ sheets: [{ name: 'S', body: strings }], strings: [], parts: { 'xl/sharedStrings.xml': `<sst><si><t>a</t></si>${'<si a="1" '.repeat(40_000)}></sst>` } }), (r) => expect(r.book.sheets[0].cells.A1).toEqual({ value: 'a' })],
  ]
  it.each(cases)('%s', async (_, f, check) => {
    const [ms, r] = await timed(f)
    check(r)
    expect(ms).toBeLessThan(1000)
  }, 120_000)
})

/** The book written, its first sheet's XML changed by `edit`, and read back. */
async function rewritten(book: BookInput, edit: (xml: string) => string) {
  const archive = openArchive((await writeXlsx({ book })).data)
  const text = new TextDecoder()
  const parts = Object.fromEntries(await Promise.all(archive.names.map(async (n) => [n, text.decode((await archive.bytes(n))!)] as const)))
  parts['xl/worksheets/sheet1.xml'] = edit(parts['xl/worksheets/sheet1.xml'])
  return readXlsx(await pack(parts))
}

/** A sheet with Excel 2010 rules (a 3Stars icon set, a data bar's 2010 part) and 2007 ones, and two validations. */
const ruled: BookInput = {
  sheets: [
    {
      name: 'S',
      cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 } },
      conditionalFormats: [
        { ranges: ['A1:A3'], type: 'iconSet', icons: '3Stars', points: [{ value: { type: 'percent', value: '0' } }, { value: { type: 'percent', value: '33' } }, { value: { type: 'percent', value: '67' } }] },
        { ranges: ['C1:C3'], type: 'dataBar', min: { type: 'min' }, max: { type: 'max' }, color: '#638EC6', direction: 'rightToLeft' },
        { ranges: ['D1:D3'], type: 'cellIs', operator: 'greaterThan', formulas: ['1'], style: { font: { bold: true } } },
      ],
      validations: [
        { range: 'B1:B3', type: 'whole', operator: 'greaterThan', formula1: '0' },
        { range: 'E1:E3', type: 'list', formula1: 'T!$A$1:$A$2' },
      ],
    },
    { name: 'T', cells: { A1: { value: 'x' } } },
  ],
}

const rulesOf = (r: Awaited<ReturnType<typeof readXlsx>>) => ({ cf: r.book.sheets[0].conditionalFormats, dv: r.book.sheets[0].validations })

describe('xlsx: the worksheet\'s own extension list (fix-u327)', () => {
  const nested = '<extLst><ext uri="{DEADBEEF-0000-4000-8000-000000000000}"/></extLst>'
  it.each([
    ['an autofilter', (x: string) => x.replace('</sheetData>', `</sheetData><autoFilter ref="A1:A3">${nested}</autoFilter>`)],
    ['a filter column', (x: string) => x.replace('</sheetData>', `</sheetData><autoFilter ref="A1:A3"><filterColumn colId="0"><filters><filter val="1"/></filters>${nested}</filterColumn></autoFilter>`)],
    ['a sort state', (x: string) => x.replace('</sheetData>', `</sheetData><sortState ref="A1:A3"><sortCondition ref="A1:A3"/>${nested}</sortState>`)],
    ['a custom sheet view', (x: string) => x.replace('</sheetData>', `</sheetData><customSheetViews><customSheetView guid="{00000000-0000-4000-8000-000000000001}">${nested}</customSheetView></customSheetViews>`)],
    ['ignored errors', (x: string) => x.replace(/<extLst><ext uri="\{78C0/, `<ignoredErrors><ignoredError sqref="A1" numberStoredAsText="1"/>${nested}</ignoredErrors>$&`)],
  ])('an extension list inside %s is not the sheet\'s: rules and validations are read', async (_, edit) => {
    const plain = rulesOf(await rewritten(ruled, (x) => x))
    expect(plain.cf?.map((x) => x.type)).toEqual(['iconSet', 'dataBar', 'cellIs'])
    expect(plain.cf?.[1]).toMatchObject({ direction: 'rightToLeft' })
    expect(plain.dv).toHaveLength(2)
    const r = await rewritten(ruled, edit)
    expect(rulesOf(r)).toEqual(plain)
    expect(r.issues.filter((i) => i.code !== 'autoFilter')).toEqual([])
  })

  it('a sheet whose last element is no extension list has no Excel 2010 rules: a nested list is not taken for its own', async () => {
    const r = await rewritten(ruled, (x) => x.replace(/<extLst><ext uri="\{78C0[\s\S]*<\/extLst>(?=<\/worksheet>)/, '').replace('</sheetData>', `</sheetData><autoFilter ref="A1:A3"><extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="expression" priority="9"><xm:f>1</xm:f></x14:cfRule><xm:sqref>F1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></autoFilter>`))
    expect(rulesOf(r).cf?.map((x) => x.ranges[0])).toEqual(['C1:C3', 'D1:D3'])
  })
})

/** The sheet's XML with SpreadsheetML bound to the prefix `x:`, as the Open XML SDK writes it: other prefixes stay. */
const prefixed = (xml: string) => xml.replace(' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', ' xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"').replace(/<(\/?)([A-Za-z][\w.-]*)(?=[\s/>])/g, '<$1x:$2')

describe('xlsx: a worksheet whose SpreadsheetML elements have a prefix (fix-u324)', () => {
  it('keeps its conditional formats, 2007 and 2010, and its validations', async () => {
    const plain = await rewritten(ruled, (x) => x)
    const r = await rewritten(ruled, prefixed)
    expect(r.book.sheets[0].cells).toEqual(plain.book.sheets[0].cells)
    expect(rulesOf(r)).toEqual(rulesOf(plain))
    expect(r.issues).toEqual([])
  })

  it('a sheet of 2007 rules alone keeps them', async () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 2 } }, conditionalFormats: [{ ranges: ['A1:A3'], type: 'cellIs', operator: 'greaterThan', formulas: ['1'], style: { font: { bold: true } } }] }] }
    expect((await rewritten(book, prefixed)).book.sheets[0].conditionalFormats).toEqual([{ ranges: ['A1:A3'], type: 'cellIs', operator: 'greaterThan', formulas: ['1'], style: { font: { bold: true } } }])
  })
})

describe('xlsx: a damaged cell (x-xlsx-4, x-xlsx-8)', () => {
  const cellsOf = (r: Awaited<ReturnType<typeof read>>) => r.book.sheets[0].cells
  const issuesOf = (r: Awaited<ReturnType<typeof read>>) => r.issues.map((i) => [i.code, i.cells ?? [], i.message])

  it('a date of a Strict file that is no ISO date, or no date at all, is left out with an issue; the file reads', async () => {
    const r = await sheet('<row r="1"><c r="A1" t="d"><v>garbage</v></c><c r="B1" t="d"><v>2025-02-30</v></c><c r="C1" t="d"><v>2025-02-28T12:00:00</v></c><c r="D1"><v>1</v></c></row>')
    expect(cellsOf(r)).toEqual({ C1: { value: 45716.5 }, D1: { value: 1 } })
    expect(issuesOf(r)).toEqual([['part', ['A1', 'B1'], 'a date of a Strict file (a cell of type d) that is not a valid ISO date is left out']])
  })

  const OPEN = 'an element of a cell left open (without its closing tag): what it holds, and the cells it runs into, are not read'
  it.each([
    ['a cell without its closing tag before the next cell', '<row r="1"><c r="A1"><v>1</v><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c></row>', { A1: { value: 1 }, A2: { value: 3 } }, ['A1']],
    ['cells without closing tags to the end', '<row r="1"><c r="A1"><v>1</v><c r="B1"><v>2</v></row><row r="2"><c r="A2"><v>3</v></row>', { A1: { value: 1 } }, ['A1']],
    ['a value without its closing tag', '<row r="1"><c r="A1"><v>1</c><c r="B1"><v>2</v></c></row>', { B1: { value: 2 } }, ['A1']],
    ['a formula without its closing tag', '<row r="1"><c r="A1"><f>1+1</c><c r="B1"><v>2</v></c></row>', { B1: { value: 2 } }, ['A1']],
    ['an inline string without its closing tag', '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></c><c r="B1"><v>2</v></c></row>', { B1: { value: 2 } }, ['A1']],
    ['a prefixed cell', '<x:row r="1"><x:c r="A1"><x:v>1</x:v><x:c r="B1"><x:v>2</x:v></x:c></x:row>', { A1: { value: 1 } }, ['A1']],
  ])('%s is reported', async (_, rows, cells, at) => {
    const r = await sheet(rows)
    expect(cellsOf(r)).toEqual(cells)
    expect(issuesOf(r)).toEqual([['part', at, OPEN]])
  })

  it('rich text, empty values and self-closing cells are no elements left open', async () => {
    const r = await sheet('<row r="1"><c r="A1" t="inlineStr"><is><r><rPr><color rgb="FFFF0000"/><charset val="1"/></rPr><t>a</t></r></is></c><c r="B1" t="str"><f>""</f><v/></c><c r="C1" s="0"/><c r="D1"><v>4</v></c></row>')
    expect(cellsOf(r)).toEqual({ A1: { value: 'a' }, B1: { formula: '""' }, D1: { value: 4 } })
    expect(r.issues.map((i) => i.code)).toEqual(['richText'])
  })

  it('cells without </sheetData> are read to the end of the part, and reported', async () => {
    const r = await read({ sheets: [{ name: 'S', body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>' }] })
    expect(cellsOf(r)).toEqual({ A1: { value: 1 }, A2: { value: 2 } })
    expect(r.issues).toEqual([{ code: 'part', sheet: 'S', message: "the sheet's cells have no closing tag (</sheetData>): they are read to the end of the part" }])
    expect((await read({ sheets: [{ name: 'S', body: '<sheetData/><mergeCells><mergeCell ref="A1:B1"/></mergeCells>' }] })).issues).toEqual([])
  })
})

describe('xlsx: a width or height that is no size (x-xlsx-6)', () => {
  it('is left out with an issue; one above the highest is the highest, as before', async () => {
    const cols = ['abc', '-5', '1e400', 'NaN', '300', '12'].map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')
    const rows = ['-1', 'abc', '1e400', '409.6', '20'].map((h, i) => `<row r="${i + 1}" ht="${h}" customHeight="1"><c r="A${i + 1}"><v>1</v></c></row>`).join('')
    const r = await read({ sheets: [{ name: 'S', body: `<cols>${cols}</cols><sheetData>${rows}</sheetData>` }] })
    const s = r.book.sheets[0]
    expect(s.columns).toEqual({ E: { width: 255 }, F: { width: 12 } })
    expect(s.rows).toEqual({ 4: { height: 409.5 }, 5: { height: 20 } })
    expect(r.issues.map((i) => [i.code, i.cells, i.message])).toEqual([
      ['part', ['A:A', 'B:B', 'C:C', 'D:D'], 'a column width that is not a number of 0 or more is left out'],
      ['part', ['1:1', '2:2', '3:3'], 'a row height that is not a number of 0 or more is left out'],
    ])
  })
})

describe('xlsx: shared strings left open (sst-open)', () => {
  const two = '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData>'
  const OPEN = 'shared strings left open (an <si> without its closing tag) are read as empty text or as part of the string after them'
  it.each([
    ['the last one', '<sst><si><t>a</t></si><si><t>b</t></sst>', { A1: { value: 'a' }, B1: { value: '' } }],
    ['one before another', '<sst><si><t>a</t><si><t>b</t></si></sst>', { A1: { value: 'ab' }, B1: { value: '' } }],
    ['prefixed', '<x:sst><x:si><x:t>a</x:t></x:si><x:si><x:t>b</x:t></x:sst>', { A1: { value: 'a' }, B1: { value: '' } }],
  ])('%s is reported', async (_, sst, cells) => {
    const r = await read({ sheets: [{ name: 'S', body: two }], strings: [], parts: { 'xl/sharedStrings.xml': sst } })
    expect(r.book.sheets[0].cells).toEqual(cells)
    expect(r.issues).toEqual([{ code: 'part', message: OPEN }])
  })

  it('closed and empty ones are not', async () => {
    const r = await read({ sheets: [{ name: 'S', body: two }], strings: [], parts: { 'xl/sharedStrings.xml': '<sst><si><t>a</t></si><si/></sst>' } })
    expect([r.book.sheets[0].cells, r.issues]).toEqual([{ A1: { value: 'a' }, B1: { value: '' } }, []])
  })
})
