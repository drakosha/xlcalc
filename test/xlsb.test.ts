import { workbookData } from './print-fixtures.js'
import { readFile } from 'node:fs/promises'
import { describe, expect, it } from 'vitest'
import { FUNCTION_COUNTS } from '../src/function-words.js'
import { compute } from '../src/index.js'
import { readXlsb } from '../src/xlsb.js'
import { readXlsx } from '../src/xlsx.js'
import { FTAB } from '../src/xlsb/ftab.js'
import { records } from '../src/xlsb/records.js'
import { Xml } from '../src/xlsb/xml.js'
import { XLFN, XLWS } from '../src/xlsx/formulas/prefixes.js'
import { zip } from '../src/xlsx/zip/write.js'
import { rels } from './xlsx-fixtures.js'

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const workbook = 'xl/workbook.bin'
const sheet = 'xl/worksheets/sheet1.bin'
const link = 'xl/externalLinks/externalLink1.bin'
const bytes = (value: string) => new TextEncoder().encode(value)
const related = (type: string, target: string, id = 'rId1') => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`
const excelPair = (name: string) => Promise.all(['xlsb', 'xlsx'].map((ext) =>
  readFile(new URL(`./excel/win/xlsb-read/${name}.${ext}`, import.meta.url))))
const sheetData = (cells: number[]) => Uint8Array.of(
  0x81, 1, 0, // BrtBeginSheet
  0x91, 1, 0, // BrtBeginSheetData
  0, 17, ...Array<number>(17).fill(0), // BrtRowHdr, row 1 without formatting or column spans
  ...cells,
  0x92, 1, 0, // BrtEndSheetData
  0x82, 1, 0, // BrtEndSheet
)

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
  it('checks native function arities and the names supplied by Excel prefixes', () => {
    const native = new Set<string>()
    const incompatible: string[] = []
    for (const entry of Object.values(FTAB)) {
      if (!entry) continue
      const [name, arity] = entry
      native.add(name)
      const range = FUNCTION_COUNTS.get(name)
      if (!range || arity === undefined) continue
      const [min, max, step] = range
      if (arity < min || arity > max || (arity - min) % step !== 0) incompatible.push(name)
    }
    expect(incompatible).toEqual([])
    const neither = [...FUNCTION_COUNTS.keys()].filter((name) => !native.has(name) && !XLFN.has(name) && !XLWS.has(name))
    expect(neither.sort()).toEqual(['EUROCONVERT', 'JIS'])
  })

  const pairs = [
    'probe', 'formulas', 'pivot-calculated', 'slicers-and-timeline', 'icon-filter', 'filter-and', 'filter-or',
    'filter-extended-and', 'filter-extended-or', 'filter-date-day', 'filter-date-month-year',
    'filter-long-value', 'filter-long-and-short-values', 'filter-no-icon', 'filter-non-blanks', 'filter-sort-stars',
    'array-exact', 'array-convert',
    'validation-multi-range', 'validation-list-comma', 'validation-invalid-reference',
    'conditional-invalid-reference', 'pivot-show-as', 'pivot-item-number',
    'union-arguments', 'unprefixed-function', 'worksheet-function-shadow', 'worksheet-function-name-case',
    'iteration-default', 'range-sheet-quote',
    'range-start-name', 'range-sheet-span', 'formula-closing-space', 'formula-small-numbers', 'formula-whole-sheet',
    'name-local-sheet', 'name-global-self',
    'table-totals-offset', 'table-totals-formula',
    'functions-001', 'functions-066', 'functions-131', 'functions-196',
    'functions-261', 'functions-326', 'functions-391', 'functions-456', 'functions-references',
  ]
  it.each(pairs)('reads both of Excel’s %s saves identically', async (name) => {
    const [binary, xml] = await excelPair(name)
    await expect(readXlsb(binary).then(workbookData)).resolves.toStrictEqual(await readXlsx(xml).then(workbookData))
  })

  it.each([
    { fixture: 'let-name-case', sheet: 1, address: 'A41', xlsx: 'LET(X,1,X+1)', xlsb: 'LET(x,1,x+1)', value: 2 },
    { fixture: 'let-name-case-spill', sheet: 0, address: 'B3', xlsx: 'LET(x,,x)', xlsb: 'LET(X,,X)', value: 0 },
  ])('uses the stored spelling for LET names in $fixture', async ({ fixture, sheet, address, xlsx, xlsb, value }) => {
    const [binary, xml] = await excelPair(fixture)
    const expected = await readXlsx(xml).then(workbookData)
    const actual = await readXlsb(binary).then(workbookData)
    const cell = expected.book.sheets[sheet].cells[address]

    expect(cell.formula).toBe(xlsx)
    expect(compute(expected.book).sheets[sheet].cells[address].value).toBe(value)
    expect(compute(actual.book).sheets[sheet].cells[address].value).toBe(value)

    // PtgName uses the same name record for x and X.
    cell.formula = xlsb
    expect(actual).toStrictEqual(expected)
  })

  it.each([100000, 165536])('keeps the result cached for an external name whose row %i Excel could not store in XLSB', async (row) => {
    const [binary, xml] = await excelPair(`external-name-${row}`)
    const expected = await readXlsx(xml).then(workbookData)
    const actual = await readXlsb(binary).then(workbookData)
    const name = expected.book.externalBooks![0].names![0]
    expect(name).toEqual({ name: 'FarName', ref: `'Source'!$A$${row}` })
    expect(compute(expected.book).sheets[0].cells.A2.value).toBe(9)
    expect(actual.cached.sheets[0].cells.A2.value).toBe(9)
    expect(compute(actual.book).sheets[0].cells.A2.value).toEqual({ error: 'REF' })

    // BrtSupNameFmla contains ExtPtgRefErr3D. Its coordinate bytes are undefined.
    name.ref = '#REF!'
    expected.issues.push({
      code: 'externalLink',
      name: 'FarName',
      message: 'the binary part xl/externalLinks/externalLink1.bin stores an invalid reference for this external name',
    })
    expect(actual).toStrictEqual(expected)
  })

  it('reports the same cells when Excel expands shared formulas in XLSB', async () => {
    const [binary, xml] = await excelPair('shared-issue-order')
    const expected = await readXlsx(xml).then(workbookData)
    const actual = await readXlsb(binary).then(workbookData)
    // The XLSX reader reports the two masters first, then each shared group. XLSB stores eight separate formulas.
    expect(expected.issues[0].cells).toEqual(['E1', 'F1', 'E2', 'E3', 'E4', 'F2', 'F3', 'F4'])
    expected.issues[0].cells = ['E1', 'F1', 'E2', 'F2', 'E3', 'F3', 'E4', 'F4']
    expect(actual).toStrictEqual(expected)
  })

  it('resolves the workbook and worksheets from their original package paths', async () => {
    const data = await fixture(Uint8Array.of(0x81, 1, 0, 0x82, 1, 0))
    const read = await readXlsb(data)
    expect(read.book.sheets).toEqual([{ name: 'S', cells: {} }])
  })

  it.each([
    [0x3a, [1, 0, 3, 0], "'S'!$D$2", false],
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

  it('refuses a cell index into a missing shared string', async () => {
    const cell = [7, 12, ...Array<number>(12).fill(0)] // BrtCellIsst, A1, shared string 0
    await expect(readXlsb(await fixture(sheetData(cell)))).rejects.toThrow(/sheet1\.bin at offset \d+: record 7: shared string index 0 out of bounds/)
  })

  it('reads a shared formula anchored away from the top-left corner of its range', async () => {
    const anchor = [5, 0, 0, 0, 1, 0, 0, 0, 0, 4, 0, 0, 0, 2, 0, 0, 0] // PtgExp and PtgExtraCol, C1
    const tokens = [0x4c, 0, 0, 0, 0, 0xff, 0xff, 0x44, 0, 0, 0, 0, 0, 0, 3] // one cell left + $A$1
    const data = await fixture(sheetData([
      2, 12, 0, 0, 0, 0, 0, 0, 0, 0, 42, 0, 0, 0, // BrtCellRk, A1 = 10
      2, 12, 1, 0, 0, 0, 0, 0, 0, 0, 82, 0, 0, 0, // BrtCellRk, B1 = 20
      9, 35, 2, ...Array<number>(17).fill(0), ...anchor, // BrtFmlaNum, C1
      0xab, 3, 39, 0, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 2, 0, 0, 0, // BrtShrFmla, B1:C2
      tokens.length, 0, 0, 0, ...tokens, 0, 0, 0, 0,
      0, 17, 1, ...Array<number>(16).fill(0), // BrtRowHdr, row 2
      2, 12, 0, 0, 0, 0, 0, 0, 0, 0, 122, 0, 0, 0, // BrtCellRk, A2 = 30
      9, 35, 1, ...Array<number>(17).fill(0), ...anchor, // BrtFmlaNum, B2
    ]))
    const { book, issues } = await readXlsb(data)
    expect(issues).toEqual([])
    expect(book.sheets[0].cells).toEqual({
      A1: { value: 10 }, B1: { value: 20 }, C1: { formula: 'B1+$A$1' },
      A2: { value: 30 }, B2: { formula: 'A2+$A$1' },
    })
    const { cells } = compute(book).sheets[0]
    expect(cells.C1.value).toBe(30)
    expect(cells.B2.value).toBe(40)
  })

  it('refuses a shared formula that points back to its own anchor', async () => {
    const anchor = [5, 0, 0, 0, 1, 0, 0, 0, 0, 4, 0, 0, 0, 0, 0, 0, 0] // PtgExp, A1
    const data = await fixture(sheetData([
      9, 35, ...Array<number>(18).fill(0), ...anchor, // BrtFmlaNum, A1
      0xab, 3, 33, ...Array<number>(16).fill(0), ...anchor, // BrtShrFmla, A1:A1
    ]))
    await expect(readXlsb(data)).rejects.toThrow(/sheet1\.bin at offset \d+: record 427: array or shared formula references another anchor/)
  })

  it('charges expanded UTF-8 XML before joining it', () => {
    const entry = new Xml(sheet, { entryBytes: 5 }, { spent: 0 })
    expect(() => entry.add('é🙂')).toThrow(/sheet1\.bin.*limits\.entryBytes/)
    const total = new Xml(sheet, { entryBytes: 6, totalBytes: 8 }, { spent: 3 })
    // Six expanded bytes plus the three bytes already read from the ZIP.
    expect(() => total.add('é🙂')).toThrow(/sheet1\.bin.*limits\.totalBytes/)
  })
})
