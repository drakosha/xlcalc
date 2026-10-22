/**
 * A table's style (fix-u330): read from a file's `tableStyleInfo`, checked as input, written back, and kept by the
 * operations on tables, in an `editStructure` and in a session of each kind. Without it a table is written with
 * Excel's default style, as before.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createSession, editStructure, InputError, type BookInput, type Execution, type StructureOp, type TableStyle } from '../src/index.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'
import { pack } from './xlsx-fixtures.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const EXECUTIONS: Execution[] = ['closures', 'eval']
const DEFAULT = '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'

afterAll(closeCalcs)

const tableBook = (style?: TableStyle): BookInput => ({
  sheets: [{ name: 'S', cells: { A1: { value: 'Qty' }, A2: { value: 1 }, A3: { value: 2 } }, tables: [{ name: 'Sales', range: 'A1:A3', columns: [{ name: 'Qty' }], ...(style ? { style } : {}) }] }],
})

/** The parts of the file written from `book`, as text. */
async function partsOf(book: BookInput): Promise<Record<string, string>> {
  const archive = openArchive((await writeXlsx({ book })).data)
  const text = new TextDecoder()
  return Object.fromEntries(await Promise.all(archive.names.map(async (n) => [n, text.decode((await archive.bytes(n))!)] as const)))
}

/** The book's file with its table's style info replaced by `info`, read back. */
async function withInfo(info: string, opts?: Parameters<typeof readXlsx>[1]) {
  const parts = await partsOf(tableBook())
  parts['xl/tables/table1.xml'] = parts['xl/tables/table1.xml'].replace(DEFAULT, info)
  return readXlsx(await pack(parts), opts)
}

const infoOf = async (book: BookInput) => /<tableStyleInfo[^>]*\/>/.exec((await partsOf(book))['xl/tables/table1.xml'])?.[0]

describe('a table\'s style in xlsx (fix-u330)', () => {
  it.each([
    ['another style with its first column emphasized and no bands', '<tableStyleInfo name="TableStyleLight1" showFirstColumn="1" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/>', { name: 'TableStyleLight1', firstColumn: true }],
    ['the default style banded by columns', '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="1" showRowStripes="0" showColumnStripes="1"/>', { name: 'TableStyleMedium2', lastColumn: true, columnStripes: true }],
    ['no style', '<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>', { rowStripes: true }],
    ['no style info at all', '', {}],
  ])('%s is read, and written back as it was', async (_, info, style) => {
    const r = await withInfo(info)
    expect(r.issues).toEqual([])
    expect(r.book.sheets[0].tables?.[0].style).toEqual(style)
    const written = await infoOf(r.book)
    expect(written).toBe(info === '' ? '<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/>' : info)
  })

  it("Excel's default style is a table without a style, written as before", async () => {
    const r = await withInfo(DEFAULT)
    expect(r.book.sheets[0].tables?.[0]).not.toHaveProperty('style')
    expect(await infoOf(tableBook())).toBe(DEFAULT)
    expect(await infoOf(tableBook({ name: 'TableStyleMedium2', rowStripes: true }))).toBe(DEFAULT)
  })

  it('a custom style is reported and read as the default', async () => {
    const info = '<tableStyleInfo name="My Style" showFirstColumn="1" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'
    const r = await withInfo(info)
    expect(r.book.sheets[0].tables?.[0]).not.toHaveProperty('style')
    expect(r.issues).toEqual([{ code: 'table', sheet: 'S', message: 'the custom table style "My Style" of the table Sales is not in the model: the table is read with Excel\'s default style' }])
    expect((await withInfo(info, { locale: ruRU })).issues[0].message).toBe('Собственный стиль «My Style» таблицы Sales не поддерживается: таблица прочитана со стилем Excel по умолчанию')
  })
})

describe("a table's style as input (fix-u330)", () => {
  it.each([
    ['not an object', 'TableStyleLight1', /book\.sheets\[0\]\.tables\[0\]\.style: expected a style such as/],
    ['an unknown field', { name: 'TableStyleLight1', bands: true }, /book\.sheets\[0\]\.tables\[0\]\.style: unknown field "bands"/],
    ['a name that is none of Excel\'s styles', { name: 'TableStyleLight22' }, /book\.sheets\[0\]\.tables\[0\]\.style\.name: expected one of Excel's table styles/],
    ['an option that is no boolean', { rowStripes: 1 }, /book\.sheets\[0\]\.tables\[0\]\.style\.rowStripes: expected a boolean/],
  ])('%s is an InputError', (_, style, message) => {
    const book = tableBook(style as TableStyle)
    expect(() => compute(book)).toThrow(InputError)
    expect(() => compute(book)).toThrow(message)
  })

  it.each([{}, { name: 'TableStyleDark11', columnStripes: true }, { name: 'TableStyleMedium28', firstColumn: false }])('%j is a style', (style) => {
    expect(() => compute(tableBook(style))).not.toThrow()
  })
})

const STYLE: TableStyle = { name: 'TableStyleLight9', firstColumn: true }
const styled = () => tableBook(STYLE)
const OPS: [string, StructureOp, (b: BookInput) => unknown][] = [
  ['renameTable', { op: 'renameTable', table: 'Sales', name: 'Orders' }, (b) => b.sheets[0].tables?.[0].style],
  ['insertRows into the table', { op: 'insertRows', sheet: 'S', row: 3 }, (b) => b.sheets[0].tables?.[0].style],
  ['moveRange of the table', { op: 'moveRange', from: { sheet: 'S', range: 'A1:A3' }, to: { sheet: 'S', cell: 'C5' } }, (b) => b.sheets[0].tables?.[0].style],
  ['copyRange of the table', { op: 'copyRange', from: { sheet: 'S', range: 'A1:A3' }, to: { sheet: 'S', range: 'C5' } }, (b) => b.sheets[0].tables?.map((t) => t.style)],
]

describe.each(EXECUTIONS)("operations keep a table's style (fix-u330), %s", (execution) => {
  describe.each(OPS)('%s', (_, op, styleOf) => {
    const expected = op.op === 'copyRange' ? [STYLE, STYLE] : STYLE
    it('in editStructure', () => {
      const r = editStructure(styled(), op)
      expect(r.ok).toBe(true)
      if (r.ok) expect(styleOf(r.book)).toEqual(expected)
    })
    describe.each(KINDS)('in %s', (kind) => {
      const open = openerFor(kind)
      it('and in the book the session holds', async () => {
        const s = await open(styled(), { execution, clock })
        expect((await s.editStructure(op)).ok).toBe(true)
        expect(styleOf(await s.getBook())).toEqual(expected)
        await s.close()
      })
    })
  })

  it('a session written to a file keeps it', async () => {
    const s = createSession(styled(), { execution, clock })
    const parts = await (async () => {
      const archive = openArchive((await writeXlsx(s.getSnapshot())).data)
      return new TextDecoder().decode((await archive.bytes('xl/tables/table1.xml'))!)
    })()
    expect(parts).toContain('<tableStyleInfo name="TableStyleLight9" showFirstColumn="1" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/>')
  })
})
