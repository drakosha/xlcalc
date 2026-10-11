/**
 * The number format is `numberFormat` in every input that takes one (cells, rows, columns, conditional formats' styles,
 * edits of them) and every answer; its old name `format` is refused (`old-names.test.ts`). A book hashes as it did
 * before the field was renamed (`inputHash` writes it by the old name), on the synthetic book and on one with formats
 * of cells, rows, columns and conditional formats. A book an answer gives (`Book`) has no old name, and its type none to
 * read. In the three kinds of session, with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, editStructure, type Book, type BookInput, type BookLayout } from '../src/index.js'
import { synthBook } from '../src/synth.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

/**
 * A book an answer gave has no `format` on its cells, rows, columns or conditional formats' styles, and its type
 * (`Book`) has none either: reading the old name does not compile (`npm run typecheck`), where it would be undefined.
 */
function noOldName(book: Book | BookLayout) {
  const sheet = book.sheets[0]
  const parts: object[] = [...Object.values('cells' in sheet ? sheet.cells : {}), ...Object.values(sheet.rows ?? {}), ...Object.values(sheet.columns ?? {})]
  for (const cf of sheet.conditionalFormats ?? []) if ('style' in cf && cf.style) parts.push(cf.style)
  expect(parts.filter((x) => 'format' in x)).toEqual([])
  expect(sheet.rows?.[3]?.numberFormat).toBeDefined()
  // @ts-expect-error a row names its number format numberFormat alone
  expect(sheet.rows?.[3]?.format).toBeUndefined()
  // @ts-expect-error so does a column
  expect(sheet.columns?.B?.format).toBeUndefined()
  const cf = sheet.conditionalFormats?.[0]
  // @ts-expect-error and a conditional format's style
  if (cf?.type === 'cellIs') expect(cf.style?.format).toBeUndefined()
  if (!('cells' in sheet)) return
  // @ts-expect-error and a cell
  expect(sheet.cells.A1.format).toBeUndefined()
}

/** A book with number formats of cells, rows, columns and a conditional format; `meta` is the host's own and may hold anything. */
const formatted = (): BookInput => ({
  styles: [{ font: { bold: true } }],
  sheets: [{
    name: 'S',
    cells: { A1: { value: 1, numberFormat: '0.00' }, A2: { value: 2, numberFormat: 14, style: 0, meta: { format: 'x', numberFormat: 'y' } }, B1: { formula: 'A1*2' } },
    columns: { B: { numberFormat: '0%', width: 12 }, 'C:D': { numberFormat: 4 } },
    rows: { 3: { numberFormat: '#,##0', height: 20 } },
    conditionalFormats: [{ type: 'cellIs', ranges: ['A1:A3'], operator: 'greaterThan', formulas: ['1'], style: { font: { bold: true }, numberFormat: '0.0' } }],
  }],
})

describe('the input hash, as before the field was renamed', () => {
  it('the synthetic book hashes as before', () => {
    const r = compute(synthBook({ seed: 1, rows: 3400 }).book, { clock: { now: '2026-10-04T12:00:00Z', seed: 7 } })
    expect([r.hash.slice(0, 16), r.inputHash.slice(0, 16)]).toEqual(['101578a9c6b600b1', '1fb3575501a5acbf'])
  })

  it('a book with formats of cells, rows, columns and a conditional format hashes as before', () => {
    const r = compute(formatted(), { clock: { now: '2026-10-04T12:00:00Z', seed: 7 } })
    // the hashes the book gave before the field was named numberFormat
    expect([r.inputHash, r.hash]).toEqual(['608a314eb5783277d10040697b035f974b1b998f25f4186d4c80b29ba0deef7f', 'd7b7f2addfc626771ef0c924e2407ce19fd3334e0856e12d3206f8fecbf40e38'])
  })
})

describe('the number format outside a session', () => {
  it('results answer numberFormat, the input not changed', () => {
    const book = formatted()
    const cells = compute(book, { clock: CLOCK }).sheets[0].cells
    expect(cells.A1).toEqual({ value: 1, text: '1.00', numberFormat: '0.00' })
    expect(cells.B1).toEqual({ value: 2, text: '200%', formula: 'A1*2' })
    expect(book.sheets[0].cells.A1).toEqual({ value: 1, numberFormat: '0.00' })
  })

  it('xlsx: a book written is read back by the same name', async () => {
    const book = formatted()
    const read = (await readXlsx((await writeXlsx({ book, result: compute(book, { clock: CLOCK }) })).data)).book
    expect([read.sheets[0].cells.A1, read.sheets[0].rows]).toEqual([{ value: 1, numberFormat: 2 }, { 3: { height: 20, numberFormat: 3 } }])
  })

  it("the books readXlsx and editStructure answer are Books: no old name, in the type either; a Book goes back as a book's input", async () => {
    const book = formatted()
    const read = (await readXlsx((await writeXlsx({ book, result: compute(book, { clock: CLOCK }) })).data)).book
    noOldName(read)
    const moved = editStructure(book, { op: 'insertColumns', sheet: 'S', column: 'Z' })
    if (!moved.ok) throw new Error(moved.message)
    noOldName(moved.book)
    expect(compute(moved.book, { clock: CLOCK }).inputHash).toBe(compute(book, { clock: CLOCK }).inputHash)
  })
})

describe.each(['closures', 'eval'] as const)('the number format in sessions, %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it('a book, edits, patches and layout edits answer and keep it', async () => {
      const s = await (await opener(kind))(formatted(), { clock: CLOCK, execution })
      expect((await s.getBook()).sheets[0]).toMatchObject(formatted().sheets[0])
      const set = await s.setCells([{ sheet: 'S', cell: 'C1', input: { value: 0.5, numberFormat: '0.0%' } }])
      expect(set.changes).toEqual([{ sheet: 'S', cell: 'C1', value: 0.5, text: '50.0%', numberFormat: '0.0%', numberFormatCode: '0.0%' }])
      await s.replaceCells([{ sheet: 'S', cell: 'C2', input: { value: 3, numberFormat: '0.0' } }])
      await s.patchCells([{ sheet: 'S', cell: 'D1', set: { numberFormat: '0.000' } }])
      await s.setLayout([{ sheet: 'S', rows: '5', set: { numberFormat: '0.0000' } }, { sheet: 'S', conditionalFormats: [{ type: 'expression', ranges: ['E1'], formula: 'TRUE', style: { numberFormat: '0' } }] }])
      const sh = (await s.getBook()).sheets[0]
      expect([sh.cells.C1, sh.cells.C2, sh.cells.D1, sh.rows?.[5], (sh.conditionalFormats?.[0] as { style?: object }).style]).toEqual([{ value: 0.5, numberFormat: '0.0%' }, { value: 3, numberFormat: '0.0' }, { numberFormat: '0.000' }, { numberFormat: '0.0000' }, { numberFormat: '0' }])
      expect((await s.getCells('S', 'C2')).C2).toEqual({ value: 3, text: '3.0', numberFormat: '0.0', numberFormatCode: '0.0' })
      // null removes the format
      await s.setCells([{ sheet: 'S', cell: 'C1', input: { numberFormat: null } }])
      expect((await s.getBook()).sheets[0].cells.C1).toEqual({ value: 0.5 })
      expect((await s.validateEntry('S', 'A1', { value: 2, numberFormat: '0' })).ok).toBe(true)
      expect((await s.getResult()).inputHash).toBe(compute(await s.getBook(), { clock: CLOCK }).inputHash)
      await s.close()
    })

    it('the books it answers are Books: getBook, getSnapshot and getLayout, a cell\'s appearance too', async () => {
      const s = await (await opener(kind))(formatted(), { clock: CLOCK, execution })
      noOldName(await s.getBook())
      noOldName((await s.getSnapshot()).book)
      noOldName(await s.getLayout())
      const look = (await s.getCells('S', 'A2')).A2.appearance?.style
      expect(look).toEqual({ font: { bold: true }, numberFormat: '0.0' })
      // @ts-expect-error a conditional format's appearance names its number format numberFormat alone
      expect(look?.format).toBeUndefined()
      await s.close()
    })
  })
})
