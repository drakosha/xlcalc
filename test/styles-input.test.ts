/** Styles and the sheet layout as input: their shape is checked, they do not touch values, they go into `inputHash` only when present. */
import { describe, expect, it } from 'vitest'
import { compute, InputError, restructure, type BookInput } from '../src/index.js'

const clock = { now: '2026-10-06T12:00:00Z', seed: 1 }

const styled = (): BookInput => ({
  sheets: [
    {
      name: 'S',
      cells: { A1: { value: 1, style: 0 }, A2: { formula: 'A1*2', style: 1, format: '0.00' }, B1: { style: 1 } },
      columns: { B: { width: 20 }, 'D:F': { width: 12, style: 0, format: '0%' } },
      rows: { 1: { height: 30 }, '4:6': { style: 1 } },
      merges: ['C3:D4', 'A8:A9'],
    },
  ],
  styles: [
    { font: { name: 'Arial', size: 12, bold: true, italic: false, underline: 'double', strike: true, color: '#1F4E79' } },
    { fill: { color: '#ffc000' }, border: { top: { style: 'thin' }, bottom: { style: 'double', color: '#000000' } }, alignment: { horizontal: 'center', vertical: 'top', wrap: true, indent: 2 } },
  ],
})

function fails(book: unknown, message: string) {
  let err: unknown
  try {
    compute(book as BookInput, { clock })
  } catch (e) {
    err = e
  }
  expect(err).toBeInstanceOf(InputError)
  expect((err as Error).message).toBe(message)
}

const withSheet = (patch: object): BookInput => {
  const b = styled()
  return { ...b, sheets: [{ ...b.sheets[0], ...patch }] }
}

describe('styles and layout as input', () => {
  it('a styled book computes the same values and values hash as without its styles', () => {
    const b = styled()
    const plain: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { formula: 'A1*2', format: '0.00' }, B1: {} } }] }
    const [x, y] = [compute(b, { clock, display: false }), compute(plain, { clock, display: false })]
    expect(x.hash).toBe(y.hash)
    expect(x.sheets[0].cells.A2).toEqual({ value: 2, formula: 'A1*2', format: '0.00' })
  })

  it('inputHash: a book without styles hashes as before; styles, columns, rows and merged areas count, their order does not', () => {
    const plain: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] }
    const bare = compute(plain, { clock }).inputHash
    expect(compute({ ...plain, styles: [] }, { clock }).inputHash).toBe(bare)
    expect(compute({ sheets: [{ ...plain.sheets[0], columns: {}, rows: {}, merges: [] }] }, { clock }).inputHash).toBe(bare)
    const b = styled()
    const h = compute(b, { clock }).inputHash
    const shuffled: BookInput = { ...b, sheets: [{ ...b.sheets[0], merges: ['A8:A9', 'C3:D4'], rows: { '4:6': { style: 1 }, 1: { height: 30 } } }] }
    expect(compute(shuffled, { clock }).inputHash).toBe(h)
    for (const changed of [withSheet({ merges: ['C3:D4'] }), withSheet({ columns: { B: { width: 21 } } }), withSheet({ rows: { 1: { height: 30 } } }), { ...b, styles: [b.styles![1], b.styles![0]] }]) expect(compute(changed, { clock }).inputHash).not.toBe(h)
  })

  it('a wrong style, layout or sheet name throws InputError with the path', () => {
    const b = styled()
    const style = (s: unknown) => ({ ...b, styles: [s] })
    fails(style({ font: { bold: 1 } }), 'book.styles[0].font.bold: expected a boolean, got number 1')
    fails(style({ font: { colour: '#000000' } }), 'book.styles[0].font: unknown field "colour" (expected name, size, bold, italic, underline, strike, color)')
    fails(style({ fill: { color: 'red' } }), 'book.styles[0].fill.color: expected a colour such as "#FF0000", got the string "red"')
    fails(style({ border: { top: { style: 'bold' } } }), 'book.styles[0].border.top.style: expected one of "thin", "medium", "thick", "dotted", "dashed", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot", got the string "bold"')
    fails(style({ alignment: { indent: 1.5 } }), 'book.styles[0].alignment.indent: expected an integer from 0 to 250, got number 1.5')
    fails({ ...b, styles: {} }, 'book.styles: expected an array of styles, got an object')
    fails(withSheet({ cells: { A1: { style: 2 } } }), 'book.sheets[0].cells["A1"].style: expected the index of a style of the book, 0 to 1, got number 2')
    fails({ sheets: [{ name: 'S', cells: { A1: { style: 0 } } }] }, 'book.sheets[0].cells["A1"].style: the book has no styles, got number 0')
    fails(withSheet({ columns: { 'B:B': { width: 1 } } }), 'book.sheets[0].columns["B:B"]: expected letters such as "C" or a span such as "B:D" within the sheet, got the string "B:B"')
    fails(withSheet({ columns: { B: { width: 300 } } }), 'book.sheets[0].columns["B"].width: expected a number from 0 to 255, got number 300')
    fails(withSheet({ columns: { 'B:D': {}, C: {} } }), 'book.sheets[0].columns: "C" overlaps "B:D"')
    fails(withSheet({ rows: { 0: { height: 1 } } }), 'book.sheets[0].rows["0"]: expected a row number such as "5" or a span such as "5:9" within the sheet, got the string "0"')
    fails(withSheet({ rows: { 5: { hidden: true } } }), 'book.sheets[0].rows["5"]: unknown field "hidden" (expected height, style, format)')
    fails(withSheet({ merges: ['C3'] }), 'book.sheets[0].merges[0]: expected a merged area such as "B2:D3" (top-left to bottom-right, two cells or more), got the string "C3"')
    fails(withSheet({ merges: ['D4:C3'] }), 'book.sheets[0].merges[0]: expected a merged area such as "B2:D3" (top-left to bottom-right, two cells or more), got the string "D4:C3"')
    fails(withSheet({ merges: ['C3:D4', 'A1:B2', 'D4:E5'] }), 'book.sheets[0].merges[2]: D4:E5 overlaps C3:D4')
    fails(withSheet({ mergers: [] }), 'book.sheets[0]: unknown field "mergers" (expected name, cells, columns, rows, merges)')
    fails({ ...b, style: [] }, 'book: unknown field "style" (expected sheets, names, styles, locale)')
    fails(withSheet({ name: 'Q1/Q2' }), 'book.sheets[0].name: a sheet name must not contain "/" (Excel forbids : \\ / ? * [ ])')
    fails(withSheet({ name: 'x'.repeat(32) }), `book.sheets[0].name: a sheet name has at most 31 characters, "${'x'.repeat(32)}" has 32`)
    fails(withSheet({ name: "'Data" }), 'book.sheets[0].name: a sheet name must not begin or end with an apostrophe')
    fails(withSheet({ name: 'HISTORY' }), 'book.sheets[0].name: Excel reserves the sheet name "History" (in any case)')
    expect(compute(withSheet({ name: "It's (2026)" }), { clock }).sheets[0].name).toBe("It's (2026)")
  })

  it('renameSheet refuses a name Excel refuses', () => {
    const b = styled()
    expect(restructure(b, { op: 'renameSheet', sheet: 'S', name: 'a:b' })).toMatchObject({ ok: false, code: 'sheet_name', message: 'a sheet name must not contain ":" (Excel forbids : \\ / ? * [ ])' })
    expect(restructure(b, { op: 'renameSheet', sheet: 'S', name: 'y'.repeat(31) }).ok).toBe(true)
    expect(restructure(b, { op: 'renameSheet', sheet: 'S', name: 'history' })).toMatchObject({ ok: false, code: 'sheet_name' })
    expect(restructure(b, { op: 'renameSheet', sheet: 'S', name: 'My History' }).ok).toBe(true)
  })
})
