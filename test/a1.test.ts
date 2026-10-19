/**
 * Addresses without the engine (`xlcalc/a1`, the same functions in the root): cells and ranges read and written, the
 * sheet's edges, layout keys as a book takes them, rectangles, and the rule names compare by, which is the engine's own.
 */
import { describe, expect, it } from 'vitest'
import * as a1 from '../src/a1.js'
import * as root from '../src/index.js'
import { caseKey } from '../src/collate.js'
import { checkBook, InputError } from '../src/input.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { cellAddress, colLetters, colNumber, linesKey, MAX_COLS, MAX_ROWS, mergedArea, nameKey, parseCell, parseLines, parseRange, quoteSheetIfNeeded, rangeAddress, rectContains, rectsMeet, sameName } from '../src/a1.js'
import { editStructure } from '../src/index.js'

describe('cells', () => {
  it('reads a cell with $ and letters in any case, rows and columns from 1', () => {
    expect(parseCell('B12')).toEqual({ row: 12, col: 2 })
    expect(parseCell('$B$12')).toEqual({ row: 12, col: 2 })
    expect(parseCell('b$12')).toEqual({ row: 12, col: 2 })
    expect(parseCell('XFD1048576')).toEqual({ row: MAX_ROWS, col: MAX_COLS })
    expect(parseCell('A01')).toEqual({ row: 1, col: 1 })
  })

  it('null past the sheet and for anything else', () => {
    for (const t of ['ZZZ1', 'XFE1', 'A1048577', 'A9999999', 'A0', 'AAAA1', 'B2:C3', 'B', '12', '', ' B2', 'B 2', 'B$$2', 'B2$', 'ßB1', '$', 'A1B']) expect(parseCell(t), t).toBeNull()
  })

  it('writes a cell and column letters', () => {
    expect(cellAddress(12, 2)).toBe('B12')
    expect(cellAddress(MAX_ROWS, MAX_COLS)).toBe('XFD1048576')
    expect([1, 26, 27, 702, 703, MAX_COLS].map(colLetters)).toEqual(['A', 'Z', 'AA', 'ZZ', 'AAA', 'XFD'])
    expect(['a', 'Z', 'aa', 'XFD', 'XFE', 'ABCD', 'A1', ''].map(colNumber)).toEqual([1, 26, 27, MAX_COLS, null, null, null, null])
  })

  it('a row or column off the sheet is an InputError', () => {
    expect(() => cellAddress(0, 1)).toThrow(InputError)
    expect(() => cellAddress(1, MAX_COLS + 1)).toThrow(/col: expected an integer from 1 to 16384, got number 16385/)
    expect(() => colLetters(1.5)).toThrow(InputError)
    expect(() => parseCell(12 as unknown as string)).toThrow(/cell: expected a string, got number 12/)
  })
})

describe('ranges', () => {
  it('reads cells, whole columns and whole rows, corners put in order', () => {
    expect(parseRange('B2:D9')).toEqual({ kind: 'cells', top: 2, left: 2, bottom: 9, right: 4 })
    expect(parseRange('D9:B2')).toEqual({ kind: 'cells', top: 2, left: 2, bottom: 9, right: 4 })
    expect(parseRange('B9:D2')).toEqual({ kind: 'cells', top: 2, left: 2, bottom: 9, right: 4 })
    expect(parseRange('$b$2')).toEqual({ kind: 'cells', top: 2, left: 2, bottom: 2, right: 2 })
    expect(parseRange('$D:b')).toEqual({ kind: 'columns', top: 1, left: 2, bottom: MAX_ROWS, right: 4 })
    expect(parseRange('9:$2')).toEqual({ kind: 'rows', top: 2, left: 1, bottom: 9, right: MAX_COLS })
  })

  it('null for anything else', () => {
    for (const t of ['B2:', ':B2', 'B2:D9:E1', 'B:2', '2:B', 'B', '2', 'A1:XFE1', '0:3', 'A$:B', 'A$B1', "'S'!A1", 'B2 ']) expect(parseRange(t), t).toBeNull()
  })

  it('writes a rectangle as Excel writes it', () => {
    expect(rangeAddress({ top: 2, left: 2, bottom: 2, right: 2 })).toBe('B2')
    expect(rangeAddress({ top: 9, left: 4, bottom: 2, right: 2 })).toBe('B2:D9')
    expect(rangeAddress({ top: 1, left: 2, bottom: MAX_ROWS, right: 4 })).toBe('B:D')
    expect(rangeAddress({ top: 2, left: 1, bottom: 9, right: MAX_COLS })).toBe('2:9')
    expect(rangeAddress({ top: 1, left: 1, bottom: MAX_ROWS, right: MAX_COLS })).toBe('1:1048576')
    for (const t of ['B2', 'B2:D9', 'B:D', '2:9', 'A:A', '5:5']) expect(rangeAddress(parseRange(t)!)).toBe(t)
    // a range over every row is the columns' range, over every column the rows', as Excel writes A1:A1048576 and A1:XFD1
    expect(rangeAddress(parseRange('A1:A1048576')!)).toBe('A:A')
    expect(rangeAddress(parseRange('A1:XFD1')!)).toBe('1:1')
    expect(() => rangeAddress({ top: 0, left: 1, bottom: 1, right: 1 })).toThrow(/rect.top: expected an integer from 1 to 1048576/)
  })

  it('rectangles meet and contain', () => {
    const b2d9 = parseRange('B2:D9')!
    expect(rectsMeet(b2d9, parseRange('D9:F12')!)).toBe(true)
    expect(rectsMeet(b2d9, parseRange('E1:F12')!)).toBe(false)
    expect(rectsMeet(b2d9, parseRange('C:C')!)).toBe(true)
    expect(rectContains(b2d9, parseRange('C3:D9')!)).toBe(true)
    expect(rectContains(b2d9, parseRange('C3:E9')!)).toBe(false)
    expect(rectContains(parseRange('2:9')!, b2d9)).toBe(true)
  })
})

describe('layout lines', () => {
  it('reads the keys a book takes and writes them', () => {
    expect(parseLines('C', 'columns')).toEqual({ first: 3, last: 3 })
    expect(parseLines('B:D', 'columns')).toEqual({ first: 2, last: 4 })
    expect(parseLines('5', 'rows')).toEqual({ first: 5, last: 5 })
    expect(parseLines('5:9', 'rows')).toEqual({ first: 5, last: 9 })
    expect(linesKey(4, 2, 'columns')).toBe('B:D')
    expect(linesKey(3, 3, 'columns')).toBe('C')
    expect(linesKey(9, 5, 'rows')).toBe('5:9')
    expect(linesKey(5, 5, 'rows')).toBe('5')
  })

  it('null for a key a book refuses, as the book check does', () => {
    const keys: [string, 'columns' | 'rows'][] = [['C:C', 'columns'], ['D:B', 'columns'], ['c', 'columns'], ['XFE', 'columns'], ['5', 'columns'], ['5:5', 'rows'], ['9:5', 'rows'], ['0', 'rows'], ['1048577', 'rows'], ['C', 'rows']]
    for (const [key, axis] of keys) {
      expect(parseLines(key, axis), key).toBeNull()
      const book = { sheets: [{ name: 'S', cells: {}, [axis]: { [key]: { hidden: true } } }] }
      expect(() => checkBook(book), key).toThrow(InputError)
    }
    expect(() => parseLines('C', 'cols' as 'columns')).toThrow(/axis: expected "columns" or "rows"/)
  })
})

describe('names', () => {
  it('compares names as the engine does, in any case by Unicode upper case', () => {
    for (const n of ['Sales', 'straße', 'Ставка', 'ǅ', 'ﬀ']) expect(nameKey(n)).toBe(caseKey(n))
    expect(sameName('Sales', 'SALES')).toBe(true)
    expect(sameName('straße', 'STRASSE')).toBe(true)
    expect(sameName('Ставка', 'СТАВКА')).toBe(true)
    expect(sameName('Sales', 'Sale')).toBe(false)
  })

  it('a book refuses two sheets of one name by the same rule', () => {
    expect(() => checkBook({ sheets: [{ name: 'straße', cells: {} }, { name: 'STRASSE', cells: {} }] })).toThrow(/duplicate sheet name/)
  })

  it('quotes a sheet name where Excel does, the locale giving the booleans', () => {
    const q = (n: string) => quoteSheetIfNeeded(n)
    expect(['Sales', 'Лист1', '_x', 'a.b', 'XFE1', 'Region', 'R0', 'C1.x', 'A1B'].map(q)).toEqual(['Sales', 'Лист1', '_x', 'a.b', 'XFE1', 'Region', 'R0', 'C1.x', 'A1B'])
    expect(["O'Brien", 'Q1 data', '2024', 'A1', 'R1C1', 'R', 'TRUE', 'false', 'R2D2', 'R1C0'].map(q)).toEqual(["'O''Brien'", "'Q1 data'", "'2024'", "'A1'", "'R1C1'", "'R'", "'TRUE'", "'false'", "'R2D2'", "'R1C0'"])
    expect(quoteSheetIfNeeded('ИСТИНА', ruRU)).toBe("'ИСТИНА'")
    expect(quoteSheetIfNeeded('TRUE', ruRU)).toBe('TRUE')
  })

  it("takes the locale's object alone: a tag, which would need the list of locales, and an object without booleans are refused", () => {
    const refused = (got: string, hint = '') => new InputError('locale', `expected a locale object such as ruRU, with its booleans, got ${got}${hint}`)
    expect(() => quoteSheetIfNeeded('ИСТИНА', 'ru-RU' as never)).toThrow(refused('the string "ru-RU"', " (it takes no tag, so that the a1 subpath carries no list of locales)"))
    expect(() => quoteSheetIfNeeded('TRUE', null as never)).toThrow(refused('null'))
    expect(() => quoteSheetIfNeeded('TRUE', {} as never)).toThrow(refused('an object'))
    expect(() => quoteSheetIfNeeded('TRUE', { booleans: { TRUE: 'ИСТИНА' } } as never)).toThrow(refused('an object'))
    expect(quoteSheetIfNeeded('ЛОЖЬ', { booleans: ruRU.booleans })).toBe("'ЛОЖЬ'")
    expect(quoteSheetIfNeeded('TRUE', undefined)).toBe("'TRUE'")
  })
})

describe('the root exports the same functions', () => {
  it('one function, two entry points', () => {
    for (const [k, v] of Object.entries(a1)) expect((root as Record<string, unknown>)[k], k).toBe(v)
  })
})

describe('mergedArea: the rectangle a merge takes in', () => {
  const R = (top: number, left: number, bottom: number, right: number) => ({ top, left, bottom, right })
  it('grown by each merged area it meets, again for those it then meets, as mergeCells grows it', () => {
    expect(mergedArea(R(2, 2, 2, 2), [R(1, 1, 2, 2), R(2, 3, 5, 3)])).toEqual(R(1, 1, 2, 2))
    expect(mergedArea(R(2, 2, 2, 3), [R(1, 1, 2, 2), R(2, 3, 5, 3), R(5, 4, 6, 4)])).toEqual(R(1, 1, 5, 3))
    expect(mergedArea(R(2, 2, 2, 2), [R(5, 5, 6, 6)])).toEqual(R(2, 2, 2, 2))
    const book = { sheets: [{ name: 'S', cells: { A1: { value: 1 } }, merges: ['A1:B2', 'C2:C5', 'D5:D6'] }] }
    const r = editStructure(book, { op: 'mergeCells', sheet: 'S', range: 'B2:C2' })
    expect(r.ok && r.book.sheets[0].merges).toEqual(['D5:D6', rangeAddress(mergedArea(R(2, 2, 2, 3), [R(1, 1, 2, 2), R(2, 3, 5, 3), R(5, 4, 6, 4)]))])
  })
  it('the input checked', () => {
    expect(() => mergedArea(R(2, 2, 2, 2), [R(0, 1, 1, 1)])).toThrow('merges[0].top')
    expect(() => mergedArea(R(2, 2, 2, 2), {} as never)).toThrow('merges')
  })
})
