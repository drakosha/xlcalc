/**
 * The references of a formula to other workbooks (`scanExternalReferences`): by a file's name as the formula bar shows it, or
 * by a link's number as a file stores it, for a host that reads those workbooks before the formula is entered.
 */
import { describe, expect, it } from 'vitest'
import { createSession, type BookInput } from '../src/index.js'
import { scanExternalReferences } from '../src/text.js'

const linked: BookInput = { sheets: [{ name: 'S', cells: {} }], externalBooks: [{ target: 'Книга.xlsx', sheets: [{ name: 'Лист', cells: { A1: 1, B1: 2 } }] }] }

describe('scanExternalReferences', () => {
  it('gives the book, the sheet and the area of a reference by name, quoted or bare', () => {
    expect(scanExternalReferences("='[Отчёт 2025.xlsm]Продажи'!$I$8")).toEqual([
      { book: 'Отчёт 2025.xlsm', sheet: 'Продажи', kind: 'cells', top: 8, left: 9, bottom: 8, right: 9 },
    ])
    expect(scanExternalReferences('[Смета.xlsx]Итог!B2:C3+1')).toEqual([{ book: 'Смета.xlsx', sheet: 'Итог', kind: 'cells', top: 2, left: 2, bottom: 3, right: 3 }])
  })

  it('takes entire columns and rows', () => {
    expect(scanExternalReferences("'[a.xlsx]S'!H:H")[0]).toMatchObject({ book: 'a.xlsx', kind: 'columns', left: 8, right: 8 })
    expect(scanExternalReferences("'[a.xlsx]S'!6:6")[0]).toMatchObject({ book: 'a.xlsx', kind: 'rows', top: 6, bottom: 6 })
  })

  it("gives a link's number as a number", () => {
    expect(scanExternalReferences('[1]Лист!A1')).toEqual([{ book: 1, sheet: 'Лист', kind: 'cells', top: 1, left: 1, bottom: 1, right: 1 }])
  })

  it('gives a defined name of another book with its name', () => {
    expect(scanExternalReferences('[Книга.xlsx]!Ставка*2')).toEqual([{ book: 'Книга.xlsx', name: 'Ставка' }])
  })

  it('skips strings, tables and references of the book itself', () => {
    for (const f of ['A1+Лист!B2', 'Таблица1[Сумма]', '[@Сумма]*2', 'HYPERLINK("[B.xlsx]S!A1")']) expect(scanExternalReferences(f)).toEqual([])
  })

  it('finds each of several references, inside functions too', () => {
    expect(scanExternalReferences("SUM('[a.xlsx]S'!A1:A5,[1]T!B2)").map((r) => r.book)).toEqual(['a.xlsx', 1])
  })

  it('gives nothing for text that does not read', () => {
    expect(scanExternalReferences('=[a.xlsx]S!A1+)')).toEqual([])
    // no sheet's `!` after the book: entering it is refused too
    expect(scanExternalReferences('[Книга.xlsx]Имя')).toEqual([])
    expect(createSession(linked).setCellText('S', 'A1', '=[Книга.xlsx]Имя')).toMatchObject({ ok: false, code: 'syntax' })
  })

  it("keeps a path and a quote in the book's name", () => {
    expect(scanExternalReferences("'C:\\путь\\[Книга.xlsx]Лист'!A1")).toEqual([{ book: 'C:\\путь\\Книга.xlsx', sheet: 'Лист', kind: 'cells', top: 1, left: 1, bottom: 1, right: 1 }])
    expect(scanExternalReferences("'[a''b.xlsx]S'!A1")[0]).toMatchObject({ book: "a'b.xlsx" })
  })

  it("gives a sheet's defined name with its sheet", () => {
    expect(scanExternalReferences("'[Книга.xlsx]Лист'!Имя")).toEqual([{ book: 'Книга.xlsx', name: 'Имя', sheet: 'Лист' }])
  })

  it('gives no area for a reference the range operator joins to something else, and the others of the formula', () => {
    for (const f of ['[Книга.xlsx]Лист!A1:B', "'[Книга.xlsx]Лист'!A1:B", '[1]Лист!A1:B', '[Книга.xlsx]Лист!A1:[Книга.xlsx]Лист!B2', 'B2:[Книга.xlsx]Лист!A1', '([Книга.xlsx]Лист!A1):B2']) expect(scanExternalReferences(f), f).toEqual([])
    expect(scanExternalReferences('[a.xlsx]S!A1:B+[c.xlsx]T!C1').map((r) => r.book)).toEqual(['c.xlsx'])
    // an area in the text, an argument of what the operator joins
    expect(scanExternalReferences('[Книга.xlsx]Лист!A1:B5')[0]).toMatchObject({ top: 1, left: 1, bottom: 5, right: 2 })
    expect(scanExternalReferences('INDEX([a.xlsx]S!A1:A5,2):B2')[0]).toMatchObject({ book: 'a.xlsx', bottom: 5 })
    // entering one over a book with no link is refused; over one with a link, it computes as Excel's operator does
    const s = createSession(linked)
    expect(s.setCellText('S', 'A1', '=[Нет.xlsx]Лист!A1:B')).toMatchObject({ ok: false, code: 'external_ref' })
    expect(s.setCellText('S', 'A1', '=[Книга.xlsx]Лист!A1:B')).toMatchObject({ ok: true })
    expect(s.getCell('S', 'A1')).toMatchObject({ value: { error: 'NAME' } })
  })

  it("reads a path with brackets of its own as the entry does, a sheet of this book: no link and #REF!", () => {
    const f = "'D:\\[old]\\[Книга.xlsx]Лист'!A1"
    expect(scanExternalReferences(f)).toEqual([])
    const s = createSession(linked)
    expect(s.setCellText('S', 'A1', `=${f}`)).toMatchObject({ ok: true })
    expect(s.getCell('S', 'A1')).toMatchObject({ value: { error: 'REF' }, formula: f })
  })

  it("reads a file's name without brackets as a sheet's, which only a link of the book makes a link", () => {
    expect(scanExternalReferences('Книга.xlsx!Имя')).toEqual([])
    const s = createSession(linked)
    s.setCellText('S', 'A1', '=Книга.xlsx!Имя')
    expect(s.getCell('S', 'A1')).toMatchObject({ formula: '[1]!Имя' })
  })
})
