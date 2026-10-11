/**
 * Book settings through xlsx: the 1904 date system (`workbookPr date1904`), iterative calculation (`calcPr iterate`,
 * its count and change) and built-in number formats by their number, written and read back, the book computing as the
 * one written. Excel's answers are checked in `test/xlsx-excel.test.ts`.
 */
import { describe, expect, it } from 'vitest'
import { storedValues } from './xlsx-fixtures.js'
import { compute, type BookInput } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import './locales.js'

const text = async (data: Uint8Array, name: string) => (await openArchive(data).text(name))!
const clock = { now: '2026-10-07T12:00:00Z', seed: 7 }
const values = (book: BookInput) => compute(book, { clock }).sheets.map((s) => Object.fromEntries(Object.entries(s.cells).map(([a, c]) => [a, [c.value, c.text]])))

describe('xlsx settings: the 1904 date system', () => {
  const book: BookInput = {
    dateSystem: 1904,
    sheets: [{ name: 'S', cells: { A1: { value: { date: '2025-01-31T18:00:00' }, numberFormat: 'yyyy-mm-dd h:mm' }, A2: { value: -1, numberFormat: 14 }, B1: { formula: 'YEAR(A1)' }, B2: { formula: 'TEXT(A2,"yyyy-mm-dd")' }, B3: { formula: 'DATE(1904,1,1)' } } }],
  }

  it('is written as the workbook says it, dates as serial numbers of 1904, and read back; the book computes as the one written', async () => {
    const { data } = await writeXlsx({ book, result: compute(book, { clock }) })
    expect(await text(data, 'xl/workbook.xml')).toContain('<workbookPr date1904="1"/><bookViews>')
    expect(await text(data, 'xl/worksheets/sheet1.xml')).toContain('<c r="A1" s="1"><v>44226.75</v></c>')
    const r = await readXlsx(data)
    expect(r.book.dateSystem).toBe(1904)
    expect(r.book.sheets[0].cells.A1).toEqual({ value: 44226.75, numberFormat: 'yyyy-mm-dd h:mm' })
    expect(storedValues(r)).toEqual({ B1: 2025, B2: '-1904-01-02', B3: 0 })
    expect(values(r.book)).toEqual(values(book))
    expect(r.issues).toEqual([])
  })

  it('a book of the 1900 system has no workbookPr', async () => {
    const { data } = await writeXlsx({ book: { sheets: [{ name: 'S', cells: { A1: { value: { date: '2025-01-31' } } } }] } })
    expect(await text(data, 'xl/workbook.xml')).not.toContain('workbookPr')
    expect((await readXlsx(data)).book).toEqual({ sheets: [{ name: 'S', cells: { A1: { value: 45688 } } }] })
  })
})

describe('xlsx settings: iterative calculation', () => {
  const cells = { A1: { formula: 'MIN(B1+1,5)' }, B1: { formula: 'A1' } }

  it('the count and the change come back as written; without them, Excel defaults', async () => {
    for (const iteration of [{ maxIterations: 7, maxChange: 0.5 }, { maxChange: 1e-7 }, {}]) {
      const book: BookInput = { iteration, sheets: [{ name: 'S', cells }] }
      const r = await readXlsx((await writeXlsx({ book, result: compute(book, { clock }) })).data)
      expect(r.book).toEqual(book)
      expect(storedValues(r)).toEqual({ A1: 5, B1: 5 })
    }
    const { data } = await writeXlsx({ book: { iteration: { maxIterations: 7, maxChange: 0.5 }, sheets: [{ name: 'S', cells }] } })
    expect(await text(data, 'xl/workbook.xml')).toContain('<calcPr fullCalcOnLoad="1" iterate="1" iterateCount="7" iterateDelta="0.5"/>')
  })

  it('a book without formulas keeps its setting; a book without the setting turns it off, as Excel keeps the one it has without calcPr', async () => {
    const { data } = await writeXlsx({ book: { iteration: {}, sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] } })
    expect(await text(data, 'xl/workbook.xml')).toContain('<calcPr iterate="1"/>')
    expect(await text((await writeXlsx({ book: { sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] } })).data, 'xl/workbook.xml')).toContain('<calcPr/>')
    const plain = await writeXlsx({ book: { sheets: [{ name: 'S', cells }] } })
    expect(await text(plain.data, 'xl/workbook.xml')).toContain('<calcPr fullCalcOnLoad="1"/>')
    expect((await readXlsx(plain.data)).book.iteration).toBeUndefined()
  })
})

describe('xlsx settings: built-in number formats by number', () => {
  const book = (locale?: string): BookInput => ({
    ...(locale ? { locale } : {}),
    sheets: [{ name: 'S', cells: { A1: { value: 45688.75, numberFormat: 14 }, A2: { value: -1234.5, numberFormat: 38 }, A3: { value: 0.5, numberFormat: 9 }, A4: { value: 45688.75, numberFormat: 22 } }, rows: { 6: { numberFormat: 44 } }, columns: { A: { width: 12 }, C: { numberFormat: 20 } } }],
  })

  it('are written by their number and read back as numbers; each locale shows them its own way', async () => {
    for (const locale of [undefined, 'ru-RU']) {
      const { data } = await writeXlsx({ book: book(locale) })
      const styles = await text(data, 'xl/styles.xml')
      expect(styles).not.toContain('<numFmts')
      for (const id of [14, 38, 9, 22, 44, 20]) expect(styles).toContain(`<xf numFmtId="${id}" `)
      const r = await readXlsx(data)
      expect(r.book).toEqual(book(locale))
      expect(values(r.book)).toEqual(values(book(locale)))
    }
    expect(compute(book('ru-RU')).sheets[0].cells.A1.text).toBe('31.01.2025')
  })

  it('a code that a locale shows otherwise than its built-in format is written as a code of its own', async () => {
    const { data } = await writeXlsx({ book: { sheets: [{ name: 'S', cells: { A1: { value: -5, numberFormat: '#,##0_);(#,##0)' }, A2: { value: 1, numberFormat: 'm/d/yy' }, A3: { value: 1, numberFormat: '0.00' } } }] } })
    const styles = await text(data, 'xl/styles.xml')
    expect(styles).toContain('<numFmt numFmtId="164" formatCode="#,##0_);(#,##0)"/><numFmt numFmtId="165" formatCode="m/d/yy"/>')
    expect((await readXlsx(data)).book.sheets[0].cells).toEqual({ A1: { value: -5, numberFormat: '#,##0_);(#,##0)' }, A2: { value: 1, numberFormat: 'm/d/yy' }, A3: { value: 1, numberFormat: 2 } })
  })
})
