import { describe, expect, it } from 'vitest'
import type { BookInput } from '../src/contract.js'
import { tableStyleFile, TABLE_STYLE_CASES } from '../bench/coverage/xlsx/table-style-names.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { libUrl } from './fixtures/target.js'

const { registerTableStyles } = await import(libUrl('table-styles.js')) as typeof import('../src/table-styles.js')
registerTableStyles()

const { readXlsx, writeXlsx } = await import(libUrl('xlsx.js')) as typeof import('../src/xlsx.js')
const { createSession } = await import(libUrl('index.js')) as typeof import('../src/index.js')
const { openKept, writeKept } = await import(libUrl('xlsx-keep.js')) as typeof import('../src/xlsx-keep.js')

describe('XLSX table style definitions and theme', () => {
  it('retains the definition and its name on the table, including an unnumbered Excel style name', async () => {
    const file = await tableStyleFile(TABLE_STYLE_CASES.find((c) => c.id === 'custom-light')!)
    const { book, issues } = await readXlsx(file)
    expect(issues).toEqual([])
    expect(book.sheets[0].tables?.[0].style).toEqual({ name: 'TableStyleLight', rowStripes: true })
    expect(book.tableStyles?.custom?.TableStyleLight).toEqual({
      pivot: false, table: true, elements: [{ type: 'headerRow', style: { fill: { color: '#FF0000' } } }],
    })
  })

  it('writes custom formats and palette, keeps them in a session and preserves them when editing existing XLSX', async () => {
    const tableStyles: NonNullable<BookInput['tableStyles']> = {
      theme: ['#FFFFFF', '#000000', '#EEEEEE', '#222222', '#126599', '#AABBCC', '#991122', '#117733', '#223344', '#AACC11', '#122113', '#765432'],
      custom: {
        'Custom & Blue': { table: true, pivot: false, elements: [
          { type: 'headerRow', style: { fill: { color: '#126599' }, font: { bold: true, color: '#FFFFFF' } } },
          { type: 'firstRowStripe', size: 2, style: { border: { bottom: { style: 'thin', color: '#991122' } } } },
        ] },
      },
    }
    const book: BookInput = {
      tableStyles,
      sheets: [{
        name: 'S', cells: { A1: { value: 'Qty' }, A2: { value: 1 } },
        tables: [{ name: 'Sales', range: 'A1:A3', columns: [{ name: 'Qty' }], style: { name: 'custom & blue', rowStripes: true } }],
      }],
    }
    const file = (await writeXlsx({ book })).data
    const contentTypes = await openArchive(file).text('[Content_Types].xml')
    expect(contentTypes).toContain('PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"')
    const opened = await openKept(file)
    expect(opened.issues).toEqual([])
    expect(opened.book.tableStyles).toEqual(tableStyles)
    expect(opened.book.sheets[0].tables?.[0].style).toEqual(book.sheets[0].tables?.[0].style)
    const session = createSession(opened.book, { cached: opened.cached })
    try {
      expect(session.getLayout().tableStyles).toEqual(tableStyles)
      session.setCells([{ sheet: 'S', cell: 'A2', input: { value: 3 } }])
      const style = { name: 'TableStyleLight3', firstColumn: true }
      expect(session.editStructure({ op: 'setTableStyle', table: 'Sales', style }).ok).toBe(true)
      const rewritten = await writeKept(file, session.getSnapshot())
      const reread = await readXlsx(rewritten.data)
      expect(reread.book.sheets[0].cells.A2.value).toBe(3)
      expect(reread.book.tableStyles).toEqual(tableStyles)
      expect(reread.book.sheets[0].tables?.[0].style).toEqual(style)
    } finally {
      session.close()
    }
  })
})
