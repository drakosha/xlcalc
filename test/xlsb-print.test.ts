import { readFile } from 'node:fs/promises'
import { describe, expect, it } from 'vitest'
import { libUrl } from './fixtures/target.js'

const { readXlsb } = await import(libUrl('xlsb.js')) as typeof import('../src/xlsb.js')
const { readXlsbWithPrint } = await import(libUrl('xlsb-print.js')) as typeof import('../src/xlsb-print.js')
const { readXlsx } = await import(libUrl('xlsx.js')) as typeof import('../src/xlsx.js')
const { readXlsxWithPrint, writeXlsxWithPrint } = await import(libUrl('xlsx-print.js')) as typeof import('../src/xlsx-print.js')

/** Excel for Mac saved all three files from the same workbook with two configured worksheets. */
const fixture = (format: string) => readFile(new URL(`./excel/print-formats/print.${format}`, import.meta.url))

describe('printing from Excel files', () => {
  it('reads XLSB print areas, titles, page setup, margins, options, headers and footers, and manual breaks', async () => {
    const data = await fixture('xlsb')
    const { print, printIssues, sheetDefaults, ...workbook } = await readXlsbWithPrint(data)
    const xlsx = await fixture('xlsx').then(readXlsxWithPrint)
    expect(xlsx.print).toMatchObject([
      {
        sheet: "O'Brien Print",
        areas: ['B2:H16', 'B20:H30'],
        repeatRows: { first: 2, last: 3 },
        repeatColumns: { first: 2, last: 2 },
        setup: { paperSize: 9, orientation: 'landscape', scale: 82, pageOrder: 'overThenDown', firstPageNumber: 7, cellComments: 'atEnd', errors: 'dash' },
        margins: { left: 0.35, right: 0.45, top: 0.65, bottom: 0.85, header: 0.2, footer: 0.3 },
        options: { horizontalCentered: true, verticalCentered: true, headings: true, gridLines: true },
        headerFooter: { oddHeader: '&CОтчёт &P', evenHeader: '&LEven', firstFooter: '&CFirst footer' },
        rowBreaks: [{ before: 8, manual: true }],
        columnBreaks: [{ before: 5, manual: true }],
      },
      { sheet: 'Fit', setup: { autoPageBreaks: false, fitToPage: true, fitToWidth: 2, fitToHeight: 0 } },
    ])
    // The binary record also stores printer defaults such as 600 DPI that Excel omits from the XML.
    expect(print).toMatchObject(xlsx.print)
    expect(print[0].setup).toMatchObject({ horizontalDpi: 600, verticalDpi: 600 })
    expect(printIssues).toEqual([])
    expect(sheetDefaults).toStrictEqual(xlsx.sheetDefaults)
    expect(workbook).toStrictEqual(await readXlsb(data))

    const written = await writeXlsxWithPrint({ book: workbook.book, result: workbook.cached }, print, { values: 'preserve' })
    expect((await readXlsxWithPrint(written.data)).print).toStrictEqual(print)
  })

  it('reads the cells, cached formula value and print settings in an XLSM through the XLSX readers', async () => {
    const data = await fixture('xlsm')
    const xlsx = await fixture('xlsx')
    const read = await readXlsx(data)
    expect(read.book.sheets[0].cells.C3.formula).toBe('B2*2')
    expect(read.cached.sheets[0].cells.C3.value).toBe(42)
    expect(read).toStrictEqual(await readXlsx(xlsx))
    await expect(readXlsxWithPrint(data)).resolves.toStrictEqual(await readXlsxWithPrint(xlsx))
  })
})
