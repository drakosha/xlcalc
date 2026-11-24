import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { openKeptWithPrint, writeKeptWithPrint } from '../src/xlsx-print-keep.js'
import type { SheetPrint } from '../src/xlsx-print.js'
import { changedParts, partsOf, withParts } from './keep-files.js'

describe('xlsx/print/keep', () => {
  it('writes print settings in place and preserves the sheet’s drawing and printer relationship reference', async () => {
    const original = await withParts({
      sheets: [{ name: 'Data', tabColor: '#AA6600', outline: { summaryBelow: false }, cells: { A1: { value: 1 }, A2: { value: 2 } } }],
    }, {
      chart: { sheet: 0, series: ["'Data'!$A$1:$A$2"], anchor: [1, 0, 8, 3] },
      printArea: { sheet: 0, ref: "'Data'!$A$1:$A$5" },
      edit: { 'xl/worksheets/sheet1.xml': (xml) => xml.replace('<drawing', '<pageSetup orientation="portrait" r:id="rId77"/><drawing') },
    })
    const opened = await openKeptWithPrint(original)
    expect(opened.print[0]).toMatchObject({ sheet: 'Data', areas: ['A1:A5'] })
    const snapshot = { book: opened.book, result: opened.cached }
    const print: SheetPrint = {
      sheet: 'Data', areas: ['A1:C12'], repeatRows: { first: 1, last: 1 },
      setup: { fitToPage: true, fitToWidth: 1, fitToHeight: 0, orientation: 'landscape' },
      margins: { left: 0.6, right: 0.6, top: 0.7, bottom: 0.7, header: 0.3, footer: 0.3 },
      headerFooter: { oddFooter: '&RPage &P of &N' },
    }
    expect((await writeKeptWithPrint(original, snapshot, [])).data).toEqual(original)
    const { data } = await writeKeptWithPrint(original, snapshot, [print])
    expect(await changedParts(original, data)).toEqual(['xl/workbook.xml', 'xl/worksheets/sheet1.xml'])
    const sheetXml = (await partsOf(data)).get('xl/worksheets/sheet1.xml')!
    expect(sheetXml).toMatch(/<tabColor[^>]*\/><outlinePr[^>]*\/><pageSetUpPr[^>]*\/><\/sheetPr>/)
    expect(sheetXml).toMatch(/<pageMargins[^>]*\/><pageSetup r:id="rId77"[^>]*\/><headerFooter[\s\S]*<\/headerFooter><drawing/)
    const reread = await openKeptWithPrint(data)
    expect(reread.print).toEqual([print])
    expect(reread.printIssues).toContainEqual({ sheet: 'Data', part: 'pageSetup', reason: 'printerSettings' })

    const cleared = await writeKeptWithPrint(data, { book: reread.book, result: reread.cached }, [{ sheet: 'Data' }])
    const left = await openKeptWithPrint(cleared.data)
    expect(left.print).toEqual([])
    expect((await partsOf(cleared.data)).get('xl/worksheets/sheet1.xml')).toContain('<pageSetup r:id="rId77"/>')
  })

  it('scopes print titles by the worksheet’s actual position, including chart sheets', async () => {
    const original = readFileSync(new URL('./excel/keep/open/chartsheet-parts.xlsx', import.meta.url))
    const read = await openKeptWithPrint(original)
    expect(read.book.sheets[0].name).toBe('Data')
    const settings: SheetPrint = { sheet: 'Data', repeatRows: { first: 2, last: 3 }, areas: ['B2:D9'] }
    const { data } = await writeKeptWithPrint(original, { book: read.book, result: read.cached }, [settings])
    const parts = await partsOf(data)
    const workbook = parts.get('xl/workbook.xml')!
    expect(workbook).toMatch(/name="_xlnm.Print_Titles" localSheetId="1"/)
    expect(workbook.indexOf('<definedNames>')).toBeGreaterThan(workbook.indexOf('</sheets>'))
    expect(workbook.indexOf('<definedNames>')).toBeLessThan(workbook.indexOf('<calcPr'))
    expect((await openKeptWithPrint(data)).print.find((sheet) => sheet.sheet === 'Data')).toEqual(settings)
    const before = await partsOf(original)
    for (const [name, xml] of before) if (name.startsWith('xl/charts')) expect(parts.get(name)).toBe(xml)
  })
})
