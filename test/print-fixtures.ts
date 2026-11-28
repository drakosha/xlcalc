import type { BookInput } from '../src/contract.js'
import type { SheetPrint } from '../src/xlsx-print.js'
import type { XlsxRead, XlsxSnapshot } from '../src/xlsx.js'

export function worksheetPrintOf(book: BookInput): SheetPrint[] {
  return book.sheets.flatMap((sheet) => sheet.print ? [{ sheet: sheet.name, ...sheet.print }] : [])
}

export function printingSnapshot(snapshot: XlsxSnapshot, settings: readonly SheetPrint[]): XlsxSnapshot {
  return {
    ...snapshot,
    book: {
      ...snapshot.book,
      sheets: snapshot.book.sheets.map((sheet) => {
        const selected = settings.find((entry) => entry.sheet === sheet.name)
        if (!selected) return sheet
        const { sheet: _, ...print } = selected
        return { ...sheet, print: Object.keys(print).length ? print : undefined }
      }),
    },
  }
}

/** These pairs compare workbook data. Excel may save different printing defaults in each format. */
export function workbookData(read: XlsxRead): XlsxRead {
  const { printIssues: _, ...data } = read
  return { ...data, book: { ...data.book, sheets: data.book.sheets.map(({ print: _print, rowDefaults: _rows, ...sheet }) => sheet) } }
}
