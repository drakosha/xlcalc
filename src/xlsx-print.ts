import { strict } from './strict-args.js'
import { readStyled } from './xlsx/read/book.js'
import { writeXlsxParts } from './xlsx/write/book.js'
import type { ReadXlsxOptions, WriteXlsxOptions, XlsxRead, XlsxSnapshot, XlsxWritten } from './xlsx/types.js'
import { printSheetGeometry as geometry } from './xlsx/print/geometry.js'
import { formatPrintHeaderFooter as formatHeaderFooter } from './xlsx/print/header-footer.js'
import { layoutPrintPages as layout } from './xlsx/print/layout.js'
import { printPaperSize as paperSize } from './xlsx/print/paper.js'
import { printReader } from './xlsx/print/read.js'
import type { PrintReadIssue, PrintSheetDefaults, SheetPrint } from './xlsx/print/types.js'
import { printParts } from './xlsx/print/write.js'

export type {
  PrintHeaderFooterContext, PrintHeaderFooterRun, PrintHeaderFooterSections, PrintHeaderFooterStyle,
} from './xlsx/print/header-footer.js'
export type { PrintGeometryOptions } from './xlsx/print/geometry.js'
export type {
  PrintAxisGeometry, PrintBreak, PrintDeviceGeometry, PrintHeaderFooter, PrintLayout, PrintLayoutIssue, PrintMargins, PrintOptions, PrintPage,
  PrintPageBand, PrintReadIssue, PrintRect,
  PrintRegion, PrintSetup, PrintSheetDefaults, PrintSheetGeometry, PrintSize, PrintSpan, PrintTarget, SheetPrint,
} from './xlsx/print/types.js'

export interface XlsxPrintRead extends XlsxRead {
  print: SheetPrint[]
  printIssues: PrintReadIssue[]
  sheetDefaults: PrintSheetDefaults[]
}

async function readXlsxWithPrint(data: Uint8Array | ArrayBuffer, opts?: ReadXlsxOptions): Promise<XlsxPrintRead> {
  const reader = printReader()
  const { read } = await readStyled(data, opts, reader.readWorksheet)
  return { ...read, print: reader.print, printIssues: reader.issues, sheetDefaults: reader.sheetDefaults }
}

async function writeXlsxWithPrint(snapshot: XlsxSnapshot, print: readonly SheetPrint[], opts?: WriteXlsxOptions): Promise<XlsxWritten> {
  return writeXlsxParts(snapshot, opts, (book) => printParts(book, print))
}

export const printSheetGeometry = /* @__PURE__ */ strict(geometry, 'printSheetGeometry', 2)
export const formatPrintHeaderFooter = /* @__PURE__ */ strict(formatHeaderFooter)
export const printPaperSize = /* @__PURE__ */ strict(paperSize)
export const layoutPrintPages = /* @__PURE__ */ strict(layout)
const read = /* @__PURE__ */ strict(readXlsxWithPrint)
const write = /* @__PURE__ */ strict(writeXlsxWithPrint)
export { read as readXlsxWithPrint, write as writeXlsxWithPrint }
