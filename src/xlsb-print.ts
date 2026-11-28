import { strict } from './strict-args.js'
import { printReader, type XlsxPrintRead } from './xlsx/print/read.js'
import type { ReadXlsxOptions } from './xlsx/types.js'
import { printRecords } from './xlsb/print-markup.js'
import { readBinaryWorkbook } from './xlsb/read.js'
import { worksheet } from './xlsb/worksheet.js'

export type { XlsxPrintRead } from './xlsx/print/read.js'

async function readXlsbWithPrint(data: Uint8Array | ArrayBuffer, opts?: ReadXlsxOptions): Promise<XlsxPrintRead> {
  const reader = printReader()
  const read = await readBinaryWorkbook(data, opts, reader.readWorksheet, (part) => worksheet(part, printRecords(part.xml)))
  return reader.result(read)
}

/** Reads an XLSB workbook with its print settings by sheet, print issues and row defaults for print geometry. */
const read = /* @__PURE__ */ strict(readXlsbWithPrint)
export { read as readXlsbWithPrint }
