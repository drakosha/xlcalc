import { strict } from './strict-args.js'
import { printReader, type XlsxPrintRead } from './xlsx/print/read.js'
import type { PrintReadIssue } from './xlsx/print/types.js'
import type { ReadXlsxOptions } from './xlsx/types.js'
import { printMarkup } from './xls/print-markup.js'
import { readXls } from './xls/read.js'

export type { XlsxPrintRead } from './xlsx/print/read.js'

/** Reads an XLS workbook with its print settings by sheet, print issues and row defaults for print geometry. */
async function readXlsWithPrint(data: Uint8Array | ArrayBuffer, opts?: ReadXlsxOptions): Promise<XlsxPrintRead> {
  const reader = printReader()
  const issues: PrintReadIssue[] = []
  const { read } = await readXls(data, opts, {
    readPrint: (records, encoding, sheet) => printMarkup(records, encoding, (part, reason) => {
      if (!issues.some((issue) => issue.sheet === sheet && issue.part === part && issue.reason === reason)) issues.push({ sheet, part, reason })
    }),
    visitWorksheet: reader.readWorksheet,
  })
  const result = reader.result(read)
  result.printIssues.push(...issues)
  return result
}

const read = /* @__PURE__ */ strict(readXlsWithPrint)
export { read as readXlsWithPrint }
