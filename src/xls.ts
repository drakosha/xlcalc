import { strict } from './strict-args.js'
import type { ReadXlsxOptions, XlsxRead } from './xlsx/types.js'
import * as xls from './xls/read.js'

/** Reads an XLS workbook, the values stored for its formulas, and issues in the requested language. */
async function readXls(data: Uint8Array | ArrayBuffer, opts?: ReadXlsxOptions): Promise<XlsxRead> {
  return (await xls.readXls(data, opts)).read
}

const read = /* @__PURE__ */ strict(readXls)
export { read as readXls }
