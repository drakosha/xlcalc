import { strict } from './strict-args.js'
import { writeKept } from './xlsx/keep/book.js'
import { openKept } from './xlsx/keep/open.js'
import type { WriteKeptOptions } from './xlsx/keep/types.js'
import { replaceKeptPrint } from './xlsx/print/keep.js'
import { printReader, type XlsxPrintRead } from './xlsx/print/read.js'
import type { SheetPrint } from './xlsx/print/types.js'
import type { ReadXlsxOptions, XlsxSnapshot, XlsxWritten } from './xlsx/types.js'

export type { XlsxPrintRead } from './xlsx/print/read.js'

async function openKeptWithPrint(data: Uint8Array | ArrayBuffer, opts?: ReadXlsxOptions): Promise<XlsxPrintRead> {
  const reader = printReader()
  return reader.result(await openKept(data, opts, reader.readWorksheet))
}

async function writeKeptWithPrint(
  original: Uint8Array | ArrayBuffer, snapshot: XlsxSnapshot, print: readonly SheetPrint[], opts?: WriteKeptOptions,
): Promise<XlsxWritten> {
  return writeKept(original, snapshot, opts, (view) => replaceKeptPrint(view, print))
}

/** Opens an xlsx file as `openKept` does, with its print settings by sheet for `writeKeptWithPrint`. */
const open = /* @__PURE__ */ strict(openKeptWithPrint)
/**
 * Writes a book over the file it was read from as `writeKept` does, replacing the print settings of each listed sheet;
 * a sheet listed by name only loses its settings. Drawings and the printer part stay as they were.
 */
const write = /* @__PURE__ */ strict(writeKeptWithPrint)
export { open as openKeptWithPrint, write as writeKeptWithPrint }
