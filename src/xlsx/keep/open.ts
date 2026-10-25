/**
 * `openKept`: an xlsx file read as `readXlsx` reads it, with the carriers `writeKept` needs among the book's names:
 * hidden names that hold the references of the parts the model does not have (`carriers.ts`).
 */
import { object } from '../../input/basic.js'
import { checkLimits, fileBytes, readXlsx } from '../read/book.js'
import type { ReadXlsxOptions, XlsxRead } from '../types.js'
import { carrierNames, carriersOf } from './carriers.js'
import { openSource } from './source.js'

/** Reads an xlsx file into a workbook as `readXlsx` does, its carriers added to the names for `writeKept`. */
export async function openKept(data: Uint8Array | ArrayBuffer, opts?: ReadXlsxOptions): Promise<XlsxRead> {
  const bytes = fileBytes(data)
  const read = await readXlsx(bytes, opts)
  const limits = opts === undefined ? undefined : object(opts, 'opts', "options such as { locale: 'ru-RU' }").limits
  const carriers = await carriersOf(await openSource(bytes, limits === undefined ? undefined : checkLimits(limits)), read.book)
  read.book.names = [...(read.book.names ?? []), ...carrierNames(carriers)]
  return read
}
