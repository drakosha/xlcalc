import { InputError } from '../input-error.js'
import { readStyledArchive, workbookInput } from '../xlsx/read/book.js'
import { relsOf } from '../xlsx/read/package.js'
import type { ReadXlsxOptions, XlsxRead } from '../xlsx/types.js'
import { openArchive } from '../xlsx/zip/read.js'
import { convertingArchive } from './archive.js'

/** Reads an XLSB workbook's cells, saved values and issues using the SpreadsheetML reader. */
export async function readXlsb(data: Uint8Array | ArrayBuffer, opts?: ReadXlsxOptions): Promise<XlsxRead> {
  const { bytes, language, limits } = workbookInput(data, opts)
  const budget = { spent: 0 }
  const source = openArchive(bytes, limits, budget)
  const root = await relsOf(source, '')
  const path = root.find((rel) => rel.kind === 'officeDocument' && !rel.external)?.target
  if (!path?.toLowerCase().endsWith('.bin')) throw new InputError('data', 'expected an .xlsb workbook')
  const unfinished = new Set<string>()
  const archive = await convertingArchive(source, path, limits ?? {}, budget, unfinished)
  const { read } = await readStyledArchive(archive, root, path, language)
  for (const part of unfinished) {
    const message = language?.tag === 'ru-RU' ? `Бинарная часть ${part} прочитана не полностью` : `some of the binary part ${part} was left out`
    read.issues.push({ code: 'part', message })
  }
  return read
}
