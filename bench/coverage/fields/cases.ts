/**
 * Fields of a workbook the model holds for a host, through Excel for Windows (`bench/windows/run.ts fields <set>`): a
 * set names its cases, each a workbook written by `writeXlsx`, or by `writeKept` over a file `writeXlsx` wrote after
 * layout edits in a session; Excel opens the file, reads each sheet's tab colour, gridlines and headings through its
 * object model, and saves it. `test/fields-win.test.ts` checks that Excel read what the workbook holds and that the
 * reader reads it back from the file Excel saved.
 */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { BookInput, LayoutEdit } from '../../../src/contract.js'
import { createSession } from '../../../src/index.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { openKept, writeKept } from '../../../src/xlsx-keep.js'

const ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '../../..')

export interface FieldsCase {
  id: string
  book: BookInput
  /** Layout edits a session makes over the book read by `openKept`, written over its file by `writeKept`. */
  kept?: LayoutEdit[]
}

/** What Excel read of each sheet: its tab colour (`#RRGGBB`, null for none), whether it shows gridlines and headings. */
export type SheetFields = [name: string, tab: string | null, gridlines: boolean, headings: boolean]

export interface FieldsAnswer {
  sheets: SheetFields[]
}

export interface FieldsSet {
  buildFieldsCases(): FieldsCase[]
}

export const fieldsFile = (set: string) => join(ROOT, 'test/excel/win/fields', `${set}.json`)
export const savedDir = (set: string) => join(ROOT, 'test/excel/win/fields', set)

/** The book a case's file holds: its book, after its layout edits for a kept case. */
export function caseBook(c: FieldsCase): BookInput {
  if (!c.kept) return c.book
  const s = createSession(c.book)
  s.setLayout(c.kept)
  const out = s.getBook()
  s.close()
  return out
}

/** A case's file: written by `writeXlsx`, or by `writeKept` over that file after its layout edits. */
export async function fieldsFileOf(c: FieldsCase): Promise<Uint8Array> {
  const data = (await writeXlsx({ book: c.book })).data
  if (!c.kept) return data
  const read = await openKept(data)
  const s = createSession(read.book, { cached: read.cached })
  s.setLayout(c.kept)
  const out = (await writeKept(data, s.getSnapshot())).data
  s.close()
  return out
}

/** What the sheets of a book hold of the fields, as Excel reads them. */
export const bookFields = (b: BookInput): SheetFields[] => b.sheets.map((s) => [s.name, s.tabColor ?? null, s.view?.gridlines !== false, s.view?.headings !== false])
