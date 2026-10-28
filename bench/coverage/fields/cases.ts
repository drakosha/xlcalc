/**
 * Fields of a workbook the model holds for a host, through Excel for Windows (`bench/windows/run.ts fields <set>`): a
 * set names its cases, each a workbook written by `writeXlsx`, or by `writeKept` over a file `writeXlsx` wrote after
 * layout edits in a session and the sheets' views a host set on the workbook it saves; Excel opens the file, reads
 * each sheet's tab colour, gridlines and headings and the calculation option through its object model, and saves it. `test/fields-win.test.ts` checks that Excel read what the workbook holds and that the
 * reader reads it back from the file Excel saved.
 */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { BookInput, LayoutEdit, SheetView } from '../../../src/contract.js'
import { withField } from '../../../src/looks/sheet.js'
import { createSession } from '../../../src/index.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { openKept, writeKept } from '../../../src/xlsx-keep.js'

const ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '../../..')

export interface FieldsCase {
  id: string
  book: BookInput
  /** Layout edits a session makes over the book read by `openKept`, and views a host sets by sheet (null removes one), written over its file by `writeKept`. */
  kept?: { layout?: LayoutEdit[]; views?: Record<string, SheetView | null>; calculation?: BookInput['calculation'] | null }
}

/** What Excel read of each sheet: its tab colour (`#RRGGBB`, null for none), whether it shows gridlines and headings. */
export type SheetFields = [name: string, tab: string | null, gridlines: boolean, headings: boolean]

/** What Excel read: the sheets' fields, and the calculation option (`automatic`, `manual`, `autoNoTable`), read since the set `book`. */
export interface FieldsAnswer {
  sheets: SheetFields[]
  calculation?: string
}

export interface FieldsSet {
  buildFieldsCases(): FieldsCase[]
}

export const fieldsFile = (set: string) => join(ROOT, 'test/excel/win/fields', `${set}.json`)
export const savedDir = (set: string) => join(ROOT, 'test/excel/win/fields', set)

/** A book with the views and the calculation option a host set on it. */
function hostSet(b: BookInput, kept: NonNullable<FieldsCase['kept']>): BookInput {
  const views = kept.views ?? {}
  const out = { ...b, sheets: b.sheets.map((s) => (s.name in views ? withField(s, 'view', views[s.name] ?? undefined) : s)) }
  return kept.calculation === undefined ? out : withField(out, 'calculation', kept.calculation ?? undefined)
}

/** The book a case's file holds: its book, after its layout edits and views for a kept case. */
export function caseBook(c: FieldsCase): BookInput {
  if (!c.kept) return c.book
  const s = createSession(c.book)
  s.setLayout(c.kept.layout ?? [])
  const out = hostSet(s.getBook(), c.kept)
  s.close()
  return out
}

/** A case's file: written by `writeXlsx`, or by `writeKept` over that file after its layout edits. */
export async function fieldsFileOf(c: FieldsCase): Promise<Uint8Array> {
  const data = (await writeXlsx({ book: c.book })).data
  if (!c.kept) return data
  const read = await openKept(data)
  const s = createSession(read.book, { cached: read.cached })
  s.setLayout(c.kept.layout ?? [])
  const snap = s.getSnapshot()
  const out = (await writeKept(data, { ...snap, book: hostSet(snap.book, c.kept) })).data
  s.close()
  return out
}

/** The calculation option of a book, as Excel names it. */
export const bookCalculation = (b: BookInput): string => b.calculation ?? 'automatic'

/** What the sheets of a book hold of the fields, as Excel reads them. */
export const bookFields = (b: BookInput): SheetFields[] => b.sheets.map((s) => [s.name, s.tabColor ?? null, s.view?.gridlines !== false, s.view?.headings !== false])
