/**
 * Protection, a sheet's and its cells', for Excel (`run.ts protect <set>`). A case's file is a book as `writeXlsx` writes
 * it, changed by hand (`raw`) as another writer leaves it, and with `kept` the file `writeKept` writes over that one
 * after a session's edits. Excel opens it, takes the case's steps through its object model (a copy, rows inserted,
 * formats cleared, cells locked or unlocked, a sheet protected as the Protect Sheet dialog does or unprotected), reads
 * every sheet's protection and the listed cells' locked and formula hidden, and saves the book. The engine's answer is
 * the same steps taken by a session over the book the reader reads from the file.
 */
import type { BookInput, CellEdit, LayoutEdit, RangeRef, SheetProtection, StructureOp } from '../../../src/contract.js'
import { compute, createSession } from '../../../src/index.js'
import { SHEET_PROTECTION } from '../../../src/input/protection.js'
import { sheetLines } from '../../../src/layout/lines.js'
import { parseAddress } from '../../../src/model.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'
import { openKept, writeKept } from '../../../src/xlsx-keep.js'
import { cellsOf } from '../cond/cases.js'
import { rawFile, type RawEdit } from '../xlsx/sent.js'

/** What a protected sheet allows, of those the object model sets and reads (`Protection.Allow…`), in the script's order. */
export const ALLOWS = ['formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks', 'deleteColumns', 'deleteRows', 'sort', 'autoFilter', 'pivotTables'] as const
export type Allow = (typeof ALLOWS)[number]

/** A step Excel takes through its object model, and a session as a host would. */
export type ProtectStep =
  /** A copy of a range to a target: Range.Copy with a destination without `paste`, else Paste Special of that kind. */
  | { copy: RangeRef; to: RangeRef; paste?: 'all' | 'formats' | 'values' }
  /** Whole rows inserted above the rows of `range` (`3:3`). */
  | { insertRows: RangeRef }
  /** Excel's Clear Formats. */
  | { clearFormats: RangeRef }
  /** A range's locked and formula hidden set. */
  | { lock: RangeRef; locked: boolean; hidden: boolean }
  /** A sheet protected as the Protect Sheet dialog protects it (objects and scenarios too), with what it allows. */
  | { protectSheet: string; allow: Allow[] }
  | { unprotectSheet: string }

/** A session's edits before `writeKept` writes over the case's file. */
export interface KeptEdits {
  cells?: CellEdit[]
  layout?: LayoutEdit[]
  op?: StructureOp
}

export interface ProtectCase {
  id: string
  /** The key of its book in the set's `books`. */
  book: string
  raw?: RawEdit[]
  kept?: KeptEdits
  steps?: ProtectStep[]
  /** The cells read: a sheet and a range of it. */
  read: [sheet: string, range: string][]
}

export interface ProtectCases {
  books: Record<string, BookInput>
  cases: ProtectCase[]
}

export interface ProtectSet {
  buildProtectCases(): ProtectCases
}

/** A sheet's protection as the object model reads it: protected contents, objects, scenarios, what selection is enabled and what it allows. */
export interface SheetGuard {
  contents: boolean
  objects: boolean
  scenarios: boolean
  selection: string
  allow: Record<Allow, boolean>
}

export interface CellGuard {
  locked: boolean
  hidden: boolean
}

export interface ProtectAnswer {
  /** Excel's error when it refused a step. */
  error?: string
  sheets: Record<string, SheetGuard>
  /** The read cells by sheet and address. */
  cells: Record<string, Record<string, CellGuard>>
}

/** A step as the script takes it. */
export function stepArgs(s: ProtectStep): unknown[] {
  if ('copy' in s) return ['copy', s.copy.sheet, s.copy.range, s.to.sheet, s.to.range, s.paste ?? '']
  if ('insertRows' in s) return ['insertRows', s.insertRows.sheet, s.insertRows.range]
  if ('clearFormats' in s) return ['clearFormats', s.clearFormats.sheet, s.clearFormats.range]
  if ('lock' in s) return ['locked', s.lock.sheet, s.lock.range, s.locked, s.hidden]
  if ('protectSheet' in s) return ['protect', s.protectSheet, s.allow]
  return ['unprotect', s.unprotectSheet]
}

/** The file of a case: its book as `writeXlsx` writes it, its changes by hand made, and with `kept` written over by `writeKept` after the edits. */
export async function caseFile(set: ProtectCases, c: ProtectCase): Promise<Uint8Array> {
  const b = set.books[c.book]
  const data = await rawFile((await writeXlsx({ book: b, result: compute(b) })).data, c.raw)
  if (!c.kept) return data
  const k = await openKept(data)
  const s = createSession(k.book, { cached: k.cached })
  if (c.kept.cells) s.setCells(c.kept.cells)
  if (c.kept.layout) s.setLayout(c.kept.layout)
  if (c.kept.op && !s.editStructure(c.kept.op).ok) throw new Error(`${c.id}: the kept operation is refused`)
  const out = await writeKept(data, s.getSnapshot())
  s.close()
  return out.data
}

/** The book after a case's steps, taken by a session over `book`. */
export function stepsBook(book: BookInput, steps: ProtectStep[] = []): BookInput {
  const s = createSession(book)
  for (const st of steps) {
    if ('copy' in st) {
      const r = s.editStructure({ op: 'copyRange', from: st.copy, to: st.to, ...(st.paste ? { paste: st.paste } : {}) })
      if (!r.ok) throw new Error(`copy refused: ${r.message}`)
    } else if ('insertRows' in st) {
      const r = s.editStructure({ op: 'insertRows', sheet: st.insertRows.sheet, row: Number(st.insertRows.range.split(':')[0]), count: rowCount(st.insertRows.range) })
      if (!r.ok) throw new Error(`insertRows refused: ${r.message}`)
    } else if ('clearFormats' in st) s.clearCells([st.clearFormats], 'formats')
    else if ('lock' in st) s.patchCells([{ ...st.lock, set: { style: { protection: { locked: st.locked, hidden: st.hidden } } } }])
    else if ('protectSheet' in st) s.setLayout([{ sheet: st.protectSheet, protection: Object.fromEntries(st.allow.map((k) => [k, true])) }])
    else s.setLayout([{ sheet: st.unprotectSheet, protection: null }])
  }
  const out = structuredClone(s.getBook())
  s.close()
  return out
}

const rowCount = (rows: string) => {
  const [a, b] = rows.split(':').map(Number)
  return b - a + 1
}

/** A sheet's protection as the object model reads it; an unprotected sheet's contents alone, which is all a test compares of it. */
export function guardOf(p: SheetProtection | undefined): SheetGuard | { contents: false } {
  if (!p) return { contents: false }
  const allowed = (k: keyof SheetProtection) => p[k] ?? SHEET_PROTECTION[k]
  return {
    contents: true,
    objects: !allowed('objects'),
    scenarios: !allowed('scenarios'),
    selection: !allowed('selectUnlockedCells') ? 'no selection' : allowed('selectLockedCells') ? 'no restrictions' : 'unlocked cells',
    allow: Object.fromEntries(ALLOWS.map((k) => [k, allowed(k)])) as Record<Allow, boolean>,
  }
}

/** The engine's answer: an unprotected sheet by its contents alone. */
export interface EngineAnswer {
  sheets: Record<string, SheetGuard | { contents: false }>
  cells: ProtectAnswer['cells']
}

/** The engine's answer of a book: every sheet's protection, and the read cells' locked and hidden by the style they show. */
export function answerOf(book: BookInput, read: ProtectCase['read']): EngineAnswer {
  const cells: ProtectAnswer['cells'] = {}
  for (const [sheet, range] of read) {
    const sh = book.sheets.find((x) => x.name === sheet)!
    const at = sheetLines(sh)
    const out = (cells[sheet] ??= {})
    for (const a of cellsOf(range)) {
      const { row, col } = parseAddress(a)
      const style = at.look(row, col).style
      const p = style === undefined ? undefined : book.styles?.[style]?.protection
      out[a] = { locked: p?.locked ?? true, hidden: p?.hidden ?? false }
    }
  }
  return { sheets: Object.fromEntries(book.sheets.map((s) => [s.name, guardOf(s.protection)])), cells }
}

/** The engine's answer to a case: its steps over the book the reader reads from the case's file. */
export async function engineCase(set: ProtectCases, c: ProtectCase): Promise<EngineAnswer> {
  return answerOf(stepsBook((await readXlsx(await caseFile(set, c))).book, c.steps), c.read)
}
