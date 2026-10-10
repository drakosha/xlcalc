/**
 * Hidden sheets for Excel (`run.ts hidden <set>`, `bench/windows/run.ts hidden <set>`). A case's book is written by
 * `writeXlsx`, or changed by hand after it (`raw`: the active tab on a hidden sheet, every sheet hidden); Excel opens it,
 * notes its active sheet, sets the visibility of sheets through its object model (`Visible`), applies the case's sheet
 * operation, reads each sheet's visibility, the active sheet and the values of the read cells, and saves the book. The
 * answer is what Excel read; the file it saved goes beside it.
 */
import type { BookInput, CellValue, StructureOp } from '../../../src/contract.js'
import { compute, restructure } from '../../../src/index.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { editParts } from '../xlsx/sent.js'

/** A sheet's visibility as Excel's object model gives it (`XlSheetVisibility`): visible, hidden, very hidden. */
export type Visible = -1 | 0 | 2

/** What Excel sets through its object model: a sheet's `Visible`. */
export interface HiddenMake {
  sheet: string
  visible: Visible
}

export interface HiddenCase {
  id: string
  /** The key of the case's book in the set's `books`. */
  book: string
  /**
   * The file changed by hand after `writeXlsx`: `activeHidden` puts the active tab on the first sheet, hidden or not (as
   * a writer that ignores hidden sheets leaves it); `allHidden` hides every sheet, which the model does not take.
   */
  raw?: 'activeHidden' | 'allHidden'
  make?: HiddenMake[]
  /** A sheet operation: renameSheet, deleteSheet, insertSheet or moveSheet. */
  op?: StructureOp
}

export interface HiddenCases {
  books: Record<string, BookInput>
  cases: HiddenCase[]
  /** The cells whose values are read after each case, on a sheet no case renames or deletes. */
  read: { sheet: string; cells: string[] }
}

export interface HiddenSet {
  buildHiddenCases(): HiddenCases
}

/** What Excel read in a case's book. */
export interface HiddenAnswer {
  /** Excel's error when it refused a step (`make: …`, `op: …`). */
  error?: string
  /** The active sheet when Excel opened the file, and after the case's steps. */
  opened: string
  active: string
  /** Each sheet's name and visibility, in the book's order. */
  sheets: [string, Visible][]
  /** The read cells' values. */
  values: CellValue[]
}

/** A sheet's visibility as the object model gives it. */
export const visibleOf = (s: BookInput['sheets'][number]): Visible => (s.hidden === 'very' ? 2 : s.hidden ? 0 : -1)

/** The structure operation that sets a sheet's visibility as `make` does. */
export const makeOp = (m: HiddenMake): StructureOp => (m.visible === -1 ? { op: 'unhideSheet', sheet: m.sheet } : { op: 'hideSheet', sheet: m.sheet, ...(m.visible === 2 ? { very: true } : {}) })

/** The file of a case: the book as `writeXlsx` writes it, changed by hand for a raw case. */
export async function hiddenFile(c: HiddenCase, set: HiddenCases): Promise<Uint8Array> {
  const data = (await writeXlsx(set.books[c.book])).data
  if (!c.raw) return data
  if (c.raw === 'allHidden') return editParts(data, (name, x) => (name === 'xl/workbook.xml' ? x.replace(/<sheet name="([^"]*)" sheetId="(\d+)"(?: state="\w+")?/g, '<sheet name="$1" sheetId="$2" state="hidden"') : undefined))
  return editParts(data, (name, x) => {
    if (name === 'xl/workbook.xml') return x.replace(/<workbookView activeTab="\d+"\/>/, '<workbookView/>')
    if (name === 'xl/worksheets/sheet1.xml') return x.replace(' tabSelected="1"', '').replace('<sheetView workbookViewId="0"', '<sheetView workbookViewId="0" tabSelected="1"')
    return /^xl\/worksheets\/sheet\d+\.xml$/.test(name) ? x.replace(' tabSelected="1"', '') : undefined
  })
}

/** Each sheet's name and visibility. */
const sheetsOf = (book: BookInput): [string, Visible][] => book.sheets.map((s) => [s.name, visibleOf(s)])

/** What the engine gives for a case: its refusal if it refuses a step, each sheet's visibility and the read cells' values then. */
export function engineAnswer(set: HiddenCases, c: HiddenCase): { refused?: string; sheets: [string, Visible][]; values: CellValue[] } {
  let book = set.books[c.book]
  let refused: string | undefined
  for (const op of [...(c.make ?? []).map(makeOp), ...(c.op ? [c.op] : [])]) {
    const r = restructure(book, op)
    if (!r.ok) {
      refused = r.code
      break
    }
    book = r.book
  }
  const cells = compute(book, { clock: { now: '2026-10-10T12:00:00Z', seed: 1 } }).sheets.find((s) => s.name === set.read.sheet)!.cells
  return { ...(refused ? { refused } : {}), sheets: sheetsOf(book), values: set.read.cells.map((a) => cells[a]?.value ?? null) }
}

/**
 * Whether Excel refused a case's steps: its error, or a sheet deleted that is still there (Excel for Mac deletes no very
 * hidden sheet and not the last visible one, and answers nothing).
 */
export const excelRefused = (set: HiddenCases, c: HiddenCase, a: HiddenAnswer) =>
  !!a.error || (c.op?.op === 'deleteSheet' && JSON.stringify(a.sheets) === JSON.stringify(sheetsOf(set.books[c.book])))
