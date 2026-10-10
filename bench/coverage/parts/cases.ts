/**
 * Sheet parts cases for Excel (`run.ts parts <set>`): notes, hyperlinks and frozen panes. A case's book is written by
 * `writeXlsx` (with threads of comments put in by hand, `threaded.ts`); Excel opens it, adds what the case makes through
 * its object model (Excel's own notes, hyperlinks and frozen panes), applies the case's structure operation, reads every
 * sheet's parts and saves the book. The answer is what Excel read, and the file it saved goes beside it.
 */
import type { BookInput, CellLink, CellNote, FrozenPanes, StructureOp } from '../../../src/contract.js'

/** What Excel adds through its object model: `AddComment`, `Hyperlinks.Add`, `FreezePanes`. */
export type PartsMake =
  | { note: string; sheet: string; text: string }
  | { link: string; sheet: string; url?: string; location?: string; tooltip?: string }
  | { freeze: string; rows: number; columns: number }

/** A thread of comments in a book's file, as Excel 365 writes one: its sheet and cell, each comment's author and text. */
export interface ThreadSpec {
  sheet: string
  cell: string
  comments: { author: string; text: string }[]
}

export interface PartsCase {
  id: string
  /** The key of the case's book in the set's `books`. */
  book: string
  make?: PartsMake[]
  op?: StructureOp
}

export interface PartsCases {
  books: Record<string, BookInput>
  cases: PartsCase[]
  /** Threads of comments put into the files of books, by book. */
  threads?: Record<string, ThreadSpec[]>
}

export interface PartsSet {
  buildPartsCases(): PartsCases
}

/** A hyperlink as Excel's object model gives it: the cells it is on, Address, SubAddress, ScreenTip, TextToDisplay (empty when none). */
export interface ExcelLink {
  range: string
  address: string
  subAddress: string
  tip: string
  display: string
}

/** A thread of comments as Excel for Windows gives it: each comment's author and text, the first one's then the replies'. */
export type ExcelThread = { author: string; text: string }[]

/** Excel's window over a sheet: frozen or not, the rows above and the columns left of the split, the first row and column its top-left pane shows. */
export interface ExcelPanes {
  freeze: boolean
  splitRow: number
  splitColumn: number
  scrollRow: number
  scrollColumn: number
}

export interface PartsSheetAnswer {
  name: string
  /** Notes by address: author and text. */
  notes: Record<string, { author: string; text: string }>
  links: ExcelLink[]
  panes: ExcelPanes
  /** Threads of comments by address (Windows). */
  threads?: Record<string, ExcelThread>
}

export interface PartsAnswer {
  /** Excel's error when it refused the operation or a step. */
  error?: string
  sheets: PartsSheetAnswer[]
}

/** A sheet's parts in one shape for the engine, the reader and Excel: notes and links by cell, frozen panes. */
export interface SheetParts {
  notes: Record<string, CellNote>
  links: Record<string, CellLink>
  frozen?: FrozenPanes
}

/** The parts of a sheet of a book. */
export function bookParts(s: BookInput['sheets'][number]): SheetParts {
  const out: SheetParts = { notes: {}, links: {} }
  for (const [addr, c] of Object.entries(s.cells)) {
    if (c.note) out.notes[addr] = c.note.author ? { author: c.note.author, text: c.note.text } : { text: c.note.text }
    if (c.link) out.links[addr] = c.link
  }
  if (s.frozen) out.frozen = s.frozen
  return out
}
