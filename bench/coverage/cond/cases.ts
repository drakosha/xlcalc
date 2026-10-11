/**
 * Conditional format cases for Excel (`run.ts cond <set>`, `bench/windows/run.ts cond <set>`). A case's book is written by
 * `writeXlsx`; Excel opens it, makes the case's rules through its object model (Windows), applies its structure
 * operation, computes the book, reads what the listed cells show (the display format: `Range.Text`, fill and font
 * colours, bold, italic) and saves the book. The answer is what Excel showed, and the file it saved goes beside it, for
 * the reader to read the rules Excel kept.
 */
import type { CellStyle, CellView, Clock, StructureOp } from '../../../src/contract.js'
import { createSession, editStructure, validateFormula } from '../../../src/index.js'
import { MAX_COLS, MAX_ROWS } from '../../../src/model.js'
import { cellCount, parseRange, rectCells } from '../../../src/structure/area.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { BUILTIN } from '../../../src/xlsx/styles/numfmt.js'
import '../locales.js'
import { takenBook, type StoredBook, type StoredRule } from '../stored.js'
import { rawFile, type RawEdit } from '../xlsx/sent.js'

/**
 * A rule Excel makes through its object model (Windows: `FormatConditions.Add` and its kin): `cellIs` with an operator
 * of `XlFormatConditionOperator` (1 between, 2 not between, 3 equal, 4 not equal, 5 greater, 6 less, 7 greater or equal,
 * 8 less or equal), `expression`, `text` with `XlContainsOperator` (0 contains, 1 does not contain, 2 begins with, 3 ends
 * with), `blanks`, `noBlanks`, `errors`, `noErrors`, `period` with `XlTimePeriods`, `top10`, `aboveAverage` with
 * `XlAboveBelow`, `unique`, `duplicate`, `colorScale` of 2 or 3 points, `dataBar`, `iconSet` with `XlIconSet`. A rule with a
 * style gets a red fill (#FFC7CE) and dark red bold font (#9C0006).
 */
export interface CondMake {
  sheet: string
  range: string
  add: 'cellIs' | 'expression' | 'text' | 'blanks' | 'noBlanks' | 'errors' | 'noErrors' | 'period' | 'top10' | 'aboveAverage' | 'unique' | 'duplicate' | 'colorScale' | 'dataBar' | 'iconSet'
  operator?: number
  f1?: string
  f2?: string
  text?: string
  /** top10: rank; colorScale: points; iconSet: the set; period: the period; aboveAverage: XlAboveBelow. */
  n?: number
  percent?: boolean
  bottom?: boolean
  stdDev?: number
}

export interface CondCase {
  id: string
  /** The key of the case's book in the set's `books`. */
  book: string
  /** Rules Excel makes (Windows only: Excel for Mac's AppleScript crashed making one). */
  make?: CondMake[]
  op?: StructureOp
  /** The cells whose look is read: a sheet and a range of it. */
  read: [sheet: string, range: string][]
  /** The file changed by hand after `writeXlsx`, as another writer leaves it: the engine holds what the reader reads from it. */
  raw?: RawEdit[]
}

export interface CondCases {
  books: Record<string, StoredBook>
  cases: CondCase[]
  /** The locale Excel runs the set in: en-US without one. */
  locale?: string
}

export interface CondSet {
  buildCondCases(): CondCases
  /** The book of a case whose rules Excel makes, with the rules it is expected to make. */
  expectedBook?(c: CondCase): StoredBook
}

/** What a cell shows: `Range.Text`, the fill (null: none), the font's colour, bold and italic, by its display format. */
export interface Look {
  text: string
  fill: string | null
  font: string
  bold: boolean
  italic: boolean
}

export interface CondAnswer {
  /** Excel's error when it refused the operation or a step. */
  error?: string
  /** The looks of the read cells, by sheet and address. */
  looks: Record<string, Record<string, Look>>
  /** The rules Excel lists on each sheet (Windows): type, applies to, formula 1, priority, stop if true. */
  rules?: Record<string, [type: number, appliesTo: string, formula: string, priority: number, stop: boolean][]>
}

const hexOf = (rgb: number[]) => `#${rgb.map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase()}`

/** A colour of COM (BGR as a number) as `#RRGGBB`. */
export const fromBgr = (n: number) => hexOf([n & 255, (n >> 8) & 255, (n >> 16) & 255])

/** A colour of AppleScript ({r, g, b}) as `#RRGGBB`. */
export const fromRgb = (rgb: number[]) => hexOf(rgb)

/** The cells of a range, row by row. */
export const cellsOf = (range: string): string[] => [...rectCells(parseRange(range)!)]

/** What a cell of a session window shows by its own style and its conditional formats' look, as Excel's display format gives it. */
export function viewLook(v: CellView | undefined, styles: CellStyle[] = []): Look {
  const own = v?.style !== undefined ? styles[v.style] : undefined
  const cf = v?.appearance?.style
  return {
    text: v?.text ?? '',
    fill: cf?.fill?.color ?? own?.fill?.color ?? null,
    font: cf?.font?.color ?? v?.color ?? own?.font?.color ?? '#000000',
    bold: cf?.font?.bold ?? own?.font?.bold ?? false,
    italic: cf?.font?.italic ?? own?.font?.italic ?? false,
  }
}

/** A rule's cells as text: each finite rectangle cell by cell, a whole row or column as written; sorted. */
function cellSet(ranges: string[]): string[] {
  const out = new Set<string>()
  for (const r of ranges) {
    const x = parseRange(r)!
    if (cellCount(x) > 10_000 || x.r2 === MAX_ROWS - 1 || x.c2 === MAX_COLS - 1) out.add(r)
    else for (const a of rectCells(x)) out.add(a)
  }
  return [...out].sort()
}

/** JSON with sorted object keys. */
const sorted = (v: unknown): unknown => (Array.isArray(v) ? v.map(sorted) : v && typeof v === 'object' ? Object.fromEntries(Object.entries(v).sort(([a], [b]) => (a < b ? -1 : 1)).map(([k, x]) => [k, sorted(x)])) : v)

/** A formula in its canonical spelling (sheets quoted), as written when it does not parse. */
function canonicalText(f: string): string {
  const r = validateFormula(f)
  return r.ok ? r.canonical : f
}

/**
 * A rule in a form that compares by meaning: its cells as a set (a file may order and split its ranges otherwise; its
 * formulas are written for the top-left of the rectangle around them either way), formulas in their canonical spelling,
 * a number format as its code without Excel's escapes of spaces (a dxf's built-in format comes back as its code), keys
 * sorted.
 */
export function ruleKey(rule: StoredRule): unknown {
  const { ranges: _, ...rest } = rule
  const canon = JSON.parse(JSON.stringify(rest), (k, v) =>
    (k === 'formula' || k === 'value') && typeof v === 'string' ? canonicalText(v) : k === 'formulas' ? (v as string[]).map(canonicalText) : k === 'numberFormat' || k === 'format' ? (typeof v === 'number' ? (v === 0 ? 'General' : BUILTIN[v]) : (v as string).replaceAll('\\ ', ' ')) : v,
  )
  return sorted({ cells: cellSet(rule.ranges).join(' '), ...canon })
}

/** A case's file: its book as `writeXlsx` writes it, changed by hand for a raw case. */
export async function condFile(set: CondCases, c: CondCase): Promise<Uint8Array> {
  return rawFile((await writeXlsx({ book: takenBook(set.books[c.book]) })).data, c.raw)
}

/** What the engine gives a case: its refusal, or each sheet's rules (`ruleKey`) and the looks of the read cells. */
export interface CondPrediction {
  refused?: string
  rules: Record<string, unknown[]>
  looks: Record<string, Record<string, Look>>
}

/** The engine's answer to a case at the moment `clock` (TODAY of time periods): the book after its operation, in a session. */
export function engineCase(set: CondCases, c: CondCase, clock: Clock, stored: StoredBook = set.books[c.book]): CondPrediction {
  const book = takenBook(stored)
  const r = c.op ? editStructure(book, c.op) : { ok: true as const, book }
  if (!r.ok) return { refused: r.code, rules: {}, looks: {} }
  const s = createSession(r.book, { clock })
  const looks: CondPrediction['looks'] = {}
  for (const [sheet, range] of c.read) {
    const views = s.getCells(sheet, range)
    looks[sheet] = Object.fromEntries(cellsOf(range).map((a) => [a, viewLook(views[a], r.book.styles)]))
  }
  // the rules as the session holds them, by the new names of their fields
  const rules = rulesOf(s.getBook())
  s.close()
  return { rules, looks }
}

/** Each sheet's rules as `ruleKey` gives them. */
export const rulesOf = (book: StoredBook): Record<string, unknown[]> => Object.fromEntries(book.sheets.map((s) => [s.name, (s.conditionalFormats ?? []).map(ruleKey)]))
