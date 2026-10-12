/**
 * Excel's outline commands, for Excel for Windows (`bench/windows/run.ts outline <set> [ru|<tag>]`): a set names its
 * books, each the sheet `S` with its rows' and columns' outline levels, hidden lines, collapsed marks and summary
 * settings, and its cases, each a book and a sequence of commands through the object model: `Range.Group` and
 * `Ungroup` of whole rows or columns, `ShowDetail` set True or False on lines, lines hidden or shown by hand,
 * `Outline.ShowLevels`. Before the first command and after each one Excel reads every line of the window (its
 * `OutlineLevel`, `Hidden`, `Summary` and a summary line's `ShowDetail`) and saves a copy of the book, whose rows, columns
 * and outline settings are read from the file: the marks Excel keeps (`collapsed`) show only there.
 */
import type { BookInput, CellInput, ColumnInput, OutlineSettings, RowInput } from '../../../src/contract.js'
import { colToIndex, indexToCol } from '../../../src/model.js'

/** A book of a set: the sheet `S`. */
export interface OutlineBook {
  rows?: Record<string, RowInput>
  columns?: Record<string, ColumnInput>
  outline?: OutlineSettings
  cells?: Record<string, CellInput>
}

/**
 * A command through the object model, on lines `5:7`, `8`, `B:D` or `C` taken whole: `Group`, `Ungroup`, `ShowDetail`
 * True (`show`) or False (`hide`), `Hidden` set by hand (`hidden`, `shown`), and `Outline.ShowLevels(RowLevels,
 * ColumnLevels)` (`levels`, 0 leaving an axis alone, as the argument left out).
 */
export type OutlineCommand =
  | { group: string }
  | { ungroup: string }
  | { show: string }
  | { hide: string }
  | { hidden: string }
  | { shown: string }
  | { levels: [rows: number, columns: number] }

/** The lines read: rows `first:last` and columns `first:last`. */
export interface OutlineWindow {
  rows: string
  columns: string
}

export const WINDOW: OutlineWindow = { rows: '1:20', columns: 'A:L' }

export interface OutlineCase {
  id: string
  book: string
  commands: OutlineCommand[]
  /** The lines read, `WINDOW` by default. */
  window?: OutlineWindow
}

export interface OutlineSet {
  BOOKS: Record<string, OutlineBook>
  buildOutlineCases(): OutlineCase[]
}

/** A line of the window as COM reads it: row number or column letters, level (0 for none), hidden, a summary line's `ShowDetail` (null for another line). */
export type ComLine = [number | string, number, boolean, boolean | null]

/** What Excel read before the first command or after one: its error (the command refused), the window's grouped, hidden and summary lines, the tags of the copy it saved. */
export interface OutlineStep {
  error?: string
  rows: ComLine[]
  columns: ComLine[]
  /** The saved sheet's `<row>` tags with an outline level, hidden or collapsed, its `<col>`, `<outlinePr>` and `<sheetFormatPr>` tags. */
  saved: string[]
}

export interface OutlineAnswer {
  steps: OutlineStep[]
}

/** A line in one shape for Excel, the predictions and the engine: level, hidden, and a summary line's `ShowDetail` (null for another line). */
export type LineView = [level: number, hidden: boolean, detail: boolean | null]

/** A saved line: level, hidden, collapsed. */
export type SavedLine = [level: number, hidden: boolean, collapsed: boolean]

/** What a file holds of the outline over the window: lines with a level, hidden or collapsed; the summary settings; the deepest levels (`sheetFormatPr`). */
export interface SavedView {
  rows: Record<string, SavedLine>
  columns: Record<string, SavedLine>
  summary: { below: boolean; right: boolean }
  levels: [rows: number, columns: number]
}

/** A step in one shape: whether the command was refused, the lines read, the file saved. */
export interface OutlineView {
  refused?: true
  rows: Record<string, LineView>
  columns: Record<string, LineView>
  saved: SavedView
}

/** The book as Excel gets it. */
export function sentBook(b: OutlineBook, locale?: string): BookInput {
  return { ...(locale ? { locale } : {}), sheets: [{ name: 'S', cells: b.cells ?? { A1: { value: 'outline' } }, ...(b.rows ? { rows: b.rows } : {}), ...(b.columns ? { columns: b.columns } : {}), ...(b.outline ? { outline: b.outline } : {}) }] }
}

/** Lines `5:7`, `8`, `B:D` or `C` as zero-based indexes, the axis by their shape. */
export function linesOf(text: string): { rows: boolean; lo: number; hi: number } {
  const [a, b = a] = text.split(':')
  const rows = /^\d/.test(a)
  const index = (x: string) => (rows ? Number(x) - 1 : colToIndex(x))
  return { rows, lo: index(a), hi: index(b) }
}

/** Lines as COM's `Range` takes them whole: `8` as `8:8`, `C` as `C:C`. */
export const comLines = (text: string): string => (text.includes(':') ? text : `${text}:${text}`)

/** The window as zero-based spans. */
export function windowOf(w: OutlineWindow = WINDOW): { rows: [number, number]; columns: [number, number] } {
  const r = linesOf(w.rows)
  const c = linesOf(w.columns)
  return { rows: [r.lo, r.hi], columns: [c.lo, c.hi] }
}

const attr = (tag: string, name: string): string | undefined => new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1]
const on = (v: string | undefined) => v === '1' || v === 'true'

/** The saved tags (`OutlineStep.saved`) over the window: each line with a level, hidden or collapsed; the settings; the deepest levels. */
export function savedView(tags: string[], window?: OutlineWindow): SavedView {
  const w = windowOf(window)
  const out: SavedView = { rows: {}, columns: {}, summary: { below: true, right: true }, levels: [0, 0] }
  for (const t of tags) {
    const name = /^<(\w+)/.exec(t)![1]
    const line = (): SavedLine => [Number(attr(t, 'outlineLevel') ?? 0), on(attr(t, 'hidden')), on(attr(t, 'collapsed'))]
    if (name === 'row') {
      const r = Number(attr(t, 'r')) - 1
      const l = line()
      if (r >= w.rows[0] && r <= w.rows[1] && (l[0] || l[1] || l[2])) out.rows[String(r + 1)] = l
    } else if (name === 'col') {
      const l = line()
      if (!(l[0] || l[1] || l[2])) continue
      for (let c = Math.max(Number(attr(t, 'min')) - 1, w.columns[0]); c <= Math.min(Number(attr(t, 'max')) - 1, w.columns[1]); c++) out.columns[indexToCol(c)] = l
    } else if (name === 'outlinePr') {
      out.summary = { below: attr(t, 'summaryBelow') === undefined || on(attr(t, 'summaryBelow')), right: attr(t, 'summaryRight') === undefined || on(attr(t, 'summaryRight')) }
    } else if (name === 'sheetFormatPr') out.levels = [Number(attr(t, 'outlineLevelRow') ?? 0), Number(attr(t, 'outlineLevelCol') ?? 0)]
  }
  return out
}

/** The tags `savedView` reads, from a sheet's XML as `xl.ps1` takes them from the file Excel saved. */
export const savedTags = (xml: string): string[] => [...xml.matchAll(/<(row|col|outlinePr|sheetFormatPr)\b[^>]*>/g)].map((m) => m[0]).filter((t) => !t.startsWith('<row') || /\b(hidden|outlineLevel|collapsed)=/.test(t))

/** A step Excel read in the shape of `OutlineView`. */
export function excelView(s: OutlineStep, window?: OutlineWindow): OutlineView {
  const lines = (ls: ComLine[]) => Object.fromEntries(ls.map(([k, level, hidden, detail]): [string, LineView] => [String(k), [level, hidden, detail]]))
  return { ...(s.error ? { refused: true as const } : {}), rows: lines(s.rows), columns: lines(s.columns), saved: savedView(s.saved, window) }
}

/** Cases of one group: ids `<group>-<n>`. */
export function collector() {
  const out: OutlineCase[] = []
  const count = new Map<string, number>()
  const add = (group: string, book: string, commands: OutlineCommand[], window?: OutlineWindow) => {
    const n = (count.get(group) ?? 0) + 1
    count.set(group, n)
    out.push({ id: `${group}-${n}`, book, commands, ...(window ? { window } : {}) })
  }
  return { out, add }
}
