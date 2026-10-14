/**
 * Excel's outline commands, for Excel for Windows (`bench/windows/run.ts outline <set> [ru|<tag>]`): a set names its
 * books, each the sheet `S` with its rows' and columns' outline levels, hidden lines, collapsed marks and summary
 * settings, and its cases, each a book and a sequence of commands through the object model: `Range.Group` and
 * `Ungroup` of whole rows or columns, `ShowDetail` set True or False on lines, lines hidden or shown by hand,
 * `Outline.ShowLevels`. Before the first command and after each one Excel reads every line of the window (its
 * `OutlineLevel`, `Hidden`, `Summary` and a summary line's `ShowDetail`) and saves a copy of the book, whose rows, columns
 * and outline settings are read from the file: the marks Excel keeps (`collapsed`) show only there.
 */
import type { BookInput, CellInput, ColumnInput, OutlineSettings, RowInput, StructureOp } from '../../../src/contract.js'
import { editStructure, InputError, outlineGroups, patchLines } from '../../../src/index.js'
import { colToIndex, indexToCol } from '../../../src/model.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { openArchive } from '../../../src/xlsx/zip/read.js'
import { groupCollector } from '../collect.js'

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

/** A command as the engine takes it: an outline operation, or lines hidden or shown by hand as `setLayout` hides them. */
export function engineCommand(c: OutlineCommand): StructureOp | { hidden: string; value: boolean } {
  if ('levels' in c) return { op: 'showLevels', sheet: 'S', ...(c.levels[0] ? { rows: c.levels[0] } : {}), ...(c.levels[1] ? { columns: c.levels[1] } : {}) }
  const [kind, text] = Object.entries(c)[0] as [string, string]
  if (kind === 'hidden' || kind === 'shown') return { hidden: text, value: kind === 'hidden' }
  const rows = /^\d/.test(text)
  if (kind === 'group' || kind === 'ungroup') return rows ? { op: kind, sheet: 'S', rows: text } : { op: kind, sheet: 'S', columns: text }
  // Show and Hide Detail take one line, as Excel's object model does (it refuses more)
  const op = kind === 'show' ? 'showDetail' : 'hideDetail'
  if (text.includes(':')) return { op, sheet: 'S', ...(rows ? { rows: text } : { columns: text }) } as unknown as StructureOp
  return rows ? { op, sheet: 'S', row: Number(text) } : { op, sheet: 'S', column: text }
}

/** What the engine holds of a book as Excel reads and saves it: the window's lines, each summary line's `ShowDetail` the group's collapsed state; the file `writeXlsx` writes. */
export async function engineView(book: BookInput, window: OutlineWindow | undefined, refused: boolean): Promise<OutlineView> {
  const sheet = book.sheets[0]
  const w = windowOf(window)
  const read = (map: Record<string, RowInput & ColumnInput> | undefined, axis: 'rows' | 'columns', [lo, hi]: [number, number], key: (i: number) => string) => {
    const detail = new Map(outlineGroups(sheet)[axis].flatMap((g) => (g.summary === null ? [] : [[g.summary - 1, !g.collapsed] as const])))
    const out: Record<string, LineView> = {}
    for (let i = lo; i <= hi; i++) {
      const e = Object.entries(map ?? {}).find(([k]) => {
        const x = linesOf(k)
        return i >= x.lo && i <= x.hi
      })?.[1]
      const level = e?.outlineLevel ?? 0
      const hidden = !!e && (e.hidden === true || (axis === 'rows' ? e.height : e.width) === 0)
      if (level || hidden || detail.has(i)) out[key(i)] = [level, hidden, detail.get(i) ?? null]
    }
    return out
  }
  const { data } = await writeXlsx({ book }, { values: 'omit' })
  const xml = (await openArchive(data).text('xl/worksheets/sheet1.xml'))!
  return {
    ...(refused ? { refused: true as const } : {}),
    rows: read(sheet.rows, 'rows', w.rows, (i) => String(i + 1)),
    columns: read(sheet.columns, 'columns', w.columns, indexToCol),
    saved: savedView(savedTags(xml), window),
  }
}

/** The engine's books of a case: the book Excel gets, then the book after each command, with whether the engine refused it (the book left as it was). */
export function engineSteps(b: OutlineBook, c: OutlineCase): { book: BookInput; refused: boolean }[] {
  let book = sentBook(b)
  const out = [{ book, refused: false }]
  for (const x of c.commands) {
    const op = engineCommand(x)
    let refused = false
    if ('hidden' in op) {
      const s = book.sheets[0]
      const key = /^\d/.test(op.hidden) ? 'rows' : 'columns'
      book = { ...book, sheets: [{ ...s, [key]: patchLines(s[key], op.hidden, { hidden: op.value ? true : null }) }] }
    } else {
      try {
        const r = editStructure(book, op)
        if (r.ok) book = r.book
        else refused = true
      } catch (e) {
        if (!(e instanceof InputError)) throw e
        refused = true
      }
    }
    out.push({ book, refused })
  }
  return out
}

/** The engine's views of a case: the book Excel gets, then after each command. */
export const engineViews = (b: OutlineBook, c: OutlineCase): Promise<OutlineView[]> => Promise.all(engineSteps(b, c).map((s) => engineView(s.book, c.window, s.refused)))

/** Cases of one group: ids `<group>-<n>`. */
export function collector() {
  const { out, add: put } = groupCollector<OutlineCase>()
  const add = (group: string, book: string, commands: OutlineCommand[], window?: OutlineWindow) => put(group, { book, commands, ...(window ? { window } : {}) })
  return { out, add }
}

/** The cases of a set by group: `g(group, book, ...commands)` puts a case of the commands on the book (`add` with a window too). */
export function outlineAdder() {
  const { out, add } = collector()
  return { out, add, g: (group: string, book: string, ...commands: OutlineCommand[]) => add(group, book, commands) }
}

/** Lines at an outline level, and hidden there. */
export const L = (outlineLevel: number): RowInput & ColumnInput => ({ outlineLevel })
export const H = (outlineLevel: number): RowInput => ({ outlineLevel, hidden: true })

/** The commands, by the lines they act on: Group, Ungroup, Show Detail, Hide Detail, and the level buttons. */
export const group = (lines: string): OutlineCommand => ({ group: lines })
export const ungroup = (lines: string): OutlineCommand => ({ ungroup: lines })
export const show = (lines: string): OutlineCommand => ({ show: lines })
export const hide = (lines: string): OutlineCommand => ({ hide: lines })
export const levels = (rows: number, columns = 0): OutlineCommand => ({ levels: [rows, columns] })
