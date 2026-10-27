/**
 * Text with a line break typed into cells (`setCellText`, `setCellTexts`, `fillCellText`) or pasted (`pasteText`), as
 * Excel takes it (`test/excel/rows/`): the cell wraps its text, its style taking `alignment.wrap` (a style of the book's
 * table found or added); a formula typed over lines does not. Typed, the row of a cell that wraps fits it as AutoFit
 * does, unless its height was set by hand or it is hidden; a paste fits no row.
 */
import type { CellInput, CellStyle } from '../contract.js'
import type { EngineState } from '../engine.js'
import { fitHeights, stateRows } from '../layout/heights.js'
import { sheetLines } from '../layout/lines.js'
import { inner, MAX_ROWS } from '../model.js'
import { outCell } from '../result.js'
import { firstEnding, lines } from '../layout/spans.js'
import { StyleTable } from '../style-table.js'
import type { SetLayoutEdit } from './book.js'
import type { CheckedPatch } from './patch.js'
import type { Rect } from '../structure/area.js'
import { patchStyle } from './styles.js'
import type { Typed } from './texts.js'

/** Whether typed text breaks a line, which makes Excel wrap the cell: text that is no formula (`input`, what it gives the cell). */
export const breaksLine = (text: string, input: CellInput | null): boolean => input?.formula === undefined && /[\r\n]/.test(text)

/**
 * The typed cells marked `wrap` given the style that wraps (their own, else their row's or column's, with `alignment.wrap`)
 * in their inputs; the book's style table when it grew.
 */
export function wrapTyped(s: EngineState, cells: Typed[]): CellStyle[] | undefined {
  if (!cells.some((x) => x.wrap)) return undefined
  const table = new StyleTable(s.input.styles)
  for (const x of cells) {
    if (!x.wrap || !x.edit.input) continue
    const lines = sheetLines(s.input.sheets[x.i], s)
    const shown = x.edit.input.style ?? lines.row(x.r)?.style ?? lines.column(x.c)?.style
    if (table.at(shown).alignment?.wrap) continue
    x.edit.input = { ...x.edit.input, style: table.indexOf(patchStyle(table.at(shown), { alignment: { wrap: true } })) }
  }
  return table.table
}

/** The heights the rows of the typed cells that wrap take now, as AutoFit gives them, where a row's height is the default or one Excel fitted and it shows: row by row. */
export function typedHeights(s: EngineState, cells: Typed[]): SetLayoutEdit[] {
  const styles = s.input.styles ?? []
  const rows = new Map<number, Set<number>>()
  const linesOf = new Map<number, ReturnType<typeof sheetLines>>()
  for (const { i, r, c } of cells) {
    let lines = linesOf.get(i)
    if (!lines) linesOf.set(i, (lines = sheetLines(s.input.sheets[i], s)))
    const style = lines.look(r, c).style
    if (style === undefined || !styles[style]?.alignment?.wrap) continue
    inner(rows, i, () => new Set()).add(r)
  }
  return [...rows].flatMap(([i, set]) => fittedHeights(s, i, spansOf([...set]), false))
}

/**
 * The rows of rectangles of sheet `i` to fit, as runs: of one over every row (entire columns, the whole sheet), only the
 * rows with cells in its columns: whether Excel fits the empty ones is not measured.
 */
function rowSpans(s: EngineState, i: number, rects: Rect[]): [number, number][] {
  const tall = rects.filter((x) => x.r1 === 0 && x.r2 === MAX_ROWS - 1)
  const spans = rects.filter((x) => !tall.includes(x)).map((x): [number, number] => [x.r1, x.r2])
  if (tall.length) for (const [r, c] of stateRows(s, i, 0, MAX_ROWS - 1).cells) if (tall.some((x) => c >= x.c1 && c <= x.c2)) spans.push([r, r])
  return merged(spans)
}

/** The heights of the rows of patches that change a style (`fittedHeights`, in runs of rows): a font may make them taller or lower. */
export function patchedHeights(s: EngineState, patches: CheckedPatch[]): SetLayoutEdit[] {
  const targets = new Map<number, CheckedPatch[]>()
  for (const p of patches) if (p.look.style !== undefined) inner(targets, p.i, () => []).push(p)
  return [...targets].flatMap(([i, list]) => fittedHeights(s, i, rowSpans(s, i, list.map((p) => p.rect)), true))
}

/** The heights of the rows of `area` of sheet `i` after a copy laid its looks there, as after a patch of styles (the paste set `fit`). */
export const pastedHeights = (s: EngineState, i: number, area: Rect): SetLayoutEdit[] => fittedHeights(s, i, rowSpans(s, i, [area]), true)

/** Rows as runs of consecutive ones, in order. */
const spansOf = (rows: number[]): [number, number][] => merged(rows.map((r) => [r, r]))

/** Spans of rows in order, those that overlap or meet joined. */
function merged(spans: [number, number][]): [number, number][] {
  const out: [number, number][] = []
  for (const [a, b] of spans.sort((x, y) => x[0] - y[0])) {
    const last = out[out.length - 1]
    if (last && a <= last[1] + 1) last[1] = Math.max(last[1], b)
    else out.push([a, b])
  }
  return out
}

/**
 * The heights the rows of `spans` (in order, apart) of sheet `i` take now, as AutoFit gives them, of those whose height
 * is the default or one Excel fitted and that show, as Excel keeps such a row fitted; `runs`: in runs of rows of one
 * height, else row by row. Rows without an entry of their own are taken by spans, not one by one.
 */
export function fittedHeights(s: EngineState, i: number, spans: [number, number][], runs: boolean): SetLayoutEdit[] {
  if (!spans.length) return []
  const [lo, hi] = [spans[0][0], spans[spans.length - 1][1]]
  const sheet = s.input.sheets[i]
  const { cells, spans: spills } = stateRows(s, i, lo, hi)
  const fits = lines(fitHeights(s.input.styles ?? [], sheet, cells, lo, hi, spills), 'r')
  const rows = lines(sheet.rows, 'r')
  const out: SetLayoutEdit[] = []
  let run: [a: number, b: number, height: number | null] | null = null
  const flush = () => {
    if (!run) return
    const [a, b, height] = run
    const set = { height, bestFit: height === null ? null : true }
    if (runs) out.push({ sheet: s.sheetNames[i], rows: a === b ? String(a + 1) : `${a + 1}:${b + 1}`, set })
    else for (let r = a; r <= b; r++) out.push({ sheet: s.sheetNames[i], rows: String(r + 1), set })
  }
  const take = (a: number, b: number, height: number | null) => {
    if (run && run[1] === a - 1 && run[2] === height) run[1] = b
    else {
      flush()
      run = [a, b, height]
    }
  }
  for (const [a, b] of spans)
    for (let f = firstEnding(fits, a); f < fits.length && fits[f][0] <= b; f++) {
      const [x, y, height] = [Math.max(a, fits[f][0]), Math.min(b, fits[f][1]), fits[f][2]]
      // the rows of the piece by their entries: one without an entry has the default height
      let at = x
      for (let e = firstEnding(rows, x); e < rows.length && rows[e][0] <= y; e++) {
        const [r1, r2, row] = rows[e]
        if (at < r1 && height !== null) take(at, r1 - 1, height)
        if (!row.hidden && (row.height === undefined || row.bestFit) && (row.height ?? null) !== height) take(Math.max(at, r1), Math.min(y, r2), height)
        at = Math.min(y, r2) + 1
      }
      if (at <= y && height !== null) take(at, y, height)
    }
  flush()
  return out
}

/**
 * The cells of numbers of patches that change a number format, which widen their columns as typing does when the
 * number no longer fits (`textWidths`), as Excel widens a column for Accounting or a decimal more.
 */
export function patchedNumbers(s: EngineState, patches: CheckedPatch[]): { i: number; r: number; c: number }[] {
  const out: { i: number; r: number; c: number }[] = []
  for (const { i, rect, look } of patches) {
    if (look.format === undefined) continue
    const cols = s.ctx.sheets[i].cols
    for (let c = rect.c1; c <= Math.min(rect.c2, cols.length - 1); c++) {
      const col = cols[c]
      if (col) for (let r = rect.r1; r <= Math.min(rect.r2, col.length - 1); r++) if (typeof outCell(col.get(r) ?? null) === 'number') out.push({ i, r, c })
    }
  }
  return out
}
