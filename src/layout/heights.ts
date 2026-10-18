/**
 * Row heights as Excel for Mac fits them (`docs/formats.md`, "Fitting rows"): AutoFit stands the cells of a row on one
 * line at its bottom, each cell one line of its font or a text that wraps its lines at its column's width, and gives the
 * row the most any cell needs above that line and the most any needs below; a cell merged down does not count, one
 * merged across is one line, and the cells of a row without a value count by the font they show (the row's style, else
 * a column's, else the book's Calibri 11), so an empty row gets the default height.
 */
import { DEFAULT_ROW_HEIGHT as DEFAULT_HEIGHT, MAX_ROW_HEIGHT as MAX_HEIGHT } from '../column-width.js'
import type { BookInput, CellStyle, CellValue, ComputeResult, SheetInput } from '../contract.js'
import { toCellValue, type EngineState } from '../engine.js'
import { layerOf } from '../looks/sheet.js'
import { MAX_COLS, parseAddress } from '../model.js'
import { columnPoints } from '../numfmt.js'
import { outValue } from '../result.js'
import { advance } from '../numfmt/advances.js'
import { resultSheet, type Placed } from './fit.js'
import { sheetLines } from './lines.js'
import { mergesOf } from './merges.js'
import { lines as spansOf, spanKey } from './spans.js'

export { DEFAULT_HEIGHT, MAX_HEIGHT }

/** The size of the book's default font, which a style without one shows. */
const DEFAULT_SIZE = 11

/** The points a wrapped text's column keeps beside its lines: two more than the padding of shown text. */
const WRAP_PADDING = 7

/** The height of one line of Calibri of `size` points, as Excel for Mac fits a row to it (`test/excel/rows/autofit.json`); a size of a half point counts as the next whole one (`held.json`). */
export const lineHeight = (size: number): number => Math.round(Math.round(size) * 1.27 + 0.96)

/**
 * The width of text in Calibri of `size` points as Excel for Mac wraps it: each character the font's advance rounded to
 * a point, as Excel shows Calibri 11 (`glyphs.ts`), at its size (`test/excel/rows/held.json`, the set `scale`).
 */
function widthIn(s: string, size: number): number {
  let w = 0
  for (const ch of s) w += Math.round((advance(ch) * size) / 2048)
  return w
}

/**
 * The lines a text takes wrapped in `space` points of Calibri of `size` points: a line at each line break, and within
 * one, the words that fit, spaces between them counted; a word wider than the room breaks between its characters.
 */
export function wrappedLines(text: string, size: number, space: number): number {
  let n = 0
  const gap = widthIn(' ', size)
  // spaces at the end of a line hang past the room
  for (const para of text.split('\n').map((x) => x.replace(/ +$/, ''))) {
    n++
    let line = -1
    for (const word of para.split(' ')) {
      const w = widthIn(word, size)
      if (line >= 0 && line + gap + w <= space) {
        line += gap + w
        continue
      }
      if (line >= 0) n++
      line = 0
      for (const ch of word) {
        const cw = widthIn(ch, size)
        if (line > 0 && line + cw > space) {
          n++
          line = 0
        }
        line += cw
      }
    }
  }
  return n
}

/** The font size of a style of the book, the default's without one. */
const sizeOf = (styles: readonly CellStyle[], style: number | undefined) => (style === undefined ? undefined : styles[style]?.font?.size) ?? DEFAULT_SIZE

/**
 * What a cell needs of its row, as Excel for Mac stands cells on one line at the bottom of the row: the points above it
 * and below it (`test/excel/rows/lines.json`). A line of Calibri of `s` points stands `s` points above and its line's
 * rest below; a wrapped text's lines above its last are a point higher than a line alone, the first 20 of them.
 */
export type Need = [above: number, below: number]

/** One line of a font of `size` points. */
const lineNeed = (size: number): Need => [Math.round(size), lineHeight(size) - Math.round(size)]

/**
 * What a cell needs (`Need`): wrapped text its lines at its column's width (`points`), anything else one line of its
 * font. `across`: the cell is merged with the cells beside it, one line, or wrapped no more than the line's rest below.
 */
export function cellNeed(value: CellValue, style: CellStyle | undefined, points: number, across = false): Need {
  const size = style?.font?.size ?? DEFAULT_SIZE
  const [above, below] = lineNeed(size)
  if (typeof value !== 'string' || !value || !style?.alignment?.wrap) return [above, below]
  if (across) return [0, below]
  const n = wrappedLines(value, size, points - WRAP_PADDING)
  return [Math.min(MAX_HEIGHT, n * lineHeight(size) + Math.min(n, 20) - below), below]
}

/** What a span of rows needs: rows `r1` to `r2`, the points above and below. */
type SpanNeed = [r1: number, r2: number, above: number, below: number]

/** The greater of two needs, side by side. */
const most = (a: Need, b: Need): Need => [Math.max(a[0], b[0]), Math.max(a[1], b[1])]

/**
 * What rows `lo` to `hi` of a sheet need beside their cells: a row without a style of its own the default font's line
 * and every styled column's (`base`), as some of its cells show them; the looks of cells without an entry one line of
 * theirs. And where the rows' entries begin and end.
 */
function lineNeeds(styles: readonly CellStyle[], sheet: SheetInput, lo: number, hi: number): { base: Need; needs: SpanNeed[]; cuts: number[] } {
  let base = lineNeed(DEFAULT_SIZE)
  for (const [, , col] of spansOf(sheet.columns, 'c')) if (col.style !== undefined) base = most(base, lineNeed(sizeOf(styles, col.style)))
  const cuts = spansOf(sheet.rows, 'r').flatMap(([a, b]) => [a, b + 1])
  const layer = layerOf(sheet)
  const needs = [...(layer?.pieces({ r1: lo, r2: hi, c1: 0, c2: MAX_COLS - 1 }) ?? [])].flatMap((p): SpanNeed[] => {
    const style = layer!.looks.list[p.look]?.style
    return style === undefined ? [] : [[p.r1, p.r2, ...lineNeed(sizeOf(styles, style))]]
  })
  return { base, needs, cuts }
}

/**
 * The heights AutoFit gives rows `lo` to `hi` of a sheet whose cells hold `cells` (any order, each row's whole, blanks
 * with an entry as null), and `spans` (rows `r1` to `r2` holding one value, a large spill's), as `RowInput.height` by
 * spans of rows of one height: null for the default height. A row is as high as its cells need above their line and
 * below it, each the most a cell needs.
 */
export function fitHeights(styles: readonly CellStyle[], sheet: SheetInput, cells: Iterable<Placed>, lo: number, hi: number, spans: [r1: number, r2: number, c: number, v: CellValue][] = []): Record<string, number | null> {
  const lines = sheetLines(sheet)
  const merges = mergesOf(sheet).filter((m) => m.r2 >= lo && m.r1 <= hi)
  const { base, needs, cuts: entries } = lineNeeds(styles, sheet, lo, hi)
  const need = (r: number, c: number, v: CellValue): Need | null => {
    const m = merges.find((x) => r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2)
    // a cell merged down counts for nothing
    if (m && m.r2 > m.r1) return null
    const style = lines.look(r, c).style
    return cellNeed(v, style === undefined ? undefined : styles[style], columnPoints(lines.column(c)?.width), !!m)
  }
  const rows = new Map<number, Need>()
  for (const [r, c, v] of cells) {
    const x = r >= lo && r <= hi ? need(r, c, v) : null
    if (x) rows.set(r, most(rows.get(r) ?? [0, 0], x))
  }
  for (const [r1, r2, c, v] of spans) {
    const x = need(r1, c, v)
    if (x) needs.push([Math.max(r1, lo), Math.min(r2, hi), ...x])
  }
  // the rows in runs of one height: their style's line or the base, the needs over them, a row's own cells
  const cuts = [...new Set([lo, hi + 1, ...entries, ...needs.flatMap(([a, b]) => [a, b + 1]), ...[...rows.keys()].flatMap((r) => [r, r + 1])])].filter((x) => x >= lo && x <= hi + 1).sort((a, b) => a - b)
  const out: Record<string, number | null> = {}
  let run: [number, number, number] | null = null
  const flush = () => {
    if (run) out[spanKey(run[0], run[1], 'r')] = run[2] === DEFAULT_HEIGHT ? null : run[2]
  }
  for (let k = 0; k + 1 < cuts.length; k++) {
    const [a, b] = [cuts[k], cuts[k + 1] - 1]
    const own = lines.row(a)?.style
    let n = own === undefined ? base : lineNeed(sizeOf(styles, own))
    for (const [x, y, above, below] of needs) if (x <= a && y >= b) n = most(n, [above, below])
    if (a === b) n = most(n, rows.get(a) ?? [0, 0])
    const h = Math.min(MAX_HEIGHT, n[0] + n[1])
    if (run && run[2] === h && run[1] === a - 1) run[1] = b
    else {
      flush()
      run = [a, b, h]
    }
  }
  flush()
  return out
}

/** The values of rows `lo` to `hi` of a sheet of the state, each row whole (blanks with an entry as null), and its large spills as spans. */
export function stateRows(s: EngineState, i: number, lo: number, hi: number): { cells: Placed[]; spans: [number, number, number, CellValue][] } {
  const data = s.ctx.sheets[i]
  const sheet = s.input.sheets[i]
  const cells: Placed[] = []
  for (let c = 0; c < data.cols.length; c++) {
    const col = data.cols[c]
    if (col)
      for (let r = lo; r <= Math.min(hi, col.length - 1); r++) {
        const v = col.get(r)
        if (v !== undefined && v !== null) cells.push([r, c, outValue(toCellValue(v))])
      }
  }
  for (const a in sheet.cells) {
    const { row, col } = parseAddress(a)
    if (row >= lo && row <= hi && (data.cols[col]?.get(row) ?? null) === null) cells.push([row, col, null])
  }
  const spans = data.fills.flatMap((f) => Array.from({ length: f.c2 - f.c1 + 1 }, (_, k): [number, number, number, CellValue] => [f.r1, f.r2, f.c1 + k, outValue(toCellValue(f.v))]))
  return { cells, spans }
}

/**
 * The heights AutoFit gives rows `5` or `2:9` of the sheet `sheet` of a computed book, from the values of its result
 * (`compute(book)`), by spans of rows of one height (`2:4`): null for the default height.
 */
export function fitRows(book: BookInput, result: ComputeResult, sheet: string, rows: string): Record<string, number | null> {
  const x = resultSheet(book, result, sheet, rows, 'r')
  const sh = x.book.sheets[x.i]
  // a cell with an entry and no value counts by its font
  const valued = new Set(x.cells.map(([r, c]) => `${r}:${c}`))
  for (const a of Object.keys(sh.cells)) {
    const { row, col } = parseAddress(a)
    if (!valued.has(`${row}:${col}`)) x.cells.push([row, col, null])
  }
  return fitHeights(x.book.styles ?? [], sh, x.cells, x.lo, x.hi)
}
