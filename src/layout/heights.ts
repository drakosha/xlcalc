/**
 * Row heights as Excel for Mac fits them (`docs/formats.md`, "Fitting rows"): AutoFit gives a row the height of its
 * highest cell, each cell one line of its font, a text that wraps its lines at its column's width; a cell merged down
 * does not count, one merged across is one line, and the cells of a row without a value count by the font they show (the
 * row's style, else a column's, else the book's Calibri 11), so an empty row gets the default height.
 */
import type { BookInput, CellStyle, CellValue, ComputeResult, SheetInput } from '../contract.js'
import { toCellValue, type EngineState } from '../engine.js'
import { checkBook, checkSheet, InputError } from '../input.js'
import { layerOf } from '../looks/sheet.js'
import { MAX_COLS, parseAddress } from '../model.js'
import { columnPoints } from '../numfmt.js'
import { outValue } from '../result.js'
import { textWidth } from '../numfmt/glyphs.js'
import { sheetKey } from '../syntax.js'
import type { Placed } from './fit.js'
import { sheetLines } from './lines.js'
import { mergesOf } from './merges.js'
import { lines as spansOf, parseSpan, spanKey } from './spans.js'

/** The height of a row without its own, in points: a line of Calibri 11, the book's default font. */
export const DEFAULT_HEIGHT = 15

/** The highest row Excel makes, in points. */
export const MAX_HEIGHT = 409.5

/** The size of the book's default font, which a style without one shows. */
const DEFAULT_SIZE = 11

/** The points a wrapped text's column keeps beside its lines: two more than the padding of shown text. */
const WRAP_PADDING = 7

/** The height of one line of Calibri of `size` points, as Excel for Mac fits a row to it (`test/excel/rows/autofit.json`). */
export const lineHeight = (size: number): number => Math.round(size * 1.27 + 0.96)

/** The height of each line of a wrapped text: a point more than a line alone. */
const wrappedLine = (size: number): number => lineHeight(size) + 1

/** The width of text in Calibri of `size` points: the widths Excel shows Calibri 11 in, scaled. */
const widthIn = (s: string, size: number): number => (textWidth(s) * size) / DEFAULT_SIZE

/**
 * The lines a text takes wrapped in `space` points of Calibri of `size` points: a line at each line break, and within
 * one, the words that fit, spaces between them counted; a word wider than the room breaks between its characters.
 */
export function wrappedLines(text: string, size: number, space: number): number {
  let n = 0
  const gap = widthIn(' ', size)
  for (const para of text.split('\n')) {
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
 * The height a cell needs: wrapped text its lines at its column's width (`points`), anything else one line of its font.
 * `across`: the cell is merged with the cells beside it, one line whatever it holds.
 */
export function cellHeight(value: CellValue, style: CellStyle | undefined, points: number, across = false): number {
  const size = style?.font?.size ?? DEFAULT_SIZE
  if (typeof value !== 'string' || !value || !style?.alignment?.wrap || across) return lineHeight(size)
  return Math.min(MAX_HEIGHT, wrappedLines(value, size, points - WRAP_PADDING) * wrappedLine(size))
}

/** A height a row span needs: rows `r1` to `r2` and the points. */
type Need = [r1: number, r2: number, points: number]

/**
 * What rows `lo` to `hi` of a sheet need beside their cells: a row without a style of its own the default font's line
 * and every styled column's (`base`), as some of its cells show them; the looks of cells without an entry one line of
 * theirs. And where the rows' entries begin and end.
 */
function lineNeeds(styles: readonly CellStyle[], sheet: SheetInput, lo: number, hi: number): { base: number; needs: Need[]; cuts: number[] } {
  let base = lineHeight(DEFAULT_SIZE)
  for (const [, , col] of spansOf(sheet.columns, 'c')) if (col.style !== undefined) base = Math.max(base, lineHeight(sizeOf(styles, col.style)))
  const cuts = spansOf(sheet.rows, 'r').flatMap(([a, b]) => [a, b + 1])
  const layer = layerOf(sheet)
  const needs = [...(layer?.pieces({ r1: lo, r2: hi, c1: 0, c2: MAX_COLS - 1 }) ?? [])].flatMap((p): Need[] => {
    const style = layer!.looks.list[p.look]?.style
    return style === undefined ? [] : [[p.r1, p.r2, lineHeight(sizeOf(styles, style))]]
  })
  return { base, needs, cuts }
}

/**
 * The heights AutoFit gives rows `lo` to `hi` of a sheet whose cells hold `cells` (any order, each row's whole, blanks
 * with an entry as null), and `spans` (rows `r1` to `r2` holding one value, a large spill's), as `RowInput.height` by
 * spans of rows of one height: null for the default height.
 */
export function fitHeights(styles: readonly CellStyle[], sheet: SheetInput, cells: Iterable<Placed>, lo: number, hi: number, spans: [r1: number, r2: number, c: number, v: CellValue][] = []): Record<string, number | null> {
  const lines = sheetLines(sheet)
  const merges = mergesOf(sheet).filter((m) => m.r2 >= lo && m.r1 <= hi)
  const { base, needs, cuts: entries } = lineNeeds(styles, sheet, lo, hi)
  const need = (r: number, c: number, v: CellValue): number | null => {
    const m = merges.find((x) => r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2)
    if (m && m.r2 > m.r1) return null
    const style = lines.look(r, c).style
    return cellHeight(v, style === undefined ? undefined : styles[style], columnPoints(lines.column(c)?.width), !!m)
  }
  const rows = new Map<number, number>()
  const take = (r: number, h: number | null) => {
    if (h !== null && h > (rows.get(r) ?? 0)) rows.set(r, h)
  }
  for (const [r, c, v] of cells) if (r >= lo && r <= hi) take(r, need(r, c, v))
  for (const [r1, r2, c, v] of spans) needs.push([Math.max(r1, lo), Math.min(r2, hi), need(r1, c, v) ?? 0])
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
    let h = own === undefined ? base : lineHeight(sizeOf(styles, own))
    for (const [x, y, p] of needs) if (x <= a && y >= b) h = Math.max(h, p)
    if (a === b) h = Math.max(h, rows.get(a) ?? 0)
    h = Math.min(MAX_HEIGHT, h)
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
    if (col) for (let r = lo; r <= Math.min(hi, col.length - 1); r++) if (col[r] !== undefined && col[r] !== null) cells.push([r, c, outValue(toCellValue(col[r]))])
  }
  for (const a in sheet.cells) {
    const { row, col } = parseAddress(a)
    if (row >= lo && row <= hi && (data.cols[col]?.[row] ?? null) === null) cells.push([row, col, null])
  }
  const spans = data.fills.flatMap((f) => Array.from({ length: f.c2 - f.c1 + 1 }, (_, k): [number, number, number, CellValue] => [f.r1, f.r2, f.c1 + k, outValue(toCellValue(f.v))]))
  return { cells, spans }
}

/**
 * The heights AutoFit gives rows `5` or `2:9` of the sheet `sheet` of a computed book, from the values of its result
 * (`compute(book)`), by spans of rows of one height (`2:4`): null for the default height.
 */
export function fitRows(book: BookInput, result: ComputeResult, sheet: string, rows: string): Record<string, number | null> {
  book = checkBook(book)
  const index = (name: string) => {
    const i = book.sheets.findIndex((x) => sheetKey(x.name) === sheetKey(name))
    return i < 0 ? undefined : i
  }
  const i = checkSheet(sheet, index, 'sheet')
  const span = typeof rows === 'string' ? parseSpan(rows, 'r') : null
  if (!span) throw new InputError('rows', `expected rows such as "5" or "2:9", got ${JSON.stringify(rows)}`)
  const values = result.sheets.find((x) => sheetKey(x.name) === sheetKey(book.sheets[i].name))
  if (!values) throw new InputError('result', `the result has no sheet ${JSON.stringify(book.sheets[i].name)}`)
  const sh = book.sheets[i]
  const cells = Object.entries(values.cells).map(([a, x]): Placed => {
    const { row, col } = parseAddress(a)
    return [row, col, x.value]
  })
  for (const a of Object.keys(sh.cells)) if (!(a in values.cells)) cells.push([parseAddress(a).row, parseAddress(a).col, null])
  return fitHeights(book.styles ?? [], sh, cells, ...span)
}

