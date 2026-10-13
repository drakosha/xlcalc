/**
 * Row heights as Excel for Mac fits them: a first guess, before Excel was asked. A row gets the height of its highest
 * cell: a line of its font, or a wrapped text's lines at its column's width; merged cells do not count, an empty row
 * gets the default height.
 */
import type { BookInput, CellStyle, ComputeResult, SheetInput } from '../contract.js'
import { epochOf } from '../dates.js'
import { checkBook, checkSheet, InputError } from '../input.js'
import { displayRules, type Rules } from '../locale/rules.js'
import { parseAddress } from '../model.js'
import { columnPoints, room } from '../numfmt.js'
import { advance } from '../numfmt/advances.js'
import { textWidth } from '../numfmt/glyphs.js'
import { hasCell } from '../structure/area.js'
import { sheetKey } from '../syntax.js'
import type { Placed } from './fit.js'
import { sheetLines } from './lines.js'
import { mergesOf } from './merges.js'
import { parseSpan, spanKey } from './spans.js'

/** The height of a row without its own, in points: a line of Calibri 11. */
export const DEFAULT_HEIGHT = 15

/** The highest row (`RowInput.height`). */
const MAX_HEIGHT = 409

/** The points of a line of Calibri of `size` points: its ascent, descent and gap (2500 of 2048). */
const lineOf = (size: number) => (size * 2500) / 2048

/** The width of text in Calibri of `size` points: 11 by the widths Excel shows, another size by the font's advances. */
const widthIn = (s: string, size: number) => (size === 11 ? textWidth(s) : ([...s].reduce((w, ch) => w + advance(ch), 0) * size) / 2048)

/** The lines a text takes wrapped in `space` points: a line at each line break, then at spaces, a word wider than the room by its characters. */
export function wrappedLines(text: string, size: number, space: number): number {
  let n = 0
  for (const para of text.split('\n')) {
    let line = 0
    n++
    for (const word of para.split(' ')) {
      const w = widthIn(word, size)
      const gap = line ? widthIn(' ', size) : 0
      if (line && line + gap + w > space) {
        n++
        line = 0
      }
      if (!line && w > space) {
        // a word wider than the room breaks between its characters
        let part = 0
        for (const ch of word) {
          const cw = widthIn(ch, size)
          if (part && part + cw > space) {
            n++
            part = 0
          }
          part += cw
        }
        line = part
      } else line += (line ? gap : 0) + w
    }
  }
  return n
}

/** The height a cell needs: one line of its font, or its wrapped lines. */
function cellHeight(text: string, style: CellStyle | undefined, space: number): number {
  const size = style?.font?.size ?? 11
  const lines = style?.alignment?.wrap ? wrappedLines(text, size, space) : 1
  return Math.min(MAX_HEIGHT, Math.round(lines * lineOf(size) + 1.6))
}

/**
 * The heights AutoFit gives the rows `lo` to `hi` of a sheet whose cells hold `cells`, as `RowInput.height`, null for
 * the default height.
 */
export function fitHeights(rules: Rules, styles: readonly CellStyle[], sheet: SheetInput, cells: Iterable<Placed>, lo: number, hi: number): Record<string, number | null> {
  const lines = sheetLines(sheet)
  const merges = mergesOf(sheet)
  const highest = new Map<number, number>()
  for (let r = lo; r <= hi; r++) highest.set(r, DEFAULT_HEIGHT)
  const take = (r: number, c: number, text: string) => {
    if (r < lo || r > hi || merges.some((m) => hasCell(m, r, c))) return
    const style = lines.look(r, c).style
    const h = cellHeight(text, style === undefined ? undefined : styles[style], room(columnPoints(lines.column(c)?.width)))
    if (h > highest.get(r)!) highest.set(r, h)
  }
  for (const [r, c, v] of cells) if (v !== null) take(r, c, rules.display(v, lines.format(r, c), Infinity).text)
  for (const a of Object.keys(sheet.cells)) {
    const { row, col } = parseAddress(a)
    if (sheet.cells[a].style !== undefined) take(row, col, '')
  }
  return Object.fromEntries([...highest].map(([r, h]) => [spanKey(r, r, 'r'), h === DEFAULT_HEIGHT ? null : h]))
}

/**
 * The heights AutoFit gives rows `5` or `2:9` of the sheet `sheet` of a computed book, from the values of its result,
 * by row number: null for the default height.
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
  const cells = Object.entries(values.cells).map(([a, x]): Placed => {
    const { row, col } = parseAddress(a)
    return [row, col, x.value]
  })
  return fitHeights(displayRules(book.locale, epochOf(book.dateSystem)), book.styles ?? [], book.sheets[i], cells, ...span)
}
