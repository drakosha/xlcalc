/**
 * Cell formats only the file tells apart: a cell, row or column whose cell format looks as an earlier one of the file
 * (`KeptStyles.twin`: a quote prefix, a theme colour, a cell style, a pattern fill the model does not hold) keeps it when
 * rows or columns move, and a column keeps it when the columns are written again. `openKept` gives each such format of a
 * sheet carriers of its rectangles (`formatAreas`): a row's run of cells of the format, the same runs of the rows below
 * in one rectangle, and whole rows and columns for those of rows and columns. `writeKept` gives a cell, row or column it
 * writes that looks as the format of the rectangle it lies in now that format (`Formats`). Rows inserted inside a
 * rectangle stretch it, so the new cells take its format where they look the same, as Excel gives them the row above's.
 */
import { colToIndex, MAX_COLS, MAX_ROWS } from '../../model.js'
import type { Rect } from '../../structure/area.js'
import type { KeptStyles } from './styles.js'
import type { Placed } from './xml.js'

export type FormatKind = 'cells' | 'rows' | 'cols'

/** A rectangle of a sheet whose cells, rows or columns have the file's cell format `s`. */
export interface FormatArea {
  kind: FormatKind
  s: number
  rect: Rect
}

/** The attributes of a column. */
const COL = /<(?:[\w.-]+:)?col(?=[\s/>])([^>]*)>/g
const attr = (name: string) => new RegExp(`\\s${name}\\s*=\\s*["']([^"']*)["']`)
const [R, STYLE, CUSTOM, MIN, MAX] = ['r', 'style', 'customFormat', 'min', 'max'].map(attr)

/**
 * A sheet's rectangles of cell formats that look as earlier ones (`twin`), cells first, and how many cells they hold. A
 * run of a row extends the rectangle of the same run of the row above. Only the cell formats are looked for, so a sheet
 * of other formats costs one search of its text.
 */
export function formatAreas(xml: string, data: Placed | undefined, cols: Placed | undefined, twin: (s: number) => boolean): { areas: FormatArea[]; cells: number } {
  const cells: FormatArea[] = []
  const lines: FormatArea[] = []
  // the last rectangle of each run, by its format and columns, of cells and of rows
  const open = { cells: new Map<number, Rect>(), rows: new Map<number, Rect>() }
  const add = (kind: 'cells' | 'rows', s: number, r: number, c1: number, c2: number) => {
    const key = (s * MAX_COLS + c1) * MAX_COLS + c2
    const x = open[kind].get(key)
    if (x && x.r2 === r - 1) x.r2 = r
    else {
      const rect = { r1: r, r2: r, c1, c2 }
      open[kind].set(key, rect)
      ;(kind === 'cells' ? cells : lines).push({ kind, s, rect })
    }
  }
  let count = 0
  let run: { s: number; r: number; c1: number; c2: number } | undefined
  // a cell's or a row's format as writers write it, ` s="3"`: one written otherwise keeps its format only where it stays
  for (let i = data ? xml.indexOf(' s="', data.open) : -1; data && i >= 0 && i < data.end; i = xml.indexOf(' s="', i)) {
    const from = i
    let s = 0
    for (let ch = xml.charCodeAt((i += 4)); ch >= 48 && ch <= 57; ch = xml.charCodeAt(++i)) s = s * 10 + ch - 48
    if (i === from + 4 || xml.charCodeAt(i) !== 34 || !twin(s)) continue
    const at = xml.lastIndexOf('<', from)
    // a cell as writers write it, `<c r="B7"`, read at once; others by their attributes
    let [row, col] = xml.startsWith('<c r="', at) ? cellAt(xml, at + 6) : [-1, -1]
    if (row < 0) {
      const tag = xml.slice(at, xml.indexOf('>', from))
      const name = /^<(?:[\w.-]+:)?(\w+)/.exec(tag)?.[1]
      const r = R.exec(tag)?.[1]
      if (name === 'row') {
        // a row without its number keeps its format only where it stays
        if (r && /^\d+$/.test(r) && /^(1|true)$/.test(CUSTOM.exec(tag)?.[1] ?? '')) add('rows', s, Number(r) - 1, 0, MAX_COLS - 1)
        continue
      }
      const ref = name === 'c' && r ? /^([A-Za-z]{1,3})(\d+)$/.exec(r) : null
      if (!ref) continue
      ;[row, col] = [Number(ref[2]) - 1, colToIndex(ref[1].toUpperCase())]
    }
    count++
    if (run && run.s === s && run.r === row && run.c2 === col - 1) run.c2 = col
    else {
      if (run) add('cells', run.s, run.r, run.c1, run.c2)
      run = { s, r: row, c1: col, c2: col }
    }
  }
  if (run) add('cells', run.s, run.r, run.c1, run.c2)
  for (const m of cols ? xml.slice(cols.open, cols.end).matchAll(COL) : []) {
    const [s, lo, hi] = [STYLE, MIN, MAX].map((x) => Number(x.exec(m[1])?.[1]))
    if (!twin(s) || !(lo >= 1 && hi >= lo && hi <= MAX_COLS)) continue
    const last = lines.at(-1)
    if (last?.kind === 'cols' && last.s === s && last.rect.c2 === lo - 2) last.rect.c2 = hi - 1
    else lines.push({ kind: 'cols', s, rect: { r1: 0, r2: MAX_ROWS - 1, c1: lo - 1, c2: hi - 1 } })
  }
  return { areas: [...cells, ...lines], cells: count }
}

/** The row and column of a cell's address at `i` of a text, ended by a quote (`B7"`); -1 for another text. */
function cellAt(xml: string, i: number): [number, number] {
  let [col, row] = [0, 0]
  let k = i
  for (let ch = xml.charCodeAt(k); ch >= 65 && ch <= 90 && k - i < 3; ch = xml.charCodeAt(++k)) col = col * 26 + ch - 64
  const digits = k
  for (let ch = xml.charCodeAt(k); ch >= 48 && ch <= 57; ch = xml.charCodeAt(++k)) row = row * 10 + ch - 48
  return k > digits && digits > i && xml.charCodeAt(k) === 34 && row >= 1 && row <= MAX_ROWS && col <= MAX_COLS ? [row - 1, col - 1] : [-1, -1]
}

/** The formats of a sheet's rectangles where they lie now: what a cell, row or column written takes. */
export class Formats {
  constructor(
    private readonly areas: FormatArea[],
    private readonly styles: KeptStyles,
  ) {}

  /**
   * The cell formats of the cells and rows of `rows` (in order): a cell's or row's of the look `xf` stands for, the
   * format of the rectangle it lies in where that looks so, else `xf`.
   */
  index(rows: number[]): { cell(r: number, c: number, xf: number): number; row(r: number, xf: number): number } {
    const cells = new Map<number, FormatArea[]>()
    const lines = new Map<number, number>()
    for (const a of this.areas) {
      if (a.kind === 'cols') continue
      let lo = 0
      let hi = rows.length
      while (lo < hi) {
        const mid = (lo + hi) >> 1
        if (rows[mid] < a.rect.r1) lo = mid + 1
        else hi = mid
      }
      for (let i = lo; i < rows.length && rows[i] <= a.rect.r2; i++)
        if (a.kind === 'rows') lines.set(rows[i], a.s)
        else (cells.get(rows[i]) ?? cells.set(rows[i], []).get(rows[i])!).push(a)
    }
    const keep = (s: number | undefined, xf: number) => (s === undefined ? xf : this.styles.keep(s, xf))
    return {
      cell: (r, c, xf) => keep(cells.get(r)?.find((a) => a.rect.c1 <= c && c <= a.rect.c2)?.s, xf),
      row: (r, xf) => keep(lines.get(r), xf),
    }
  }

  /** The writer's `<cols>` with each column's cell format the format of the columns it lies in now where that looks so. */
  cols(text: string): string {
    const areas = this.areas.filter((a) => a.kind === 'cols').sort((a, b) => a.rect.c1 - b.rect.c1)
    if (!areas.length) return text
    return text.replace(/<col min="(\d+)" max="(\d+)"([^>]*?)\/>/g, (_, min: string, max: string, rest: string) => {
      const xf = Number(/ style="(\d+)"/.exec(rest)?.[1] ?? 0)
      const bare = rest.replace(/ style="\d+"/, '')
      const [lo, hi] = [Number(min) - 1, Number(max) - 1]
      const pieces: [number, number, number][] = []
      const piece = (c1: number, c2: number, s: number) => {
        const last = pieces.at(-1)
        if (last && last[2] === s) last[1] = c2
        else pieces.push([c1, c2, s])
      }
      let c = lo
      for (const a of areas) {
        if (a.rect.c2 < c || a.rect.c1 > hi) continue
        if (a.rect.c1 > c) piece(c, a.rect.c1 - 1, xf)
        const end = Math.min(hi, a.rect.c2)
        piece(Math.max(c, a.rect.c1), end, this.styles.keep(a.s, xf))
        c = end + 1
      }
      if (c <= hi) piece(c, hi, xf)
      return pieces.map(([c1, c2, s]) => `<col min="${c1 + 1}" max="${c2 + 1}"${bare}${s ? ` style="${s}"` : ''}/>`).join('')
    })
  }
}
