/** A selection as plain-text clipboard data: visible lines, computed values and formats without column clipping. */
import type { CopyTextOptions } from '../contract.js'
import { lookShown, windowLooks } from '../cond/looks.js'
import { toCellValue, valueAt, type EngineState } from '../engine.js'
import { allot, CalcError, paced, type Steps } from '../exec/budget.js'
import { checkBoolean, checkNumber, object, onlyOptions } from '../input.js'
import { hiddenSpansOf } from '../layout/hidden.js'
import { sheetLines } from '../layout/lines.js'
import { mergesOf } from '../layout/merges.js'
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { lowerBound } from '../sorted.js'
import type { Rect } from '../structure/area.js'
import { formatClipboard } from '../text.js'
import { sheetExtent } from './extent.js'
import { windowLines } from './viewport.js'

const MAX_CELLS = 20_000
const MAX_CHARS = 16 * 1024 * 1024

export function checkCopyOptions(v: unknown): CopyTextOptions {
  if (v === undefined) return {}
  const o = object(v, 'opts', 'copy options such as { trailingNewline: false, maxCells: 20000 }')
  onlyOptions(o, 'opts', ['trailingNewline', 'maxCells', 'maxChars'])
  if (o.trailingNewline !== undefined) checkBoolean(o.trailingNewline, 'opts.trailingNewline')
  if (o.maxCells !== undefined) checkNumber(o.maxCells, 'opts.maxCells', 1, MAX_ROWS * MAX_COLS, true)
  if (o.maxChars !== undefined) checkNumber(o.maxChars, 'opts.maxChars', 1, 0x1fffffe8, true)
  return o as CopyTextOptions
}

/** The visible line count, before allocating or reading any cells. */
function visibleCount(lo: number, hi: number, spans: [number, number][]): number {
  let n = hi - lo + 1
  for (const [a, b] of spans) n -= Math.max(0, Math.min(hi, b) - Math.max(lo, a) + 1)
  return n
}

/** Visible zero-based lines, skipping whole hidden spans. */
function* visibleLines(lo: number, hi: number, spans: [number, number][]): Generator<number> {
  let next = lo
  for (const [a, b] of spans) {
    if (b < next) continue
    if (a > hi) break
    for (let i = next; i < a; i++) yield i
    next = b + 1
  }
  for (let i = next; i <= hi; i++) yield i
}

export function* copyText(state: EngineState, s: number, asked: Rect, o: CopyTextOptions): Steps<string> {
  const sheet = state.input.sheets[s]
  const extent = sheetExtent(state, s)
  const x = { ...asked }
  // An empty sheet copies one blank field. Whole rows/columns stop at the used extent; explicit rectangles keep blanks.
  if (!extent.rows && !extent.columns) { x.r2 = x.r1; x.c2 = x.c1 }
  else {
    if (x.r1 === 0 && x.r2 === MAX_ROWS - 1) x.r2 = Math.max(x.r1, extent.rows - 1)
    if (x.c1 === 0 && x.c2 === MAX_COLS - 1) x.c2 = Math.max(x.c1, extent.columns - 1)
  }
  const hiddenRows = hiddenSpansOf(sheet.rows, 'r')
  const hiddenCols = hiddenSpansOf(sheet.columns, 'c')
  const size = visibleCount(x.r1, x.r2, hiddenRows) * visibleCount(x.c1, x.c2, hiddenCols)
  const max = o.maxCells ?? MAX_CELLS
  if (size > max) throw new CalcError('limit', `copy contains ${size} cells, over the limit of ${max}`, { limit: 'cells', max, size })
  allot(size)
  if (!size) return ''
  const rows = [...visibleLines(x.r1, x.r2, hiddenRows)]
  const cols = [...visibleLines(x.c1, x.c2, hiddenCols)]
  const w = windowLines(x, { rows: rows.map((r) => r + 1), columns: cols.map((c) => c + 1), maxCells: max })
  const cond = yield* windowLooks(state, s, w.rects)
  const covered = new Set<number>()
  for (const m of mergesOf(sheet)) {
    for (let ri = lowerBound(rows, m.r1, (r) => r); ri < rows.length && rows[ri] <= m.r2; ri++)
      for (let ci = lowerBound(cols, m.c1, (c) => c); ci < cols.length && cols[ci] <= m.c2; ci++) {
        if (paced(1)) yield
        const r = rows[ri], c = cols[ci]
        if (r !== m.r1 || c !== m.c1) covered.add(r * MAX_COLS + c)
      }
  }
  const formats = sheetLines(sheet, state)
  const lines: string[] = []
  const maxChars = o.maxChars ?? MAX_CHARS
  let chars = o.trailingNewline === false ? 0 : 2
  const countText = (text: string, separator: number) => {
    chars += separator + text.length + (/[\t\r\n]/.test(text) ? 2 + (text.match(/"/g)?.length ?? 0) : 0)
    if (chars > maxChars) throw new CalcError('limit', `copy text exceeds ${maxChars} characters`, { max: maxChars, size: chars })
  }
  for (const r of rows) {
    countText('', lines.length ? 2 : 0)
    const fields: string[] = []
    for (const c of cols) {
      if (paced(1)) yield
      const k = r * MAX_COLS + c
      if (covered.has(k)) { countText('', fields.length ? 1 : 0); fields.push(''); continue }
      const value = toCellValue(valueAt(state, s, r, c))
      const view = state.ctx.rules.display(value, formats.format(r, c), Infinity)
      const look = cond.get(k)
      if (look) lookShown(view, look, (format) => state.ctx.rules.display(value, format, Infinity))
      countText(view.text, fields.length ? 1 : 0)
      fields.push(view.text)
    }
    lines.push(formatClipboard([fields], { trailingNewline: false }))
  }
  return lines.join('\r\n') + (o.trailingNewline === false ? '' : '\r\n')
}
