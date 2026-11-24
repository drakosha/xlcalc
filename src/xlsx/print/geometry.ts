/** Point coordinates of stored sheet geometry, with one adjustment per explicit size span. */
import { bookGeometry, columnPoints, MAX_ROW_HEIGHT } from '../../column-width.js'
import { checkNumber, integer } from '../../input/basic.js'
import type { BookInput, RowInput, SheetInput } from '../../contract.js'
import { firstEnding, isHidden, lines, type Line } from '../../layout/spans.js'
import { MAX_COLS, MAX_ROWS } from '../../model.js'
import type { PrintAxisGeometry, PrintSheetDefaults, PrintSheetGeometry, PrintSpan } from './types.js'

export interface PrintGeometryOptions extends Omit<PrintSheetDefaults, 'sheet'> {
  font?: BookInput['font']
  usedRange?: string
}

function axis<T extends { hidden?: boolean; height?: number; width?: number }>(
  entries: Line<T>[], count: number, defaultSize: number, sizeOf: (entry: T) => number,
): PrintAxisGeometry {
  let adjustment = 0
  const spans = entries.flatMap(([first, last, entry]) => {
    const perLine = (isHidden(entry) ? 0 : sizeOf(entry)) - defaultSize
    if (!perLine) return []
    const stop = last + 1
    const span = { first, stop, startAdjustment: adjustment, perLine }
    adjustment += perLine * (stop - first)
    return [span]
  })
  return {
    count,
    edge(line) {
      const before = line - 1
      let low = 0
      let high = spans.length
      while (low < high) {
        const middle = (low + high) >> 1
        if (spans[middle].first <= before) low = middle + 1
        else high = middle
      }
      const span = spans[low - 1]
      return before * defaultSize + (span ? span.startAdjustment + (Math.min(before, span.stop) - span.first) * span.perLine : 0)
    },
  }
}

/** Keep explicit height overrides inside the visible spans of a sheet hidden by default. */
function visibleRowLines(rows: Line<RowInput>[], visible: readonly PrintSpan[]): Line<RowInput>[] {
  const out: Line<RowInput>[] = []
  let previous = 0
  for (const [i, span] of visible.entries()) {
    const path = `options.visibleRows[${i}]`
    const first = integer(span.first, `${path}.first`, previous + 1, MAX_ROWS) - 1
    const last = integer(span.last, `${path}.last`, first + 1, MAX_ROWS) - 1
    previous = last + 1
    let next = first
    for (let at = firstEnding(rows, first); rows[at]?.[0] <= last; at++) {
      const [start, end, entry] = rows[at]
      if (next < start) out.push([next, start - 1, {}])
      out.push([Math.max(next, start), Math.min(end, last), entry])
      next = end + 1
    }
    if (next <= last) out.push([next, last, {}])
  }
  return out
}

export function printSheetGeometry(
  sheet: Pick<SheetInput, 'rows' | 'columns' | 'defaultColumnWidth'>,
  options: PrintGeometryOptions = {},
): PrintSheetGeometry {
  const [digit, defaultHeight] = bookGeometry(options.font)
  const rowHeight = options.defaultRowHeight ?? defaultHeight
  checkNumber(rowHeight, 'options.defaultRowHeight', 0, MAX_ROW_HEIGHT)
  const colWidth = columnPoints(sheet.defaultColumnWidth, digit)
  const rows = lines(sheet.rows, 'r')
  const visibleRows = options.visibleRows === undefined ? rows : visibleRowLines(rows, options.visibleRows)
  const implicitRowHeight = options.visibleRows === undefined ? rowHeight : 0
  return {
    rows: axis(visibleRows, MAX_ROWS, implicitRowHeight, (entry) => entry.height ?? rowHeight),
    columns: axis(lines(sheet.columns, 'c'), MAX_COLS, colWidth, (entry) => entry.width === undefined ? colWidth : columnPoints(entry.width, digit)),
    ...(options.usedRange === undefined ? {} : { usedRange: options.usedRange }),
  }
}
