import type { PrintBreak, PrintSpan, SheetInput } from '../contract.js'
import { sheetKey } from '../syntax.js'
import { areaText, LIMIT, parseRange, rectOf, type Area, type Axis, type Rect } from './area.js'
import type { Transform } from './rewrite.js'

function asArea(rect: Rect, kind: Area['kind']): Area {
  const span = (first: number, last: number) => ({ lo: { i: first, abs: true }, hi: { i: last, abs: true } })
  return { kind, r: span(rect.r1, rect.r2), c: span(rect.c1, rect.c2) }
}

function lineArea(axis: Axis, first: number, last: number): Area {
  return axis === 'r'
    ? asArea({ r1: first - 1, r2: last - 1, c1: 0, c2: LIMIT.c - 1 }, 'rows')
    : asArea({ c1: first - 1, c2: last - 1, r1: 0, r2: LIMIT.r - 1 }, 'cols')
}

/** The references in print settings use the same transform as sheet-local defined names. */
export function printAfter(sheet: SheetInput, transform: Transform, axis?: Axis): void {
  const { name, print, rowDefaults } = sheet
  if (!print && !rowDefaults?.visibleRows) return
  const key = sheetKey(name)
  const move = (area: Area): Area | undefined => {
    const result = transform.ref({ sheet: name, area })
    if (result === undefined) return area
    if (typeof result !== 'object' || sheetKey(result.sheet) !== key) return undefined
    return result.area
  }
  const moveLines = (along: Axis, span: PrintSpan): PrintSpan | undefined => {
    const moved = move(lineArea(along, span.first, span.last))
    return moved && { first: moved[along].lo.i + 1, last: moved[along].hi.i + 1 }
  }
  if (print) {
    const next = { ...print }
    if (print.areas) {
      const areas = print.areas.flatMap((text) => {
        const rect = parseRange(text.replace(/\$/g, '').toUpperCase())!
        const kind = rect.r1 === 0 && rect.r2 === LIMIT.r - 1 ? 'cols'
          : rect.c1 === 0 && rect.c2 === LIMIT.c - 1 ? 'rows' : text.includes(':') ? 'range' : 'cell'
        const moved = move(asArea(rect, kind))
        return moved ? [areaText(rectOf(moved))] : []
      })
      if (areas.length) next.areas = areas
      else delete next.areas
    }
    for (const along of ['r', 'c'] as const) {
      const titles = along === 'r' ? 'repeatRows' : 'repeatColumns'
      const repeat = next[titles]
      if (repeat) {
        const moved = moveLines(along, repeat)
        if (moved) next[titles] = moved
        else delete next[titles]
      }
      const breaksKey = along === 'r' ? 'rowBreaks' : 'columnBreaks'
      const breaks = next[breaksKey]
      if (breaks && axis === along) {
        const kept: PrintBreak[] = []
        for (const br of breaks) {
          const moved = moveLines(along, { first: br.before, last: br.before })
          if (moved) kept.push({ ...br, before: moved.first })
        }
        if (kept.length) next[breaksKey] = kept
        else delete next[breaksKey]
      }
    }
    if (Object.keys(next).length) sheet.print = next
    else delete sheet.print
  }
  if (axis === 'r' && rowDefaults?.visibleRows) {
    const visibleRows = rowDefaults.visibleRows.flatMap((span) => moveLines('r', span) ?? [])
    sheet.rowDefaults = { ...rowDefaults, visibleRows }
  }
}
