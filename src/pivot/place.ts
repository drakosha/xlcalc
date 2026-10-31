/**
 * PivotTables laid out written to their sheets as one step (a refresh of a cache shared by several): every new area
 * checked first, then the old areas cleared and the new cells written, the definitions with their new layout and place.
 */
import type { BookInput, AlignmentStyle, CellInput, PivotTable, SheetInput } from '../contract.js'
import { CONTENT, holds } from '../array-areas.js'
import { dateToSerial } from '../date-serial.js'
import { say } from '../locale/messages.js'
import { formatAddress, parseAddress, readAddress } from '../model.js'
import { contains, cornersText, hasCell, meets, parseRange, type Rect } from '../structure/area.js'
import type { Applied } from '../structure/geometry.js'
import { sheetIndex } from '../structure/rewrite.js'
import { ERROR_TEXT } from '../syntax/errors.js'
import { internStyle } from '../style-table.js'
import { withField } from '../looks/sheet.js'
import { cellAlignment, labelMerges } from './looks.js'
import { pivotArea, pivotParts } from './area.js'
import type { PivotLaidOut } from './layout.js'

/** A PivotTable to write: its sheet, its definition as refreshed (its `name` the one on the sheet), and its layout. */
export interface Placement {
  sheet: string
  table: PivotTable
  laid: PivotLaidOut
}

/** A cell's fields that a PivotTable's layout writes: its content and number format; its style, note and link stay. */
const LAID = [...CONTENT, 'numberFormat'] as const

const cleared = (c: CellInput): CellInput | undefined => {
  const out: CellInput = { ...c }
  for (const k of LAID) delete out[k]
  return Object.keys(out).length ? out : undefined
}

/** What a PivotTable laid out over cells it cannot take becomes: refused (`pivot`), or spilled, as Excel 365 refreshes it. */
export type Blocked = 'spill' | 'refuse'

/**
 * Writes the PivotTables `placed` to their sheets at once (`Placement`). An area over cells that hold something, a table,
 * another PivotTable or a PivotTable laid before it in `placed` is refused (`pivot`) with `blocked: 'refuse'` (Excel's
 * new and moved PivotTable); with `spill`, as a refresh in Excel 365, the PivotTable shrinks to the top-left cell of its
 * old range, which shows #SPILL!, keeping its new lines, its old cells cleared and the others left as they are.
 */
export function placePivots(book: BookInput, placed: Placement[], blocked: Blocked = 'spill'): Applied {
  const steps = placed.map((p) => {
    const s = sheetIndex(book, p.sheet)
    const pivots = book.sheets[s]?.pivotTables ?? []
    const at = pivots.findIndex((t) => t.name === p.table.name)
    if (at < 0) throw new Error(`no PivotTable ${p.table.name} on ${p.sheet}`)
    return { ...p, s, at, old: pivotArea(pivots[at]), parts: pivotParts(pivots[at]), spill: false }
  })
  // each new area against the cells it takes: cells holding something outside its own old area, tables, the other
  // PivotTables where they are, and those laid before it where they go
  for (const [n, p] of steps.entries()) {
    const sheet = book.sheets[p.s]
    const x = p.laid.area
    const others = (sheet.pivotTables ?? []).flatMap((t, i) => {
      if (i === p.at) return []
      const q = steps.find((y) => y.s === p.s && y.at === i)
      const areas = [pivotArea(t), q && steps.indexOf(q) < n && !q.spill ? q.laid.area : null]
      return areas.some((a) => a && meets(x, a)) ? [t.name] : []
    })
    const table = (sheet.tables ?? []).find((t) => {
      const r = parseRange(t.range)
      return r && meets(x, r)
    })
    // the cells its new parts take, not the blank row under its filters nor beside them, which hold anything
    const data = pivotParts({ location: p.laid.location }).some((part) => filled(sheet, part, p.parts))
    if (!others.length && !table && !data) continue
    if (blocked === 'spill') {
      p.spill = true
      continue
    }
    const more = others.length ? { key: 'pivotOverPivot', other: others[0] } : table ? { key: 'pivotOverTable', table: table.name } : { key: 'pivotOverData', sheet: sheet.name, range: cornersText(x) }
    return { ok: false, code: 'pivot', message: say(book.locale, { pivot: p.table.name, ...more } as never) }
  }
  const sheets = [...book.sheets]
  let styles = book.styles ?? []
  // the look Excel keeps in the cells of the area (`looks.ts`), over the cell's own style
  const aligned = (c: CellInput, alignment: AlignmentStyle): CellInput => {
    const base = c.style === undefined ? {} : styles[c.style]
    const next = internStyle(styles, { ...base, alignment: { ...base.alignment, ...alignment } })
    styles = next.styles
    return { ...c, style: next.index }
  }
  for (const p of steps) {
    const sheet = sheets[p.s]
    const cells: SheetInput['cells'] = { ...sheet.cells }
    // its old cells cleared, not those beside its filters (the corpus: a value right of a page field kept)
    for (const part of p.parts)
      for (const [addr, c] of cellsIn(cells, part)) {
        const rest = cleared(c)
        if (rest) cells[addr] = rest
        else delete cells[addr]
      }
    const old = sheet.pivotTables![p.at].location.range.split(':')[0]
    const laid: PivotLaidOut = p.spill ? { ...p.laid, location: { range: old, firstHeaderRow: 0, firstDataRow: 0, firstDataColumn: 0 }, cells: [{ ...parseAddress(old), error: 'SPILL' }] } : p.laid
    for (const c of laid.cells) {
      const addr = formatAddress(c.row, c.col)
      let out: CellInput = { ...cells[addr] }
      if (c.error) out.formula = ERROR_TEXT[c.error]
      // a date as Excel keeps it in the cell, its serial number in the date format (the corpus: a page field's date item)
      else if (c.value !== undefined) out.value = typeof c.value === 'object' ? dateToSerial(c.value.date, { dateSystem: book.dateSystem }) : c.value
      if (c.numberFormat !== undefined) out.numberFormat = c.numberFormat
      const alignment = p.spill ? undefined : cellAlignment(p.table, laid, c)
      if (alignment) out = aligned(out, alignment)
      cells[addr] = out
    }
    const pivotTables = [...(sheet.pivotTables ?? [])]
    pivotTables[p.at] = { ...p.table, location: laid.location, rowLines: laid.rowLines, columnLines: laid.columnLines }
    // the merged labels: the old area's go with it, where its PivotTable merged them, the new ones are added
    const was = sheet.pivotTables![p.at].mergeLabels && p.old ? sheet.merges?.filter((m) => !contains(p.old!, parseRange(m)!)) : sheet.merges
    const merged = p.spill ? [] : labelMerges(p.table, laid).map(cornersText)
    const merges = [...(was ?? []), ...merged]
    sheets[p.s] = withField({ ...sheet, cells, pivotTables }, 'merges', merges.length ? merges : undefined)
  }
  return { ok: true, book: { ...book, sheets, ...(styles.length ? { styles } : {}) }, geometry: [] }
}

/** The cells of `x` among `cells`, with their addresses. */
function cellsIn(cells: SheetInput['cells'], x: Rect): [string, CellInput][] {
  const at = { r: 0, c: 0 }
  return Object.entries(cells).filter(([addr]) => {
    readAddress(addr, at)
    return hasCell(x, at.r, at.c)
  })
}

/** Whether a cell of `x` outside the areas `own` holds a value or a formula. */
function filled(sheet: SheetInput, x: Rect, own: Rect[]): boolean {
  const at = { r: 0, c: 0 }
  for (const [addr, c] of cellsIn(sheet.cells, x)) {
    readAddress(addr, at)
    if (holds(c) && !own.some((o) => hasCell(o, at.r, at.c))) return true
  }
  return false
}
