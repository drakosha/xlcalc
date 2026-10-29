/** The first workbook and sheet views, passed between the file and presentation metadata. */
import type { BookInput, FrozenPanes, SheetInput, SheetView } from '../contract.js'
import { checkSheetView } from '../input-view.js'
import { formatAddress, parseAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import { attributes, element, elements, flag, type Element } from './xml/scan.js'
import { escapeAttr } from './xml/text.js'

export function readSheetView(first: Element | undefined): SheetView | undefined {
  if (!first) return undefined
  const attrs = attributes(first.attrs)
  const body = first.body ?? ''
  const paneElement = element(body, 'pane')
  const pane = attributes(paneElement?.attrs ?? '')
  const selections = elements(body, 'selection').map((s) => attributes(s.attrs))
  // the active pane's selection, a pane named by neither the top left one; without panes the first selection
  const selected = (paneElement ? selections.find((s) => (s.pane ?? 'topLeft') === (pane.activePane ?? 'topLeft')) : selections[0]) ?? {}
  const out: SheetView = {}
  const topLeft = scrollOf(attrs.topLeftCell, pane)
  if (topLeft !== undefined) out.topLeft = topLeft
  if (attrs.zoomScale !== undefined) out.zoom = Number(attrs.zoomScale)
  for (const [key, attr] of VIEW_FLAGS) if (attrs[attr] !== undefined && !flag(attrs[attr])) out[key] = false
  if (selected.activeCell !== undefined) out.activeCell = selected.activeCell
  if (selected.sqref?.trim()) out.selection = selected.sqref.trim().split(/\s+/)
  // Malformed optional view fields do not make a valid workbook unreadable.
  for (const key of Object.keys(out) as (keyof SheetView)[]) {
    try { checkSheetView({ [key]: out[key] }, 'view') } catch { delete out[key] }
  }
  return Object.keys(out).length ? out : undefined
}

const CELL = /^[A-Z]{1,3}[1-9]\d*$/

/** The view's flags shown by default, by their attributes: written as `0` where the model has false. */
export const VIEW_FLAGS = [['gridlines', 'showGridLines'], ['headings', 'showRowColHeaders']] as const

/**
 * The scroll position a view saved: its top left cell, but over frozen panes whose pane after the split is scrolled past
 * it, that pane's top left cell on each frozen axis (as `frozenScroll` writes it and Excel keeps it).
 */
function scrollOf(top: string | undefined, pane: Record<string, string>): string | undefined {
  if ((pane.state !== 'frozen' && pane.state !== 'frozenSplit') || !CELL.test(pane.topLeftCell ?? '')) return top
  const p = parseAddress(pane.topLeftCell!)
  const v = top !== undefined && CELL.test(top) ? parseAddress(top) : { row: 0, col: 0 }
  const [ys, xs] = [Math.trunc(Number(pane.ySplit ?? 0)), Math.trunc(Number(pane.xSplit ?? 0))]
  // a pane at the split is not scrolled past it
  if ((!(ys > 0) || p.row === v.row + ys) && (!(xs > 0) || p.col === v.col + xs)) return top
  return formatAddress(ys > 0 ? p.row : v.row, xs > 0 ? p.col : v.col)
}

/**
 * A saved scroll position over frozen panes as Excel keeps it: on a frozen axis the view starts at the first row or
 * column (frozen panes count from there, as the reader counts them) and the pane after the split holds the scroll, at
 * the split or past it; on an axis not frozen the view holds it. The view's and the pane's top left cells.
 */
export function frozenScroll(topLeft: string | undefined, f: FrozenPanes): { view?: string; pane: string } {
  const [rows, cols] = [f.rows ?? 0, f.columns ?? 0]
  const t = topLeft === undefined ? undefined : parseAddress(topLeft)
  return {
    ...(t ? { view: formatAddress(rows ? 0 : t.row, cols ? 0 : t.col) } : {}),
    pane: formatAddress(rows ? Math.max(t?.row ?? 0, rows) : 0, cols ? Math.max(t?.col ?? 0, cols) : 0),
  }
}

export function readActiveSheet(xml: string, declared: { name: string }[], sheets: SheetInput[]): string | undefined {
  const view = attributes(element(element(xml, 'bookViews')?.body ?? '', 'workbookView')?.attrs ?? '')
  const index = Number(view.activeTab)
  const name = view.activeTab !== undefined && Number.isInteger(index) ? declared[index]?.name : undefined
  return sheets.find((s) => s.name === name && !s.hidden)?.name
}

export function activeSheetIndex(book: BookInput): number {
  return book.activeSheet === undefined ? book.sheets.findIndex((s) => !s.hidden) : book.sheets.findIndex((s) => sheetKey(s.name) === sheetKey(book.activeSheet!))
}

/**
 * Keep the existing frozen pane and replace only its active selection when the saved view gives one. Only the active
 * sheet's tab is selected: a group of selected tabs, kept apart from `activeSheet`, would open the book grouped.
 */
export function sheetViewXml(view: SheetView | undefined, pane: string, active: boolean): string {
  const attrs = ` workbookViewId="0"${active ? ' tabSelected="1"' : ''}${view?.topLeft ? ` topLeftCell="${escapeAttr(view.topLeft)}"` : ''}${view?.zoom !== undefined ? ` zoomScale="${view.zoom}"` : ''}${VIEW_FLAGS.map(([key, attr]) => (view?.[key] === false ? ` ${attr}="0"` : '')).join('')}`
  let body = pane
  if (view?.activeCell !== undefined || view?.selection !== undefined) {
    const paneName = attributes(element(pane, 'pane')?.attrs ?? '').activePane
    const selection = `<selection${paneName ? ` pane="${paneName}"` : ''}${view.activeCell ? ` activeCell="${escapeAttr(view.activeCell)}"` : ''}${view.selection ? ` sqref="${escapeAttr(view.selection.join(' '))}"` : ''}/>`
    body = pane.replace(/<selection[^>]*\/>/g, '') + selection
  }
  return `<sheetViews><sheetView${attrs}${body ? `>${body}</sheetView>` : '/>'}</sheetViews>`
}
