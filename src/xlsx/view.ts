/** The first workbook and sheet views, passed between the file and presentation metadata. */
import type { BookInput, SheetInput, SheetView } from '../contract.js'
import { checkSheetView } from '../input-view.js'
import { sheetKey } from '../syntax.js'
import { attributes, element, elements, flag } from './xml/scan.js'
import { escapeAttr } from './xml/text.js'

export function readSheetView(head: string): SheetView | undefined {
  const first = element(element(head, 'sheetViews')?.body ?? '', 'sheetView')
  if (!first) return undefined
  const attrs = attributes(first.attrs)
  const body = first.body ?? ''
  const activePane = attributes(element(body, 'pane')?.attrs ?? '').activePane
  const selections = elements(body, 'selection').map((s) => attributes(s.attrs))
  const selected = selections.find((s) => s.pane === activePane) ?? selections[0] ?? {}
  const out: SheetView = {}
  if (attrs.tabSelected !== undefined) out.tabSelected = flag(attrs.tabSelected)
  if (attrs.topLeftCell !== undefined) out.topLeft = attrs.topLeftCell
  if (attrs.zoomScale !== undefined) out.zoom = Number(attrs.zoomScale)
  if (selected.activeCell !== undefined) out.activeCell = selected.activeCell
  if (selected.sqref?.trim()) out.selection = selected.sqref.trim().split(/\s+/)
  // Malformed optional view fields do not make a valid workbook unreadable.
  for (const key of Object.keys(out) as (keyof SheetView)[]) {
    try { checkSheetView({ [key]: out[key] }, 'view') } catch { delete out[key] }
  }
  return Object.keys(out).length ? out : undefined
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

/** Keep the existing frozen pane and replace only its active selection when the saved view gives one. */
export function sheetViewXml(view: SheetView | undefined, pane: string, active: boolean): string {
  const selected = view?.tabSelected ?? active
  const attrs = ` workbookViewId="0"${selected ? ' tabSelected="1"' : view?.tabSelected === false ? ' tabSelected="0"' : ''}${view?.topLeft ? ` topLeftCell="${escapeAttr(view.topLeft)}"` : ''}${view?.zoom !== undefined ? ` zoomScale="${view.zoom}"` : ''}`
  let body = pane
  if (view?.activeCell !== undefined || view?.selection !== undefined) {
    const paneName = attributes(element(pane, 'pane')?.attrs ?? '').activePane
    const selection = `<selection${paneName ? ` pane="${paneName}"` : ''}${view.activeCell ? ` activeCell="${escapeAttr(view.activeCell)}"` : ''}${view.selection ? ` sqref="${escapeAttr(view.selection.join(' '))}"` : ''}/>`
    body = pane.replace(/<selection[^>]*\/>/g, '') + selection
  }
  return `<sheetViews><sheetView${attrs}${body ? `>${body}</sheetView>` : '/>'}</sheetViews>`
}
