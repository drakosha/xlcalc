import { MAX_ROW_HEIGHT } from '../../column-width.js'
import { MAX_COLS, MAX_ROWS } from '../../model.js'
import { areaText } from '../../structure/area.js'
import { sheetKey } from '../../syntax/print.js'
import { areasOf } from '../keep/refs.js'
import { attrsOf, childrenOf, rootOf, type Placed } from '../keep/xml.js'
import type { WorksheetPart } from '../read/book.js'
import type { XlsxRead } from '../types.js'
import { attr, attributes, flag, within } from '../xml/scan.js'
import { readText } from '../xml/text.js'
import {
  MAX_PAGE_FIELD, type PrintBreak, type PrintHeaderFooter, type PrintMargins, type PrintOptions, type PrintReadIssue,
  type PrintSetup, type PrintSheetDefaults, type PrintSpan, type SheetPrint,
} from './types.js'
import { BREAK, HEADER_ATTRS, HEADER_TEXT, MARGINS, OPTIONS, PAGE_FIELD_WRAP, PROPERTIES, SETUP, readFields } from './xml.js'

const PRINT_AREA = '_xlnm.print_area'
const PRINT_TITLES = '_xlnm.print_titles'
type PrintName = typeof PRINT_AREA | typeof PRINT_TITLES
type Name = { name: PrintName; ref: string }

function workbookPrintNames(xml: string): Map<number, Name[]> {
  const names = new Map<number, Name[]>()
  for (const node of within(xml, 'definedNames', 'definedName')) {
    const attrs = attributes(node.attrs)
    const name = readText(attrs.name ?? '').toLowerCase()
    if (name !== PRINT_AREA && name !== PRINT_TITLES) continue
    const local = attrs.localSheetId
    if (local === undefined || !/^\d+$/.test(local)) continue
    const index = Number(local)
    const entry: Name = { name, ref: readText(node.body ?? '') }
    const list = names.get(index)
    if (list) list.push(entry)
    else names.set(index, [entry])
  }
  return names
}

function readPrintName(settings: SheetPrint, fileName: string, { name, ref }: Name): boolean {
  const areas = areasOf(ref)
  if (!areas?.length || areas.some((area) => !('rect' in area) || sheetKey(area.sheet) !== sheetKey(fileName))) return false
  const rects = areas.flatMap((area) => 'rect' in area ? [area.rect] : [])
  if (name === PRINT_AREA) {
    if (settings.areas) return false
    settings.areas = rects.map(areaText)
    return true
  }
  if (settings.repeatRows || settings.repeatColumns) return false
  let repeatRows: SheetPrint['repeatRows']
  let repeatColumns: SheetPrint['repeatColumns']
  for (const rect of rects) {
    if (!repeatRows && rect.c1 === 0 && rect.c2 === MAX_COLS - 1) repeatRows = { first: rect.r1 + 1, last: rect.r2 + 1 }
    else if (!repeatColumns && rect.r1 === 0 && rect.r2 === MAX_ROWS - 1) repeatColumns = { first: rect.c1 + 1, last: rect.c2 + 1 }
    else return false
  }
  if (repeatRows) settings.repeatRows = repeatRows
  if (repeatColumns) settings.repeatColumns = repeatColumns
  return true
}

function breaks(xml: string, node: Placed, limit: number, across: number, invalid: () => void): PrintBreak[] {
  const out: PrintBreak[] = []
  for (const child of childrenOf(xml, node)) {
    if (child.local !== 'brk') continue
    const data = readFields<{ id?: number; min?: number; max?: number; man?: boolean; pt?: boolean }>(BREAK, attrsOf(xml, child), invalid)
    const { id, min, max, man, pt } = data
    if (id === undefined || id >= limit || (max !== undefined && (max >= across || max < (min ?? 0))) || (min ?? 0) >= across) {
      invalid()
      continue
    }
    out.push({
      before: id + 1,
      ...(min === undefined ? {} : { first: min + 1 }),
      ...(max === undefined || max === across - 1 ? {} : { last: max + 1 }),
      ...(man === undefined ? {} : { manual: man }),
      ...(pt === undefined ? {} : { pivot: pt }),
    })
  }
  return out
}

function headerFooter(xml: string, node: Placed, invalid: () => void): PrintHeaderFooter {
  const header = readFields<PrintHeaderFooter>(HEADER_ATTRS, attrsOf(xml, node), invalid)
  for (const child of childrenOf(xml, node)) {
    const key = HEADER_TEXT.find((candidate) => candidate === child.local)
    if (key) header[key] = readText(child.end === child.open ? '' : xml.slice(child.open, xml.lastIndexOf('</', child.end - 1)))
  }
  return header
}

/** When zeroHeight is set, a row only takes space if the file wrote it and left it visible. */
function visibleRowSpans(xml: string, data: Placed | undefined, invalid: () => void): PrintSpan[] {
  const spans: PrintSpan[] = []
  if (!data) return spans
  let previous = 0
  for (const node of childrenOf(xml, data)) {
    if (node.local !== 'row') continue
    const attrs = attrsOf(xml, node)
    const row = attrs.r === undefined ? previous + 1 : Number(attrs.r)
    if (!Number.isInteger(row) || row <= previous || row > MAX_ROWS) {
      invalid()
      continue
    }
    previous = row
    if (flag(attrs.hidden) || (attrs.ht !== undefined && Number(attrs.ht) === 0)) continue
    const last = spans.at(-1)
    if (last?.last === row - 1) last.last = row
    else spans.push({ first: row, last: row })
  }
  return spans
}

function worksheetDefaults(xml: string, nodes: Placed[], sheet: string, invalid: () => void): PrintSheetDefaults | undefined {
  const format = nodes.find((node) => node.local === 'sheetFormatPr')
  if (!format) return undefined
  const attrs = attrsOf(xml, format)
  const defaults: PrintSheetDefaults = { sheet }
  if (attrs.defaultRowHeight !== undefined) {
    const height = Number(attrs.defaultRowHeight)
    if (attrs.defaultRowHeight.trim() && Number.isFinite(height) && height >= 0 && height <= MAX_ROW_HEIGHT) defaults.defaultRowHeight = height
    else invalid()
  }
  if (flag(attrs.zeroHeight)) {
    const data = nodes.find((node) => node.local === 'sheetData')
    defaults.visibleRows = visibleRowSpans(xml, data, invalid)
  }
  return defaults
}

export interface XlsxPrintRead extends XlsxRead {
  print: SheetPrint[]
  printIssues: PrintReadIssue[]
  sheetDefaults: PrintSheetDefaults[]
}

export function printReader() {
  const print: SheetPrint[] = []
  const sheetDefaults: PrintSheetDefaults[] = []
  const issues: PrintReadIssue[] = []
  let names: ReturnType<typeof workbookPrintNames> | undefined

  function readWorksheet(part: WorksheetPart) {
    const settings: SheetPrint = { sheet: part.name }
    const { xml } = part
    const report = (element: string, reason: PrintReadIssue['reason'] = 'invalid') => {
      if (!issues.some((issue) => issue.sheet === part.name && issue.part === element && issue.reason === reason)) {
        issues.push({ sheet: part.name, part: element, reason })
      }
    }
    names ??= workbookPrintNames(part.workbook)
    for (const name of names.get(part.index) ?? []) {
      if (!readPrintName(settings, part.fileName, name)) report(name.name)
    }
    const root = rootOf(xml)
    if (!root) return
    const setup: PrintSetup = {}
    const nodes = childrenOf(xml, root)
    const defaults = worksheetDefaults(xml, nodes, part.name, () => report('sheetFormatPr'))
    if (defaults) sheetDefaults.push(defaults)
    for (const node of nodes) {
      const invalid = () => report(node.local)
      switch (node.local) {
        case 'sheetPr': {
          const page = childrenOf(xml, node).find((child) => child.local === 'pageSetUpPr')
          if (page) Object.assign(setup, readFields(PROPERTIES, attrsOf(xml, page), () => report(page.local)))
          break
        }
        case 'pageSetup': {
          const attrs = attrsOf(xml, node)
          const page = readFields<PrintSetup>(SETUP, attrs, invalid)
          if (page.firstPageNumber !== undefined && page.firstPageNumber > MAX_PAGE_FIELD) page.firstPageNumber -= PAGE_FIELD_WRAP
          Object.assign(setup, page)
          if (attr(attrs, 'id') !== undefined) report(node.local, 'printerSettings')
          break
        }
        case 'pageMargins': {
          const margins = readFields<PrintMargins>(MARGINS, attrsOf(xml, node), invalid)
          if (Object.keys(margins).length === Object.keys(MARGINS).length) settings.margins = margins
          else invalid()
          break
        }
        case 'printOptions': {
          const options = readFields<PrintOptions>(OPTIONS, attrsOf(xml, node), invalid)
          if (Object.keys(options).length) settings.options = options
          break
        }
        case 'headerFooter':
          settings.headerFooter = headerFooter(xml, node, invalid)
          break
        case 'rowBreaks':
          settings.rowBreaks = breaks(xml, node, MAX_ROWS, MAX_COLS, invalid)
          break
        case 'colBreaks':
          settings.columnBreaks = breaks(xml, node, MAX_COLS, MAX_ROWS, invalid)
          break
        case 'legacyDrawingHF':
        case 'drawingHF':
          report(node.local, 'headerFooterDrawing')
          break
      }
    }
    if (Object.keys(setup).length) settings.setup = setup
    if (Object.keys(settings).length > 1) print.push(settings)
  }

  const result = (read: XlsxRead): XlsxPrintRead => ({ ...read, print, printIssues: issues, sheetDefaults })
  return { readWorksheet, result }
}
