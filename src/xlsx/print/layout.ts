/** Print areas laid out independently, with a common sheet scale and page coordinates in points. */
import { checkNumber, integer } from '../../input/basic.js'
import { InputError } from '../../input-error.js'
import { MAX_COLS, MAX_ROWS } from '../../model.js'
import { parseRange, rectText } from '../../structure/area.js'
import { noteCells, notePages } from './notes.js'
import { MAX_PAGE_FIELD } from './types.js'
import type {
  PrintAxisGeometry, PrintBreak, PrintDeviceGeometry, PrintLayout, PrintLayoutIssue, PrintMargins, PrintNotes, PrintPage, PrintPageBand, PrintRect,
  PrintRegion, PrintSheetGeometry, PrintSpan, PrintTarget, PrintWorkbookLayout, PrintWorkbookOptions, PrintWorkbookPage,
  PrintWorkbookSheet, SheetPrint,
} from './types.js'

const POINTS_PER_INCH = 72
const DEFAULT_MARGINS: PrintMargins = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }
const EPSILON = 1e-8
const MAX_PAGES = 10_000

interface Axis {
  geometry: PrintAxisGeometry
  deviceEdge?: PrintDeviceGeometry['rowEdge']
  range: PrintSpan
  titles?: PrintSpan
  breaks: readonly number[]
}

interface Slice {
  bands: PrintSpan[]
  size: number
  clipped: boolean
}

interface Area {
  rows: Axis
  columns: Axis
  index: number
}

type Edge = PrintAxisGeometry['edge']
const length = (edge: Edge, span: PrintSpan) => edge(span.last + 1) - edge(span.first)
function scaledEdge(axis: Axis, scale: number): Edge {
  const device = axis.deviceEdge
  return device ? (line) => device(line, scale) : (line) => axis.geometry.edge(line) * scale
}

function titleBands(body: PrintSpan, titles: PrintSpan | undefined): PrintSpan[] {
  if (!titles || body.first <= titles.first) return [body]
  if (titles.last + 1 < body.first) return [titles, body]
  return [{ first: titles.first, last: Math.max(body.last, titles.last) }]
}

function measured(edge: Edge, titles: PrintSpan | undefined, first: number, last: number) {
  const bands = titleBands({ first, last }, titles)
  return { bands, size: bands.reduce((sum, band) => sum + length(edge, band), 0) }
}

function* slices(axis: Axis, available: number, scale: number): Generator<Slice> {
  const edge = scaledEdge(axis, scale)
  let first = axis.range.first
  for (const breakAt of [...axis.breaks, axis.range.last + 1]) {
    const end = breakAt - 1
    while (first <= end && length(edge, { first, last: end }) > 0) {
      // End positions are monotonic, including hidden lines whose edges are identical.
      let low = first
      let high = end
      while (low < high) {
        const next = Math.ceil((low + high) / 2)
        if (measured(edge, axis.titles, first, next).size <= available + EPSILON) low = next
        else high = next - 1
      }
      let last = low
      if (length(edge, { first, last }) === 0) {
        // Repeated titles can leave too little room for even one visible line.
        // Absorb all preceding hidden lines into the clipped page.
        let visible = end
        while (last < visible) {
          const middle = Math.floor((last + visible) / 2)
          if (length(edge, { first, last: middle }) > 0) visible = middle
          else last = middle + 1
        }
        last = visible
      }
      const got = measured(edge, axis.titles, first, last)
      yield { ...got, clipped: got.size > available + EPSILON }
      first = last + 1
    }
    first = breakAt
  }
}

function matchesFit(axis: Axis, available: number, count: number, scale: number): boolean {
  if (!count) return true
  let pages = 0
  for (const slice of slices(axis, available, scale)) {
    if (slice.clipped || ++pages > count) return false
  }
  return true
}

function fitScale(areas: Area[], width: number, height: number, settings: SheetPrint, headings: PrintTarget['headings']): number {
  const fitWidth = integer(settings.setup?.fitToWidth ?? 1, 'print.setup.fitToWidth', 0, MAX_PAGE_FIELD)
  const fitHeight = integer(settings.setup?.fitToHeight ?? 1, 'print.setup.fitToHeight', 0, MAX_PAGE_FIELD)
  const fits = (percent: number) => {
    const scale = percent / 100
    const availableWidth = width - (headings?.width ?? 0) * scale
    const availableHeight = height - (headings?.height ?? 0) * scale
    return areas.every((area) =>
      matchesFit(area.columns, availableWidth, fitWidth, scale) && matchesFit(area.rows, availableHeight, fitHeight, scale))
  }
  if (!fits(10)) return 0
  let low = 10
  let high = 100
  while (low < high) {
    const next = Math.ceil((low + high) / 2)
    if (fits(next)) low = next
    else high = next - 1
  }
  return low / 100
}

function geometryCheck(geometry: PrintAxisGeometry, path: string, limit: number) {
  integer(geometry.count, `${path}.count`, 0, limit)
  if (typeof geometry.edge !== 'function' || geometry.edge(1) !== 0) {
    throw new InputError(`${path}.edge`, 'expected sheet coordinates starting at zero')
  }
  checkNumber(geometry.edge(geometry.count + 1), `${path}.edge(count + 1)`, 0, Number.MAX_SAFE_INTEGER)
}

function titleSpan(span: PrintSpan | undefined, geometry: PrintAxisGeometry, key: string): PrintSpan | undefined {
  if (!span) return undefined
  const first = integer(span.first, `print.${key}.first`, 1, geometry.count)
  const last = integer(span.last, `print.${key}.last`, first, geometry.count)
  return { first, last }
}

function manualCuts(breaks: readonly PrintBreak[] | undefined, axis: PrintSpan, across: PrintSpan): number[] {
  const cuts = new Set<number>()
  for (const item of breaks ?? []) {
    if (!item.manual || item.before <= axis.first || item.before > axis.last) continue
    const first = item.first ?? 1
    const last = item.last ?? Infinity
    // Excel makes a cut across this entire print area when the break overlaps any part of it.
    if (first <= across.last && last >= across.first) cuts.add(item.before)
  }
  return [...cuts].sort((a, b) => a - b)
}

function printAreas(
  settings: SheetPrint, geometry: PrintSheetGeometry, device: PrintDeviceGeometry | undefined,
): Area[] {
  const { rows, columns } = geometry
  const used = geometry.usedRange === undefined ? undefined : parseRange(geometry.usedRange)
  if (geometry.usedRange !== undefined && !used) throw new InputError('geometry.usedRange', 'expected a sheet-local A1 range')
  const ranges = settings.areas?.length ? settings.areas : geometry.usedRange ? [geometry.usedRange] : []
  const repeatRows = titleSpan(settings.repeatRows, rows, 'repeatRows')
  const repeatColumns = titleSpan(settings.repeatColumns, columns, 'repeatColumns')
  return ranges.flatMap((range, i) => {
    const path = `print.areas[${i}]`
    const area = parseRange(range.replace(/\$/g, '').toUpperCase())
    if (!area) throw new InputError(path, 'expected a sheet-local A1 range')
    const rowSpan = { first: area.r1 + 1, last: area.r2 + 1 }
    const colSpan = { first: area.c1 + 1, last: area.c2 + 1 }
    // A whole-row or whole-column print area uses the supplied extent along its open axis.
    if (area.r1 === 0 && rowSpan.last === MAX_ROWS) rowSpan.last = used ? used.r2 + 1 : rows.count
    if (area.c1 === 0 && colSpan.last === MAX_COLS) colSpan.last = used ? used.c2 + 1 : columns.count
    if (rowSpan.last > rows.count || colSpan.last > columns.count) throw new InputError(path, 'print area extends beyond the supplied sheet geometry')
    if (rowSpan.last < rowSpan.first || colSpan.last < colSpan.first) return []
    const index = i + 1
    return [{
      index,
      rows: { geometry: rows, deviceEdge: device?.rowEdge, range: rowSpan, titles: repeatRows, breaks: settings.setup?.fitToPage ? [] : manualCuts(settings.rowBreaks, rowSpan, colSpan) },
      columns: { geometry: columns, deviceEdge: device?.columnEdge, range: colSpan, titles: repeatColumns, breaks: settings.setup?.fitToPage ? [] : manualCuts(settings.columnBreaks, colSpan, rowSpan) },
    }]
  })
}

function marginsAt(settings: SheetPrint, target: PrintTarget): { body: PrintRect; margins: PrintMargins } {
  const margins = settings.margins ?? DEFAULT_MARGINS
  const inset = (side: 'top' | 'bottom' | 'left' | 'right') => {
    const printer = target.insets?.[side] ?? 0
    checkNumber(printer, `target.insets.${side}`, 0, Number.MAX_SAFE_INTEGER)
    return Math.max(printer, margins[side] * POINTS_PER_INCH)
  }
  for (const [side, size] of Object.entries(margins)) checkNumber(size, `print.margins.${side}`, 0, Number.MAX_SAFE_INTEGER)
  const x = inset('left')
  const y = inset('top')
  const width = target.width - x - inset('right')
  const height = target.height - y - inset('bottom')
  if (!(width > 0 && height > 0)) throw new InputError('target', 'paper has no body after margins and printer insets')
  return { body: { x, y, width, height }, margins }
}

function pageRegions(area: Area, rows: Slice, columns: Slice, x: number, y: number, scale: number): PrintRegion[] {
  const regions: PrintRegion[] = []
  const sheetRows = (line: number) => area.rows.geometry.edge(line)
  const sheetColumns = (line: number) => area.columns.geometry.edge(line)
  const printRows = scaledEdge(area.rows, scale)
  const printColumns = scaledEdge(area.columns, scale)
  for (const row of rows.bands) {
    const height = length(printRows, row)
    if (!height) continue
    let left = x
    for (const col of columns.bands) {
      const width = length(printColumns, col)
      if (!width) continue
      const range = rectText({ r1: row.first - 1, r2: row.last - 1, c1: col.first - 1, c2: col.last - 1 })
      const source = { x: sheetColumns(col.first), y: sheetRows(row.first), width: length(sheetColumns, col), height: length(sheetRows, row) }
      const target = { x: left, y, width, height }
      regions.push({ range, source, target })
      left += target.width
    }
    y += height
  }
  return regions
}

function headerFooter(settings: SheetPrint, page: PrintPage, target: PrintTarget, margins: PrintMargins, pageIndex: number) {
  const given = settings.headerFooter
  if (!given) return
  const variant = given.differentFirst && pageIndex === 0 ? 'first' : given.differentOddEven && pageIndex % 2 === 1 ? 'even' : 'odd'
  const horizontal = given.alignWithMargins === false
    ? { x: target.insets?.left ?? 0, width: target.width - (target.insets?.left ?? 0) - (target.insets?.right ?? 0) }
    : { x: page.body.x, width: page.body.width }
  const headerTop = Math.max(target.insets?.top ?? 0, margins.header * POINTS_PER_INCH)
  const footerEnd = target.height - Math.max(target.insets?.bottom ?? 0, margins.footer * POINTS_PER_INCH)
  const band = (text: string | undefined, y: number, end: number): PrintPageBand | undefined => text === undefined ? undefined : {
    text,
    bounds: { ...horizontal, y, height: Math.max(0, end - y) },
    scale: given.scaleWithDoc === false ? 1 : page.scale,
  }
  const bodyEnd = page.body.y + page.body.height
  const header = band(given[`${variant}Header`], headerTop, page.body.y)
  const footer = band(given[`${variant}Footer`], bodyEnd, footerEnd)
  if (header) page.header = header
  if (footer) page.footer = footer
}

/** Page regions map sheet points to paper points, independently of the host's cell and drawing renderer. */
export function layoutPrintPages(
  settings: SheetPrint, geometry: PrintSheetGeometry, target: PrintTarget, notes?: PrintNotes,
): PrintLayout {
  return layoutSheetPages(settings, geometry, target, 0, 1, notes)
}

function layoutSheetPages(
  settings: SheetPrint, geometry: PrintSheetGeometry, target: PrintTarget, precedingPages: number, automaticNumber: number, notes?: PrintNotes,
): PrintLayout {
  geometryCheck(geometry.rows, 'geometry.rows', MAX_ROWS)
  geometryCheck(geometry.columns, 'geometry.columns', MAX_COLS)
  checkNumber(target.width, 'target.width', EPSILON, Number.MAX_SAFE_INTEGER)
  checkNumber(target.height, 'target.height', EPSILON, Number.MAX_SAFE_INTEGER)
  const maxPages = integer(target.maxPages ?? MAX_PAGES, 'target.maxPages', 1, Number.MAX_SAFE_INTEGER)
  const issues: PrintLayoutIssue[] = []
  const report = (area: number, reason: PrintLayoutIssue['reason']) => {
    if (!issues.some((issue) => issue.area === area && issue.reason === reason)) issues.push({ area, reason })
  }
  const { body, margins } = marginsAt(settings, target)
  const areas = printAreas(settings, geometry, target.deviceGeometry)
  const headings = settings.options?.headings ? target.headings : undefined
  if (settings.options?.headings && !headings) throw new InputError('target.headings', 'expected row and column heading sizes in points')
  if (headings) {
    checkNumber(headings.width, 'target.headings.width', 0, Number.MAX_SAFE_INTEGER)
    checkNumber(headings.height, 'target.headings.height', 0, Number.MAX_SAFE_INTEGER)
  }
  const reserveWidth = target.paginationReserve?.width ?? 0
  const reserveHeight = target.paginationReserve?.height ?? 0
  checkNumber(reserveWidth, 'target.paginationReserve.width', 0, body.width - EPSILON)
  checkNumber(reserveHeight, 'target.paginationReserve.height', 0, body.height - EPSILON)
  const fitWidth = body.width - reserveWidth
  const fitHeight = body.height - reserveHeight
  const fitted = settings.setup?.fitToPage ? fitScale(areas, fitWidth, fitHeight, settings, headings) : undefined
  const scale = fitted === undefined ? integer(settings.setup?.scale || 100, 'print.setup.scale', 10, 400) / 100 : fitted || 0.1
  const headingWidth = (headings?.width ?? 0) * scale
  const headingHeight = (headings?.height ?? 0) * scale
  const availableWidth = fitWidth - headingWidth
  const availableHeight = fitHeight - headingHeight
  const firstNumber = settings.setup?.useFirstPageNumber
    ? integer(settings.setup.firstPageNumber ?? 1, 'print.setup.firstPageNumber', -MAX_PAGE_FIELD, MAX_PAGE_FIELD)
    : automaticNumber
  const pages: PrintPage[] = []
  const paper = { width: target.width, height: target.height }
  const append = (area: number, regions: PrintRegion[], noteLines?: PrintPage['noteLines']) => {
    if (pages.length === maxPages) throw new InputError('target.maxPages', 'print layout exceeds the page limit')
    const page: PrintPage = { number: firstNumber + pages.length, area, scale, paper, body, regions }
    if (noteLines) page.noteLines = noteLines
    headerFooter(settings, page, target, margins, precedingPages + pages.length)
    pages.push(page)
  }
  const endNotes = notes && settings.setup?.cellComments === 'atEnd' ? noteCells(notes) : undefined
  const notesInOrder = []

  for (const area of areas) {
    const areaNotes = endNotes?.filter(({ row, col }) =>
      row >= area.rows.range.first && row <= area.rows.range.last
      && col >= area.columns.range.first && col <= area.columns.range.last)
    for (const note of areaNotes ?? []) notesInOrder.push(note)
    if (fitted === 0) report(area.index, 'fitLimit')
    const collect = (axis: Axis, available: number) => {
      const out: Slice[] = []
      for (const slice of slices(axis, available, scale)) {
        if (out.length === maxPages) throw new InputError('target.maxPages', 'print layout exceeds the page limit')
        if (slice.clipped) report(area.index, 'clippedLine')
        out.push(slice)
      }
      return out
    }
    const rows = collect(area.rows, availableHeight)
    const columns = rows.length ? collect(area.columns, availableWidth) : []
    if (!rows.length || !columns.length) {
      // An area with end notes still has a grid page when its cells are all hidden.
      if (areaNotes?.length) append(area.index, [])
      continue
    }
    if (rows.length * columns.length > maxPages - pages.length) throw new InputError('target.maxPages', 'print layout exceeds the page limit')
    const appendGrid = (row: Slice, column: Slice) => {
      const width = column.size + headingWidth
      const height = row.size + headingHeight
      const x = body.x + (settings.options?.horizontalCentered ? Math.max(0, body.width - width) / 2 : 0)
      const y = body.y + (settings.options?.verticalCentered ? Math.max(0, body.height - height) / 2 : 0)
      const regions = pageRegions(area, row, column, x + headingWidth, y + headingHeight, scale)
      append(area.index, regions)
    }
    if (settings.setup?.pageOrder === 'overThenDown') {
      for (const row of rows) for (const column of columns) appendGrid(row, column)
    } else {
      for (const column of columns) for (const row of rows) appendGrid(row, column)
    }
  }
  if (notes && notesInOrder.length) {
    for (const lines of notePages(notesInOrder, notes, body, scale)) append(0, [], lines)
  }
  return { pages, issues }
}

/** Print the supplied visible worksheets in order, with Excel's workbook-level numbering and header variants. */
export function layoutPrintWorkbook(sheets: readonly PrintWorkbookSheet[], options: PrintWorkbookOptions = {}): PrintWorkbookLayout {
  if (!Array.isArray(sheets)) throw new InputError('sheets', 'expected an array of worksheets')
  const maxPages = integer(options.maxPages ?? MAX_PAGES, 'options.maxPages', 1, Number.MAX_SAFE_INTEGER)
  const fromPage = integer(options.fromPage ?? 1, 'options.fromPage', 1, Number.MAX_SAFE_INTEGER)
  const toPage = integer(options.toPage ?? Number.MAX_SAFE_INTEGER, 'options.toPage', fromPage, Number.MAX_SAFE_INTEGER)
  const pages: PrintWorkbookPage[] = []
  const issues: PrintWorkbookLayout['issues'] = []
  let nextNumber = 1
  for (const { settings, geometry, target, notes, hidden } of sheets as readonly PrintWorkbookSheet[]) {
    if (hidden) continue
    const remaining = maxPages - pages.length
    const sheetLimit = Math.max(1, remaining)
    const requestedLimit = integer(target.maxPages ?? sheetLimit, 'target.maxPages', 1, Number.MAX_SAFE_INTEGER)
    const sheetTarget = { ...target, maxPages: Math.min(requestedLimit, sheetLimit) }
    const plan = layoutSheetPages(settings, geometry, sheetTarget, pages.length, nextNumber, notes)
    for (const issue of plan.issues) issues.push({ ...issue, sheet: settings.sheet })
    if (!plan.pages.length) continue
    if (plan.pages.length > remaining) throw new InputError('options.maxPages', 'print layout exceeds the page limit')
    // Excel starts the first subsequent automatic sheet at the first explicit number even if the first sheet spanned pages.
    const firstNumber = plan.pages[0].number
    nextNumber = pages.length === 0 && settings.setup?.useFirstPageNumber ? firstNumber : firstNumber + plan.pages.length
    for (const page of plan.pages) pages.push({ ...page, sheet: settings.sheet, position: pages.length + 1 })
  }
  const totalPages = pages.length
  return { pages: fromPage === 1 && toPage >= totalPages ? pages : pages.slice(fromPage - 1, toPage), totalPages, issues }
}
