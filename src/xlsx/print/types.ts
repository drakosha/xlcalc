import type { BookInput, CellNote } from '../../contract.js'

/** Excel's bounds for page counts and explicit first page numbers. */
export const MAX_PAGE_FIELD = 32767

/** A span of sheet lines, inclusive and one-based, including any hidden lines. */
export interface PrintSpan {
  first: number
  last: number
}

/** Physical page settings stored by the worksheet. Paper size is a SpreadsheetML paper-size code. */
export interface PrintSetup {
  autoPageBreaks?: boolean
  fitToPage?: boolean
  paperSize?: number
  paperHeight?: string
  paperWidth?: string
  scale?: number
  fitToWidth?: number
  fitToHeight?: number
  pageOrder?: 'downThenOver' | 'overThenDown'
  orientation?: 'default' | 'portrait' | 'landscape'
  usePrinterDefaults?: boolean
  blackAndWhite?: boolean
  draft?: boolean
  cellComments?: 'none' | 'asDisplayed' | 'atEnd'
  errors?: 'displayed' | 'blank' | 'dash' | 'NA'
  firstPageNumber?: number
  useFirstPageNumber?: boolean
  horizontalDpi?: number
  verticalDpi?: number
  copies?: number
}

/** Margins in inches, as saved in an XLSX worksheet. */
export interface PrintMargins {
  left: number
  right: number
  top: number
  bottom: number
  header: number
  footer: number
}

export interface PrintOptions {
  horizontalCentered?: boolean
  verticalCentered?: boolean
  headings?: boolean
  gridLines?: boolean
  gridLinesSet?: boolean
}

export interface PrintHeaderFooter {
  differentOddEven?: boolean
  differentFirst?: boolean
  scaleWithDoc?: boolean
  alignWithMargins?: boolean
  /** Header or footer text in Excel's syntax, e.g. `&LReport&RPage &P of &N`. */
  oddHeader?: string
  oddFooter?: string
  evenHeader?: string
  evenFooter?: string
  firstHeader?: string
  firstFooter?: string
}

export interface PrintBreak {
  /** The one-based row or column immediately after this break. */
  before: number
  /** Inclusive bounds along the other axis. Omit for a break across the entire print area. */
  first?: number
  last?: number
  manual?: boolean
  pivot?: boolean
}

export interface SheetPrint {
  sheet: string
  /** An ordered list of sheet-local A1 ranges, one print area per entry. */
  areas?: readonly string[]
  repeatRows?: PrintSpan
  repeatColumns?: PrintSpan
  setup?: PrintSetup
  margins?: PrintMargins
  options?: PrintOptions
  headerFooter?: PrintHeaderFooter
  rowBreaks?: readonly PrintBreak[]
  columnBreaks?: readonly PrintBreak[]
}

export interface PrintReadIssue {
  sheet: string
  part: string
  reason: 'invalid' | 'printerSettings' | 'headerFooterDrawing'
}

export interface PrintSize {
  width: number
  height: number
}

export interface PrintRect extends PrintSize {
  x: number
  y: number
}

/** One worksheet axis in points, including zero-width hidden lines. */
export interface PrintAxisGeometry {
  count: number
  /** Position from the sheet origin of the edge before line 1 through count + 1. */
  edge(line: number): number
}

export interface PrintSheetDefaults {
  sheet: string
  /** Default height of a stored worksheet's rows, in points. */
  defaultRowHeight?: number
  /** When the file hides rows by default, the spans it writes as visible. */
  visibleRows?: readonly PrintSpan[]
}

export interface PrintSheetGeometry {
  rows: PrintAxisGeometry
  columns: PrintAxisGeometry
  /** Extent of printable content, including objects when supplied by the host. */
  usedRange?: string
}

export interface PrintDeviceGeometry {
  /** Paper points before a sheet row, from row 1 through row count + 1, at the requested scale factor. */
  rowEdge?: (line: number, scale: number) => number
  /** Paper points before a sheet column, from column 1 through column count + 1, at the requested scale factor. */
  columnEdge?: (line: number, scale: number) => number
}

/** The effective, oriented paper size in points and the device’s printable geometry. */
export interface PrintTarget extends PrintSize {
  /** Unprintable edges of that paper, in points. */
  insets?: Partial<Pick<PrintMargins, 'left' | 'right' | 'top' | 'bottom'>>
  /** When printing row and column headings, their unscaled width and height in points. */
  headings?: PrintSize
  /** The device's printable grid across its supported scale factors. Missing axes use scaled sheet points. */
  deviceGeometry?: PrintDeviceGeometry
  /** Extra trailing space the print device reserves when fitting complete lines into the body, in points. */
  paginationReserve?: Partial<PrintSize>
  /** The maximum number of pages to materialize (10,000 by default). */
  maxPages?: number
}

/** A rectangular portion of the sheet mapped to the page. Rectangles are in points from their respective origins. */
export interface PrintRegion {
  range: string
  source: PrintRect
  target: PrintRect
}

export interface PrintPageBand {
  /** The selected text in Excel's header/footer syntax. */
  text: string
  bounds: PrintRect
  scale: number
}

/** Text measurements for end notes at the current print scale, in paper points. */
export interface PrintNoteMetrics {
  lineHeight: number
  measureText(text: string): number
}

export interface PrintNotes {
  /** Cells keyed by address, such as the cells returned by readXlsxWithPrint. */
  cells: Readonly<Record<string, { note?: CellNote }>>
  /** The workbook's Normal font. The default measurements use Calibri 11 when omitted. */
  font?: BookInput['font']
  labels?: { address: string; body: string }
  /** Override nominal measurements with the renderer's font and device metrics. */
  metrics?: (scale: number) => PrintNoteMetrics
}

export interface PrintNoteLine {
  cell: string
  kind: 'address' | 'body'
  text: string
  bounds: PrintRect
  label?: { text: string; bounds: PrintRect }
}

export interface PrintPage {
  number: number
  /** The one-based print area index, or zero on a sheet-wide page of notes. */
  area: number
  scale: number
  paper: PrintSize
  /** The printable body after page margins. Regions are clipped to this rectangle. */
  body: PrintRect
  regions: PrintRegion[]
  /** Positioned text lines on an end-notes page. */
  noteLines?: PrintNoteLine[]
  header?: PrintPageBand
  footer?: PrintPageBand
}

export interface PrintLayoutIssue {
  /** One-based print area index. */
  area: number
  reason: 'fitLimit' | 'clippedLine'
}

export interface PrintLayout {
  pages: PrintPage[]
  issues: PrintLayoutIssue[]
}

export interface PrintWorkbookSheet {
  settings: SheetPrint
  geometry: PrintSheetGeometry
  target: PrintTarget
  notes?: PrintNotes
  hidden?: boolean | 'very'
}

export interface PrintWorkbookOptions {
  /** Maximum number of pages in the job before selecting output pages (10,000 by default). */
  maxPages?: number
  /** Inclusive interval of physical pages to output, from one. */
  fromPage?: number
  toPage?: number
}

export interface PrintWorkbookPage extends PrintPage {
  sheet: string
  /** Physical page in the job, from one. The page number field may be restarted by the sheet. */
  position: number
}

export interface PrintWorkbookLayout {
  pages: PrintWorkbookPage[]
  /** Full job page count, including pages outside the selected interval. */
  totalPages: number
  issues: (PrintLayoutIssue & { sheet: string })[]
}
