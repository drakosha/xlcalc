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

/** Print flags of a worksheet: centering on the page, row and column headings, gridlines. */
export interface PrintOptions {
  horizontalCentered?: boolean
  verticalCentered?: boolean
  headings?: boolean
  gridLines?: boolean
  gridLinesSet?: boolean
}

/** A worksheet's headers and footers: odd pages, and even and first pages when their flags are set. */
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

/** A manual or automatic page break, before a row or a column. */
export interface PrintBreak {
  /** The one-based row or column immediately after this break. */
  before: number
  /** Inclusive bounds along the other axis. Omit for a break across the entire print area. */
  first?: number
  last?: number
  manual?: boolean
  pivot?: boolean
}

/** A worksheet's print settings: areas, repeated titles, page setup, margins, flags, headers and footers, breaks. */
export interface WorksheetPrintSettings {
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

/** The saved default row geometry of a worksheet. */
export interface SheetRowDefaults {
  /** Default height of a stored worksheet's rows, in points. */
  defaultRowHeight?: number
  /** When the file hides rows by default, the spans it writes as visible. */
  visibleRows?: readonly PrintSpan[]
}
