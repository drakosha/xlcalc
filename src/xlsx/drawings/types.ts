import type { PrintRect, PrintSize } from '../print/types.js'

/** One-based worksheet lines, with offsets in points from the upper-left corner of the cell. */
export interface DrawingMarker {
  row: number
  col: number
  x: number
  y: number
}

/** Where a drawing sits: between two cells, at a cell with a size, or at an absolute position in points. */
export type DrawingAnchor =
  | { kind: 'twoCell'; from: DrawingMarker; to: DrawingMarker; editAs: 'twoCell' | 'oneCell' | 'absolute' }
  | { kind: 'oneCell'; from: DrawingMarker; size: PrintSize }
  | { kind: 'absolute'; bounds: PrintRect }

interface DrawingFrame {
  /** The sheet part, drawing part and position, stable between readers of the same file. */
  id: string
  sheet: string
  sheetKind: 'worksheet' | 'chartsheet'
  anchor: DrawingAnchor
  name?: string
  hidden: boolean
  printable: boolean
}

/** An embedded picture over the grid, its part loaded by `readImage`. */
export interface PictureDrawing extends DrawingFrame {
  kind: 'picture'
  part: string
  contentType?: string
  alt?: string
  /** Clockwise degrees around the center of the picture. */
  rotation?: number
  flipX?: boolean
  flipY?: boolean
  /** Fractions trimmed from each edge of the original image. */
  crop?: { left: number; top: number; right: number; bottom: number }
}

/** A note's box from the VML drawing; its text is the cell's `note`. */
export interface NoteDrawing extends DrawingFrame {
  kind: 'note'
  /** Cell whose note text the workbook supplies. */
  cell: string
  fill?: string
  stroke?: string
}

/** A chart frame, its ChartML part read by `readXlsxCharts`. */
export interface ChartDrawing extends DrawingFrame {
  kind: 'chart'
  part: string
}

/** A picture, note box or chart frame of a sheet. */
export type XlsxDrawing = PictureDrawing | NoteDrawing | ChartDrawing

/** A drawing the reader skipped: a bad anchor, a missing or external part, an unsupported drawing or chart. */
export interface DrawingReadIssue {
  sheet: string
  part: string
  reason: 'invalidAnchor' | 'missingPart' | 'externalImage' | 'unsupportedDrawing' | 'unsupportedChart'
}

/** A file's drawings in drawing order, the skipped ones, and the loader of picture bytes. */
export interface XlsxDrawings {
  drawings: XlsxDrawing[]
  issues: DrawingReadIssue[]
  /** Load a reported picture's part within the archive limits. The returned bytes are shared. */
  readImage(part: string): Promise<Uint8Array | undefined>
}
