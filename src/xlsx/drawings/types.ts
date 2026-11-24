import type { PrintRect, PrintSize } from '../print/types.js'

/** One-based worksheet lines, with offsets in points from the upper-left corner of the cell. */
export interface DrawingMarker {
  row: number
  col: number
  x: number
  y: number
}

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

export interface NoteDrawing extends DrawingFrame {
  kind: 'note'
  /** Cell whose note text the workbook supplies. */
  cell: string
  fill?: string
  stroke?: string
}

export interface ChartDrawing extends DrawingFrame {
  kind: 'chart'
  part: string
}

export type XlsxDrawing = PictureDrawing | NoteDrawing | ChartDrawing

export interface DrawingReadIssue {
  sheet: string
  part: string
  reason: 'invalidAnchor' | 'missingPart' | 'externalImage' | 'unsupportedDrawing' | 'unsupportedChart'
}

export interface XlsxDrawings {
  drawings: XlsxDrawing[]
  issues: DrawingReadIssue[]
  /** Load a reported picture's part within the archive limits. The returned bytes are shared. */
  readImage(part: string): Promise<Uint8Array | undefined>
}
