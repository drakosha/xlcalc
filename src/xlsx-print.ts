import { strict } from './strict-args.js'
import { readXlsxDrawings as readDrawings } from './xlsx/drawings/read.js'
import { printDrawingBounds as drawingBounds, projectPrintDrawings as projectDrawings } from './xlsx/print/drawings.js'
import { printSheetGeometry as geometry } from './xlsx/print/geometry.js'
import { formatPrintHeaderFooter as formatHeaderFooter } from './xlsx/print/header-footer.js'
import { layoutPrintPages as layout, layoutPrintWorkbook as workbookLayout } from './xlsx/print/layout.js'
import { printPaperSize as paperSize } from './xlsx/print/paper.js'
import { projectPrintRect as project } from './xlsx/print/project.js'

export type {
  PrintHeaderFooterContext, PrintHeaderFooterRun, PrintHeaderFooterSections, PrintHeaderFooterStyle,
} from './xlsx/print/header-footer.js'
export type { ChartDrawing, DrawingAnchor, DrawingMarker, DrawingReadIssue, NoteDrawing, PictureDrawing, XlsxDrawing, XlsxDrawings } from './xlsx/drawings/types.js'
export type { PrintDrawingOptions, PrintDrawingProjection } from './xlsx/print/drawings.js'
export type { PrintRectProjection } from './xlsx/print/project.js'
export type { PrintGeometryOptions } from './xlsx/print/geometry.js'
export type {
  PrintAxisGeometry, PrintBreak, PrintDeviceGeometry, PrintHeaderFooter, PrintLayout, PrintLayoutIssue, PrintMargins, PrintOptions, PrintPage,
  PrintNoteLine, PrintNoteMetrics, PrintNotes, PrintPageBand, PrintReadIssue, PrintRect,
  PrintRegion, PrintSetup, SheetRowDefaults, WorksheetPrintSettings, PrintSheetGeometry, PrintSize, PrintSpan, PrintTarget,
  PrintWorkbookLayout, PrintWorkbookOptions, PrintWorkbookPage, PrintWorkbookSheet, SheetPrint,
} from './xlsx/print/types.js'


export const printSheetGeometry = /* @__PURE__ */ strict(geometry, 'printSheetGeometry', 2)
export const formatPrintHeaderFooter = /* @__PURE__ */ strict(formatHeaderFooter)
export const printPaperSize = /* @__PURE__ */ strict(paperSize)
export const readXlsxDrawings = /* @__PURE__ */ strict(readDrawings)
export const printDrawingBounds = /* @__PURE__ */ strict(drawingBounds)
export const projectPrintDrawings = /* @__PURE__ */ strict(projectDrawings, 'projectPrintDrawings', 4)
export const projectPrintRect = /* @__PURE__ */ strict(project)
export const layoutPrintPages = /* @__PURE__ */ strict(layout)
export const layoutPrintWorkbook = /* @__PURE__ */ strict(workbookLayout, 'layoutPrintWorkbook', 2)
