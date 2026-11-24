import type { CellValue } from '../../contract.js'
import type { ChartDrawing, DrawingReadIssue } from '../drawings/types.js'

/** A cached chart point: its index in the source (indexes can skip) and value. */
export interface ChartPoint {
  index: number
  value: CellValue
}

/** A chart's data source: its reference, point count, cached points and number format. */
export interface ChartSource {
  /** The chart's stored formula, for resolving its points against the current workbook. */
  reference?: string
  count: number
  points: ChartPoint[]
  /** Levels of a cached hierarchical category axis, in file order. */
  levels?: ChartPoint[][]
  format?: string
}

/** A chart element's basic fill and line as saved: colors as `#RRGGBB` or `none`, width in points. */
export interface ChartStyle {
  fill?: string
  stroke?: string
  strokeWidth?: number
}

/** A chart series: name, categories and values (or x and y for scatter), style and marker. */
export interface ChartSeries {
  name?: ChartSource
  categories?: ChartSource
  values?: ChartSource
  x?: ChartSource
  y?: ChartSource
  style?: ChartStyle
  marker?: { symbol?: string; size?: number }
}

/** A group of series of one chart type over a pair of axes. */
export interface ChartGroup {
  type: 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter'
  axes: string[]
  series: ChartSeries[]
  direction?: string
  grouping?: string
  scatterStyle?: string
  gapWidth?: number
  holeSize?: number
}

/** A chart axis: kind, id and crossing axis, title, bounds, number format and visibility. */
export interface ChartAxis {
  kind: 'category' | 'date' | 'value' | 'series'
  id: string
  crossesAxis?: string
  position?: string
  title?: ChartSource
  min?: number
  max?: number
  logBase?: number
  majorUnit?: number
  format?: string
  reversed: boolean
  hidden: boolean
  majorGridlines: boolean
}

/** A chart read from ChartML, joined to its frame from `readXlsxDrawings` by `frame.id`. */
export interface XlsxChart {
  frame: ChartDrawing
  title?: ChartSource
  legend?: { position?: string; overlay: boolean }
  groups: ChartGroup[]
  axes: ChartAxis[]
  plotVisibleOnly: boolean
  emptyCells?: 'gap' | 'span' | 'zero'
  style?: ChartStyle
}

/** A file's charts and the chart parts the reader skipped. */
export interface XlsxCharts {
  charts: XlsxChart[]
  issues: DrawingReadIssue[]
}

/** Current values for a chart's sources: a session, or a host reader that also resolves names. */
export interface ChartValueReader {
  /** A session satisfies this interface directly. */
  getValues(sheet: string, range: string): readonly (readonly CellValue[])[]
  /** Optionally evaluate a formula or name with the host's workbook context. Undefined keeps the file's cached points. */
  resolve?(reference: string, sheet: string): readonly CellValue[] | undefined
}
