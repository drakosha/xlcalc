import type { CellValue } from '../../contract.js'
import type { ChartDrawing, DrawingReadIssue } from '../drawings/types.js'

export interface ChartPoint {
  index: number
  value: CellValue
}

export interface ChartSource {
  /** The chart's stored formula, for resolving its points against the current workbook. */
  reference?: string
  count: number
  points: ChartPoint[]
  /** Levels of a cached hierarchical category axis, in file order. */
  levels?: ChartPoint[][]
  format?: string
}

export interface ChartStyle {
  fill?: string
  stroke?: string
  strokeWidth?: number
}

export interface ChartSeries {
  name?: ChartSource
  categories?: ChartSource
  values?: ChartSource
  x?: ChartSource
  y?: ChartSource
  style?: ChartStyle
  marker?: { symbol?: string; size?: number }
}

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

export interface XlsxCharts {
  charts: XlsxChart[]
  issues: DrawingReadIssue[]
}

export interface ChartValueReader {
  /** A session satisfies this interface directly. */
  getValues(sheet: string, range: string): readonly (readonly CellValue[])[]
  /** Optionally evaluate a formula or name with the host's workbook context. Undefined keeps the file's cached points. */
  resolve?(reference: string, sheet: string): readonly CellValue[] | undefined
}
