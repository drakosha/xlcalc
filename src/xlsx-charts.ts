import { strict } from './strict-args.js'
import { readXlsxCharts as read } from './xlsx/charts/read.js'
import { resolveChartData as resolve } from './xlsx/charts/values.js'

export type { ChartAxis, ChartGroup, ChartPoint, ChartSeries, ChartSource, ChartStyle, ChartValueReader, XlsxChart, XlsxCharts } from './xlsx/charts/types.js'
export type { ChartDrawing, DrawingReadIssue } from './xlsx/drawings/types.js'

export const readXlsxCharts = /* @__PURE__ */ strict(read)
export const resolveChartData = /* @__PURE__ */ strict(resolve)
