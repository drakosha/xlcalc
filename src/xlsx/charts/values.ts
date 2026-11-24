import type { CellValue } from '../../contract.js'
import { InputError } from '../../input-error.js'
import { FormulaError } from '../../syntax/ast.js'
import { parse } from '../../syntax/parse.js'
import type { ChartSeries, ChartSource, ChartValueReader, XlsxChart } from './types.js'

function directValues(reference: string, sheet: string, reader: ChartValueReader): CellValue[] | undefined {
  let ast
  try {
    ast = parse(reference.replace(/^=/, ''))
  } catch (error) {
    if (error instanceof FormulaError) return
    throw error
  }
  if (ast.k !== 'ref' || ast.ref.spill) return
  try {
    const range = ast.ref.text.replace(/\$/g, '').toUpperCase()
    return reader.getValues(ast.sheet ?? sheet, range).flat()
  } catch (error) {
    if (error instanceof InputError && error.path === 'sheet') return
    throw error
  }
}

/** Return fresh chart data using current workbook values wherever a source can be resolved. */
export function resolveChartData(chart: XlsxChart, reader: ChartValueReader): XlsxChart {
  const { sheet } = chart.frame
  const source = (data: ChartSource | undefined): ChartSource | undefined => {
    if (!data?.reference || data.levels) return data
    const values = reader.resolve?.(data.reference, sheet) ?? directValues(data.reference, sheet, reader)
    return values === undefined ? data : { ...data, count: values.length, points: values.map((value, index) => ({ index, value })) }
  }
  const series = (item: ChartSeries): ChartSeries => ({
    ...item, name: source(item.name), categories: source(item.categories), values: source(item.values), x: source(item.x), y: source(item.y),
  })
  return {
    ...chart, title: source(chart.title), groups: chart.groups.map((group) => ({ ...group, series: group.series.map(series) })),
    axes: chart.axes.map((axis) => ({ ...axis, title: source(axis.title) })),
  }
}
