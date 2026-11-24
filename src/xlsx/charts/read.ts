/** ChartML types, axes, styles, and data references with their saved caches. */
import type { CellValue } from '../../contract.js'
import { ERROR_CODE } from '../../syntax/errors.js'
import { xmlNumber } from '../drawings/anchors.js'
import { drawingParts, drawingTheme, openDrawingPackage } from '../drawings/package.js'
import { frames } from '../drawings/read.js'
import type { ChartDrawing, DrawingReadIssue } from '../drawings/types.js'
import { attrsOf, childrenOf, rootOf, type Placed } from '../keep/xml.js'
import { palette } from '../styles/color.js'
import type { ReadXlsxOptions } from '../types.js'
import { attributes, attrsAt, element, elements, flag, has, ownEntry, placed } from '../xml/scan.js'
import { readText } from '../xml/text.js'
import type { ChartAxis, ChartGroup, ChartPoint, ChartSeries, ChartSource, ChartStyle, XlsxChart, XlsxCharts } from './types.js'

const GROUP: Readonly<Record<string, ChartGroup['type']>> = {
  barChart: 'bar', lineChart: 'line', areaChart: 'area', pieChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter',
}
const AXIS: Readonly<Record<string, ChartAxis['kind']>> = {
  catAx: 'category', dateAx: 'date', valAx: 'value', serAx: 'series',
}

const val = (xml: string, name: string) => attrsAt(xml, name).val
const inside = (xml: string, name: string) => element(xml, name)?.body ?? ''
const bodyOf = (xml: string, node?: Placed) =>
  node && node.end !== node.open ? xml.slice(node.open, xml.lastIndexOf('</', node.end - 1)) : ''
const childOf = (xml: string, parent: Placed, name: string) => childrenOf(xml, parent).find((node) => node.local === name)
const tag = (xml: string, name: string) => {
  const node = placed(xml, name, true)[0]
  return node ? xml.slice(node.start, node.end) : ''
}
const text = (xml: string, name: string) => readText(inside(xml, name))
const whole = (value: number | undefined) => value !== undefined && Number.isSafeInteger(value) && value >= 0 ? value : undefined

function points(xml: string, numeric: boolean): ChartPoint[] {
  const out = new Map<number, CellValue>()
  for (const point of elements(xml, 'pt')) {
    const index = whole(xmlNumber(attributes(point.attrs).idx))
    const raw = point.body && element(point.body, 'v')?.body
    if (index === undefined || raw == null) continue
    const value = readText(raw)
    const error = numeric ? ownEntry(ERROR_CODE, value) : undefined
    out.set(index, error ? { error } : numeric ? xmlNumber(value) ?? null : value)
  }
  return [...out].sort(([a], [b]) => a - b).map(([index, value]) => ({ index, value }))
}

function chartSource(xml: string, numeric = false): ChartSource | undefined {
  if (!xml) return
  const formula = text(xml, 'f').trim()
  const reference = formula ? { reference: formula } : {}
  const numberData = element(xml, 'numCache') ?? element(xml, 'numLit')
  const data = numberData ?? element(xml, 'strCache') ?? element(xml, 'strLit') ?? element(xml, 'multiLvlStrCache')
  if (data) {
    const cache = data.body ?? ''
    const levels = elements(cache, 'lvl').map((level) => points(level.body ?? '', false))
    const cached = levels[0] ?? points(cache, numberData !== undefined)
    const greatest = levels.reduce((max, level) => Math.max(max, level.at(-1)?.index ?? -1), cached.at(-1)?.index ?? -1)
    const count = Math.max(whole(xmlNumber(val(cache, 'ptCount'))) ?? 0, greatest + 1)
    const format = text(cache, 'formatCode')
    return { ...reference, count, points: cached, ...(format ? { format } : {}), ...(levels.length ? { levels } : {}) }
  }
  const direct = element(xml, 'v')?.body
  const rich = elements(xml, 't').map((part) => readText(part.body ?? '')).join('')
  if (direct != null || rich) {
    const value = direct == null ? rich : readText(direct)
    return { ...reference, count: 1, points: [{ index: 0, value: numeric ? xmlNumber(value) ?? null : value }] }
  }
  if (formula) return { reference: formula, count: 0, points: [] }
}

const THEME_COLORS = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']

function color(xml: string, theme: string[]): string | undefined {
  if (has(xml, 'noFill')) return 'none'
  const fill = inside(xml, 'solidFill')
  const value = val(fill, 'srgbClr') ?? theme[THEME_COLORS.indexOf(val(fill, 'schemeClr') ?? '')]
  if (value && /^[0-9a-f]{6}$/i.test(value)) return '#' + value.toUpperCase()
}

function style(xml: string, theme: string[]): ChartStyle | undefined {
  const root = rootOf(xml)
  if (!root) return
  const direct = childrenOf(xml, root)
  const proper = (name: string) => direct.find((node) => node.local === name)
  const part = (name: string) => {
    const child = proper(name)
    return child ? xml.slice(child.start, child.end) : ''
  }
  const fill = proper('noFill') ? 'none' : color(part('solidFill'), theme)
  const line = part('ln')
  const stroke = color(line, theme)
  const width = xmlNumber(attrsAt(line, 'ln').w)
  if (fill || stroke || width !== undefined) return {
    ...(fill ? { fill } : {}), ...(stroke ? { stroke } : {}), ...(width === undefined ? {} : { strokeWidth: width / 12700 }),
  }
}

function series(xml: string, theme: string[]): ChartSeries {
  const sources = {
    name: chartSource(inside(xml, 'tx')),
    categories: chartSource(inside(xml, 'cat')),
    values: chartSource(inside(xml, 'val'), true),
    x: chartSource(inside(xml, 'xVal'), true),
    y: chartSource(inside(xml, 'yVal'), true),
  }
  const paint = style(tag(xml, 'spPr'), theme)
  const marker = inside(xml, 'marker')
  return { ...sources, style: paint, ...(marker ? { marker: { symbol: val(marker, 'symbol'), size: xmlNumber(val(marker, 'size')) } } : {}) }
}

function group(xml: string, type: ChartGroup['type'], theme: string[]): ChartGroup {
  const ordered = elements(xml, 'ser').map((node, index) => ({ node, order: xmlNumber(val(node.body ?? '', 'order')) ?? index }))
  ordered.sort((a, b) => a.order - b.order)
  return {
    type, axes: elements(xml, 'axId').flatMap((node) => {
      const id = attributes(node.attrs).val
      return id ? [id] : []
    }),
    series: ordered.map(({ node }) => series(node.body ?? '', theme)),
    direction: val(xml, 'barDir'), grouping: val(xml, 'grouping'), scatterStyle: val(xml, 'scatterStyle'),
    gapWidth: xmlNumber(val(xml, 'gapWidth')), holeSize: xmlNumber(val(xml, 'holeSize')),
  }
}

function axis(xml: string, kind: ChartAxis['kind']): ChartAxis | undefined {
  const id = val(xml, 'axId')
  if (!id) return
  const scale = inside(xml, 'scaling')
  return {
    id, kind, crossesAxis: val(xml, 'crossAx'), position: val(xml, 'axPos'), title: chartSource(inside(inside(xml, 'title'), 'tx')),
    min: xmlNumber(val(scale, 'min')), max: xmlNumber(val(scale, 'max')), logBase: xmlNumber(val(scale, 'logBase')),
    majorUnit: xmlNumber(val(xml, 'majorUnit')), format: attrsAt(xml, 'numFmt').formatCode,
    reversed: val(scale, 'orientation') === 'maxMin', hidden: flag(val(xml, 'delete')), majorGridlines: has(xml, 'majorGridlines'),
  }
}

function chart(xml: string, frame: ChartDrawing, theme: string[], issues: DrawingReadIssue[]): XlsxChart | undefined {
  const report = (reason: DrawingReadIssue['reason']) => issues.push({ sheet: frame.sheet, part: frame.part, reason })
  const root = rootOf(xml)
  const rootChildren = root ? childrenOf(xml, root) : []
  const chartNode = rootChildren.find((node) => node.local === 'chart')
  const plot = chartNode && childOf(xml, chartNode, 'plotArea')
  if (!chartNode || !plot) {
    report('unsupportedChart')
    return
  }
  const groups: ChartGroup[] = []
  const axes: ChartAxis[] = []
  for (const node of childrenOf(xml, plot)) {
    const type = ownEntry(GROUP, node.local)
    const kind = ownEntry(AXIS, node.local)
    const content = bodyOf(xml, node)
    if (type) groups.push(group(content, type, theme))
    else if (kind) {
      const model = axis(content, kind)
      if (model) axes.push(model)
    } else if (node.local.endsWith('Chart')) report('unsupportedChart')
  }
  if (!groups.length) {
    if (!issues.some((issue) => issue.part === frame.part && issue.reason === 'unsupportedChart')) report('unsupportedChart')
    return
  }
  const chartChildren = childrenOf(xml, chartNode)
  const chartChild = (name: string) => chartChildren.find((node) => node.local === name)
  const legend = chartChild('legend')
  const legendBody = bodyOf(xml, legend)
  const title = bodyOf(xml, chartChild('title'))
  const visibility = chartChild('plotVisOnly')
  const blanks = chartChild('dispBlanksAs')
  const none = blanks && attrsOf(xml, blanks).val
  const chartStyle = rootChildren.find((node) => node.local === 'spPr')
  return {
    frame, groups, axes, title: chartSource(inside(title, 'tx')),
    ...(legend ? { legend: { position: val(legendBody, 'legendPos'), overlay: flag(val(legendBody, 'overlay')) } } : {}),
    plotVisibleOnly: !visibility || flag(attrsOf(xml, visibility).val, true),
    ...(none === 'gap' || none === 'span' || none === 'zero' ? { emptyCells: none } : {}),
    style: chartStyle ? style(xml.slice(chartStyle.start, chartStyle.end), theme) : undefined,
  }
}

export async function readXlsxCharts(data: Uint8Array | ArrayBuffer, opts?: ReadXlsxOptions): Promise<XlsxCharts> {
  const pkg = await openDrawingPackage(data, opts)
  const issues: DrawingReadIssue[] = []
  const charts: XlsxChart[] = []
  let theme: string[] | undefined
  for await (const part of drawingParts(pkg, issues, false)) {
    for (const frame of frames(part, issues, true)) {
      if (frame.kind !== 'chart') continue
      const xml = await pkg.archive.text(frame.part)
      if (xml === undefined) {
        issues.push({ sheet: frame.sheet, part: frame.part, reason: 'missingPart' })
        continue
      }
      theme ??= palette(await drawingTheme(pkg), undefined).theme
      const drawn = chart(xml, frame, theme, issues)
      if (drawn) charts.push(drawn)
    }
  }
  return { charts, issues }
}
