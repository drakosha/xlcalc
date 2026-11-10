/** Criteria and sorting shared by worksheet and table writers; stored hidden rows are left as supplied. */
import type { ColumnFilter, ConditionalStyle, FilterCriteria, FilterSort, SortCondition } from '../../contract.js'
import { escapeAttr } from '../xml/text.js'
import type { CellFormats } from './styles.js'

type Styles = Pick<CellFormats, 'dxf'>

/** The icon sets of Excel 2010, which an Excel 2010 element holds. */
const X14_ICONS = ['3Stars', '3Triangles', '5Boxes']

const attrs = (a: Record<string, string | number | boolean | undefined>) =>
  Object.entries(a)
    .filter(([, v]) => v !== undefined)
    .map(([k, v]) => ` ${k}="${escapeAttr(typeof v === 'boolean' ? (v ? '1' : '0') : String(v))}"`)
    .join('')
const tag = (name: string, a: Record<string, string | number | boolean | undefined>, body?: string) => `<${name}${attrs(a)}${body ? `>${body}</${name}>` : '/>'}`
const x14 = (body: string, fallback = '') =>
  `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><mc:Choice Requires="x14">${body}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`

/**
 * The differential format of a color a filter or a sort goes by: a cell pattern's foreground, as Excel writes it for a filter by either
 * color and a sort by a cell's; a font's color for a sort by the font's (the sets `filter-criteria-made`, sheet-font and sheet-sort-font).
 */
function color(style: ConditionalStyle, font: boolean, styles: Styles, sort = false): number {
  const c = font ? (style.font?.color ?? style.fill?.color) : (style.fill?.color ?? style.font?.color)
  if (font && sort && c) return styles.dxf({ font: { color: c } })
  return styles.dxf(c ? { fill: { color: c } } : {}, true)
}

function criterion(f: ColumnFilter, styles: Styles): string {
  switch (f.type) {
    case 'values': {
      const values = f.values ?? []
      const plain = (list: string[]) => list.map((val) => tag('filter', { val })).join('')
      // a value over 255 characters goes in Excel 2010's element, the others in the fallback too
      const list = values.some((v) => v.length > 255) ? x14(values.map((val) => tag('x14:filter', { val })).join(''), plain(values.filter((v) => v.length <= 255))) : plain(values)
      const dates = (f.dates ?? [])
        .map((d) => tag('dateGroupItem', { year: d.year, month: d.month, day: d.day, hour: d.hour, minute: d.minute, second: d.second, dateTimeGrouping: d.level }))
        .join('')
      return tag('filters', { blank: f.blank, calendarType: f.calendar }, list + dates)
    }
    case 'custom': {
      const prefix = f.conditions.some((c) => c.value.length > 255) ? 'x14:' : ''
      const text = tag(`${prefix}customFilters`, { and: f.and }, f.conditions.map((c) => tag(`${prefix}customFilter`, { operator: c.operator, val: c.value })).join(''))
      return prefix ? x14(text) : text
    }
    case 'top':
      return tag('top10', { top: f.bottom === undefined ? undefined : !f.bottom, percent: f.percent, val: f.value, filterVal: f.filterValue })
    case 'dynamic':
      return tag('dynamicFilter', { type: f.kind, val: f.value, maxVal: f.maxValue, valIso: f.valueIso, maxValIso: f.maxValueIso })
    case 'color':
      return tag('colorFilter', { dxfId: color(f.style, !!f.font, styles), cellColor: f.font === undefined ? undefined : !f.font })
    case 'icon': {
      const prefix = X14_ICONS.includes(f.set) ? 'x14:' : ''
      const text = tag(`${prefix}iconFilter`, { iconSet: f.set, iconId: f.icon })
      return prefix ? x14(text) : text
    }
  }
}

function condition(c: SortCondition, styles: Styles): string {
  const prefix = c.iconSet !== undefined && X14_ICONS.includes(c.iconSet) ? 'x14:' : ''
  const dxfId = c.style === undefined ? undefined : color(c.style, c.by === 'fontColor', styles, true)
  const text = tag(`${prefix}sortCondition`, { ref: c.range, descending: c.descending, sortBy: c.by, customList: c.customList, dxfId, iconSet: c.iconSet, iconId: c.icon })
  return prefix ? x14(text) : text
}

/** The last sort, without changing the order of any cells. */
export const sortXml = (s: FilterSort | undefined, styles: Styles): string =>
  s
    ? tag(
        'sortState',
        { ref: s.range, caseSensitive: s.caseSensitive, columnSort: s.byColumns, sortMethod: s.method },
        (s.conditions ?? []).map((c) => condition(c, styles)).join(''),
      )
    : ''

/** A filter and its criteria, or `fallback` (the filter mode's criteria every row meets) where none were given. */
export function filterXml(range: string, criteria: FilterCriteria | undefined, styles: Styles, fallback = ''): string {
  const columns = (criteria?.columns ?? [])
    .map((c) => tag('filterColumn', { colId: c.column, hiddenButton: c.hiddenButton, showButton: c.showButton }, c.filter ? criterion(c.filter, styles) : ''))
    .join('')
  return tag('autoFilter', { ref: range }, (criteria === undefined ? fallback : columns) + sortXml(criteria?.sort, styles))
}
