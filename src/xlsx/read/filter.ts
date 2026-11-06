/**
 * A filter's criteria and sort as a file holds them (ECMA-376 18.3.2, 18.3.1.91-92), a sheet's AutoFilter's or a table's.
 * Excel 2010's forms in `AlternateContent` (`x14:filter`, `x14:customFilters`, `x14:iconFilter`, `x14:sortCondition`, [MS-XLSX]
 * 2.2.6) read as their kinds, the Fallback left. Dates also lie in an extension Excel 365 writes where its object model chose
 * them (`filter-criteria-made`). What the model has no place for, or a file does not spell as the schema does, is left out
 * with its part named: the rich values' extensions and an entry out of its lists.
 */
import type { Color, ColumnFilter, ConditionalStyle, DateGroup, DynamicFilter, DynamicFilterKind, FilterColumn, FilterCriteria, FilterOperator, FilterSort, IconFilter, SortCondition } from '../../contract.js'
import { DATE_LEVELS, DYNAMIC_KINDS, FILTER_OPERATORS, iconSet, SORT_BY, SORT_METHODS } from '../../filter/check.js'
import { parseRange, rectText } from '../../structure/area.js'
import type { Dxf } from '../styles/dxf.js'
import { attributes, element, elements, flag, placed, type Element } from '../xml/scan.js'
import { readText } from '../xml/text.js'

/** Excel 365's extension of a column that holds its dates there (no [MS-XLSX] entry; the set `filter-criteria-made`). */
const EXT_DATES = '{1AD28BCE-077C-4C59-8B6E-1921CE8616D4}'

export interface FilterRead {
  dxfs: Dxf[]
  prefixes?: ReadonlySet<string>
  /** A part of the filter left out, by its name in the file. */
  lost(part: string): void
}

const number = (v: string | undefined) => (v === undefined || v === '' || !Number.isFinite(Number(v)) ? undefined : Number(v))

/** The text with `parts` each replaced by what `to` gives. */
function replaced(xml: string, parts: { start: number; end: number }[], to: (i: number) => string): string {
  let out = ''
  let at = 0
  parts.forEach((p, i) => {
    out += xml.slice(at, p.start) + to(i)
    at = p.end
  })
  return out + xml.slice(at)
}

/** Each `AlternateContent` as what a reader of Excel 2010 takes of it: its Choice of `x14`, else its Fallback. */
/** Prefixes bound to the supported extension namespace, which may be declared on the part or the Choice. */
export const filterPrefixes = (xml: string): Set<string> => new Set([...xml.matchAll(/xmlns:([\w.-]+)=["']http:\/\/schemas\.microsoft\.com\/office\/spreadsheetml\/2009\/9\/main["']/g)].map(m=>m[1]))

function chosen(xml: string,ctx: FilterRead): string {
  if (!xml.includes('AlternateContent>')) return xml
  const supported=new Set([...ctx.prefixes??[],...filterPrefixes(xml)])
  const blocks=placed(xml,'AlternateContent')
  return replaced(xml,blocks,i=>{
    const body=blocks[i].body??''
    const choice=elements(body,'Choice').find(c=>(attributes(c.attrs).Requires??'').split(/\s+/).filter(Boolean).every(p=>supported.has(p)))
    return chosen((choice??element(body,'Fallback'))?.body??'',{...ctx,prefixes:supported})
  })
}

/** The color a filter's differential format names: a cell pattern's foreground as Excel writes it there, else its background or font as a rule's. */
function colorOfDxf(d: Dxf): Color | undefined {
  return d.pattern ?? d.style.fill?.color ?? d.style.font?.color
}

/** The style holding a filter's or sort's color by its differential format `id`, in `font` for a font's; undefined where the file has none. */
function styleOf(id: string | undefined, font: boolean, ctx: FilterRead): ConditionalStyle | undefined {
  if (id === undefined) return {}
  const d = ctx.dxfs[Number(id)]
  if (!d) return undefined
  const color = colorOfDxf(d)
  return !color ? {} : font ? { font: { color } } : { fill: { color } }
}

/** A date or time chosen; undefined without its level or year. */
function dateOf(e: Element): DateGroup | undefined {
  const a = attributes(e.attrs)
  const level = DATE_LEVELS.find((x) => x === a.dateTimeGrouping)
  const year = number(a.year)
  if (!level || year === undefined) return undefined
  const out: DateGroup = { level, year }
  for (const k of ['month', 'day', 'hour', 'minute', 'second'] as const) {
    const v = number(a[k])
    if (v !== undefined) out[k] = v
  }
  return out
}

function valuesOf(e: Element, ctx: FilterRead): ColumnFilter {
  const a = attributes(e.attrs)
  const body = e.body ?? ''
  const values = elements(body, 'filter').map((x) => readText(attributes(x.attrs).val ?? ''))
  const dates = elements(body, 'dateGroupItem').flatMap((x) => {
    const d = dateOf(x)
    if (!d) ctx.lost('dateGroupItem')
    return d ? [d] : []
  })
  return {
    type: 'values',
    ...(values.length ? { values } : {}),
    ...(flag(a.blank) ? { blank: true } : {}),
    ...(dates.length ? { dates } : {}),
    ...(a.calendarType === undefined ? {} : { calendar: readText(a.calendarType) }),
  }
}

function customOf(e: Element, ctx: FilterRead): ColumnFilter | undefined {
  const conditions = elements(e.body ?? '', 'customFilter').map((x) => {
    const a = attributes(x.attrs)
    return { operator: a.operator as FilterOperator | undefined, value: readText(a.val ?? '') }
  })
  if (conditions.length < 1 || conditions.length > 2 || conditions.some((c) => c.operator !== undefined && !FILTER_OPERATORS.includes(c.operator))) return void ctx.lost('customFilters')
  return {
    type: 'custom',
    ...(flag(attributes(e.attrs).and) ? { and: true } : {}),
    conditions: conditions.map(({ operator, value }) => (operator === undefined ? { value } : { operator, value })),
  }
}

function topOf(e: Element, ctx: FilterRead): ColumnFilter | undefined {
  const a = attributes(e.attrs)
  const value = number(a.val)
  if (value === undefined) return void ctx.lost('top10')
  const filterValue = number(a.filterVal)
  return {
    type: 'top',
    ...(flag(a.top, true) ? {} : { bottom: true }),
    ...(flag(a.percent) ? { percent: true } : {}),
    value,
    ...(filterValue === undefined ? {} : { filterValue }),
  }
}

function dynamicOf(e: Element, ctx: FilterRead): ColumnFilter | undefined {
  const a = attributes(e.attrs)
  if (!DYNAMIC_KINDS.includes(a.type as DynamicFilterKind)) return void ctx.lost('dynamicFilter')
  const out: DynamicFilter = { type: 'dynamic', kind: a.type as DynamicFilterKind }
  const [value, maxValue] = [number(a.val), number(a.maxVal)]
  if (value !== undefined) out.value = value
  if (maxValue !== undefined) out.maxValue = maxValue
  if (a.valIso !== undefined) out.valueIso = readText(a.valIso)
  if (a.maxValIso !== undefined) out.maxValueIso = readText(a.maxValIso)
  return out
}

function colorOf(e: Element, ctx: FilterRead): ColumnFilter | undefined {
  const a = attributes(e.attrs)
  const font = !flag(a.cellColor, true)
  const style = styleOf(a.dxfId, font, ctx)
  if (!style) return void ctx.lost('colorFilter')
  return { type: 'color', style, ...(font ? { font: true } : {}) }
}

function iconOf(e: Element, ctx: FilterRead): ColumnFilter | undefined {
  const a = attributes(e.attrs)
  if (a.iconSet === undefined || !iconSet(a.iconSet)) return void ctx.lost('iconFilter')
  const out: IconFilter = { type: 'icon', set: a.iconSet }
  const icon = number(a.iconId)
  if (icon !== undefined) out.icon = icon
  return out
}

const KINDS: [string, (e: Element, ctx: FilterRead) => ColumnFilter | undefined][] = [
  ['filters', valuesOf],
  ['customFilters', customOf],
  ['top10', topOf],
  ['dynamicFilter', dynamicOf],
  ['colorFilter', colorOf],
  ['iconFilter', iconOf],
]

/** What a column of a filter keeps, from its own elements or its dates' extension. */
function filterOf(body: string, exts: string[], ctx: FilterRead): ColumnFilter | undefined {
  for (const [local, read] of KINDS) {
    const e = element(body, local)
    if (e) return read(e, ctx)
  }
  let out: ColumnFilter | undefined
  const mark = element(body, 'extLst')
  for (const ext of mark ? elements(exts[Number(attributes(mark.attrs).n)], 'ext') : []) {
    const uri = (attributes(ext.attrs).uri ?? '').toUpperCase()
    const held = uri === EXT_DATES ? element(ext.body ?? '', 'filters') : undefined
    if (held) out = valuesOf(held, ctx)
    else ctx.lost(`filterColumn extension ${uri}`)
  }
  return out
}

/** The columns of a filter `width` columns wide: each once, inside it, else left out. */
function columnsOf(body: string, width: number, ctx: FilterRead): FilterColumn[] {
  const seen = new Set<number>()
  // an extension holds a `filterColumn` of its own: each list is set aside first, a mark for it in its place
  const lists = placed(body, 'extLst')
  const exts = lists.map((e) => e.body ?? '')
  return elements(chosen(replaced(body, lists, (i) => `<extLst n="${i}"/>`),ctx), 'filterColumn').flatMap((e) => {
    const a = attributes(e.attrs)
    const column = Number(a.colId)
    if (!Number.isInteger(column) || column < 0 || column >= width || seen.has(column)) {
      ctx.lost('filterColumn')
      return []
    }
    seen.add(column)
    const filter = filterOf(e.body ?? '', exts, ctx)
    return [{ column, ...(flag(a.hiddenButton) ? { hiddenButton: true } : {}), ...(flag(a.showButton, true) ? {} : { showButton: false }), ...(filter ? { filter } : {}) }]
  })
}

/** A cell or range as the model holds it; undefined for anything else, a list of areas among them. */
function rangeOf(ref: string | undefined): string | undefined {
  const x = ref === undefined ? null : parseRange(ref)
  return x ? rectText(x) : undefined
}

function conditionOf(e: Element, ctx: FilterRead): SortCondition | undefined {
  const a = attributes(e.attrs)
  const range = rangeOf(a.ref)
  const by = SORT_BY.find((x) => x === a.sortBy)
  if (!range || (a.sortBy !== undefined && !by)) return void ctx.lost('sortCondition')
  const style = by === 'cellColor' || by === 'fontColor' ? styleOf(a.dxfId, by === 'fontColor', ctx) : undefined
  const icon = number(a.iconId)
  return {
    range,
    ...(flag(a.descending) ? { descending: true } : {}),
    ...(by && by !== 'value' ? { by } : {}),
    ...(a.customList === undefined ? {} : { customList: readText(a.customList) }),
    ...(style ? { style } : {}),
    ...(a.iconSet !== undefined && iconSet(a.iconSet) ? { iconSet: a.iconSet } : {}),
    ...(icon === undefined ? {} : { icon }),
  }
}

/** A `sortState` as the model holds it; undefined without a range. */
export function readSort(e: Element, ctx: FilterRead): FilterSort | undefined {
  const a = attributes(e.attrs)
  const range = rangeOf(a.ref)
  if (!range) return void ctx.lost('sortState')
  const body = e.body ?? ''
  for (const ext of elements(element(body, 'extLst')?.body ?? '', 'ext')) ctx.lost(`sortState extension ${attributes(ext.attrs).uri ?? ''}`)
  const conditions = elements(chosen(body,ctx), 'sortCondition').flatMap((c) => conditionOf(c, ctx) ?? [])
  const method = SORT_METHODS.find((x) => x === a.sortMethod)
  return {
    range,
    ...(flag(a.caseSensitive) ? { caseSensitive: true } : {}),
    ...(flag(a.columnSort) ? { byColumns: true } : {}),
    ...(method ? { method } : {}),
    ...(conditions.length ? { conditions } : {}),
  }
}

/**
 * The criteria of an `autoFilter` over `width` columns, with its sort, or the one beside it (`sort`, as a table's part holds
 * its own); undefined without either.
 */
export function readCriteria(filter: Element, width: number, sort: Element | undefined, ctx: FilterRead): FilterCriteria | undefined {
  const body = filter.body ?? ''
  // the sort's conditions hold no filterColumn: the columns are read without it
  const own = element(body, 'sortState')
  const held = own ?? sort
  const read = held && readSort(held, ctx)
  const columns = columnsOf(body, width, ctx)
  const out: FilterCriteria = { ...(columns.length ? { columns } : {}), ...(read ? { sort: read } : {}) }
  return Object.keys(out).length ? out : undefined
}
