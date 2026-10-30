/**
 * A PivotTable part (`pivotTableDefinition`) as a sheet's `PivotTable`: its place, fields on their axes with their items,
 * subtotals, layout and sorting, value fields with Show Values As (Excel 2010's calculations from its extension), filters,
 * options, captions, style, formats of its areas and the saved layout of its lines. What the model has no place for
 * (chart formats) is reported.
 */
import type {
  CellStyle,
  NumberFormat,
  PivotArea,
  PivotAreaReference,
  PivotAxisField,
  PivotDataField,
  PivotField,
  PivotFieldItem,
  PivotFilter,
  PivotFilterType,
  PivotFormat,
  PivotFunction,
  PivotLayout,
  PivotLine,
  PivotRuleScope,
  PivotShowAs,
  PivotTable,
} from '../../contract.js'
import type { DateSystemOptions } from '../../date-serial.js'
import { serialToUtcMs } from '../../model.js'
import type { Dxf } from '../styles/dxf.js'
import { attributes, element, elements, flag, type Element } from '../xml/scan.js'
import { attrText, take, values } from './pivot-cache.js'

/** The functions of subtotals by the names of the file's item types (`countA` counts every value, `count` numbers). */
export const ITEM_FUNCTIONS: Record<string, PivotFunction> = { sum: 'sum', countA: 'count', avg: 'average', max: 'max', min: 'min', product: 'product', count: 'countNums', stdDev: 'stdDev', stdDevP: 'stdDevp', var: 'var', varP: 'varp' }

/** Show Values As by the file's names, Excel 2010's in its extension (`pivotShowAs`) among them. */
const SHOW_AS: Record<string, PivotShowAs> = { percentOfCol: 'percentOfColumn', percentOfParentCol: 'percentOfParentColumn' }
const showAs = (v: string): PivotShowAs => SHOW_AS[v] ?? (v as PivotShowAs)

const AXES: Record<string, PivotArea['axis']> = { axisRow: 'rows', axisCol: 'columns', axisPage: 'pages', axisValues: 'values' }

/** Excel's base items before and after the one shown, `previous` and `next`. */
const PREVIOUS = 1048828

/** A field's index, `values` for the value fields (-2, or 4294967294 where the file writes it unsigned). */
const axisField = (v: string | undefined): PivotAxisField => {
  const n = Number(v ?? 0)
  return n === -2 || n === 4294967294 ? 'values' : n
}

/** A form by the flags of its compact and outline forms: tabular without outline. */
const layoutOf = (compact: boolean, outline: boolean): PivotLayout => (!outline ? 'tabular' : compact ? 'compact' : 'outline')

/** The subtotals a reference or a field picks by its `…Subtotal` flags. */
const subtotals = (a: Record<string, string>): PivotFunction[] => Object.keys(ITEM_FUNCTIONS).flatMap((k) => (flag(a[`${k}Subtotal`]) ? [ITEM_FUNCTIONS[k]] : []))

export function readArea(e: Element): PivotArea {
  const a = attributes(e.attrs)
  const area: PivotArea = {}
  if (a.type !== undefined && a.type !== 'normal') area.type = a.type as PivotArea['type']
  if (a.field !== undefined) area.field = axisField(a.field)
  if (a.axis !== undefined) area.axis = AXES[a.axis]
  take(area, a, 'fieldPosition:n dataOnly:t labelOnly:f grandRow:f grandCol=grandColumn:f outline:t cacheIndex:f collapsedLevelsAreSubtotals:f offset')
  const refs = elements(element(e.body ?? '', 'references')?.body ?? '', 'reference').map((r): PivotAreaReference => {
    const ra = attributes(r.attrs)
    const ref: PivotAreaReference = { field: axisField(ra.field) }
    const items = values(r.body ?? '') as number[]
    if (items.length) ref.items = items
    take(ref, ra, 'selected:t byPosition:f relative:f')
    const sub: PivotAreaReference['subtotals'] = [...(flag(ra.defaultSubtotal) ? ['subtotal' as const] : []), ...subtotals(ra)]
    if (sub.length) ref.subtotals = sub
    return ref
  })
  if (refs.length) area.references = refs
  return area
}

const itemOf = (a: Record<string, string>): PivotFieldItem => take({ ...(flag(a.sd, true) ? {} : { collapsed: true }) }, a, 'x=item:n n=caption h=hidden:f f=calculated:f m=missing:f')

function fieldOf(e: Element, layout: PivotLayout, format: (id: string | undefined) => NumberFormat | undefined): PivotField {
  const a = attributes(e.attrs)
  const body = e.body ?? ''
  const f = take<PivotField>({}, a, 'name subtotalCaption subtotalTop:t insertBlankRow=blankLine:f insertPageBreak=pageBreak:f includeNewItemsInFilter=includeNewItems:f multipleItemSelectionAllowed=multipleItems:f')
  const items = elements(element(body, 'items')?.body ?? '', 'item')
    .map((i) => attributes(i.attrs))
    .filter((i) => i.t === undefined || i.t === 'data')
    .map(itemOf)
  if (items.length) f.items = items
  // custom subtotals by their functions' flags, which Excel writes without taking the automatic one away; none without either
  const custom = subtotals(a)
  if (custom.length || !flag(a.defaultSubtotal, true)) f.subtotals = custom
  const own = layoutOf(flag(a.compact, true), flag(a.outline, true))
  if (own !== layout) f.layout = own
  // Show items with no data: `showAll`, true by default, which Excel writes false
  if (flag(a.showAll, true)) f.showEmpty = true
  if (flag(attributes(element(body, 'pivotField')?.attrs ?? '').fillDownLabels)) f.repeatLabels = true
  if (a.sortType === 'ascending' || a.sortType === 'descending') {
    f.sort = { order: a.sortType }
    const scope = element(body, 'autoSortScope')
    const area = scope && element(scope.body ?? '', 'pivotArea')
    const refs = area ? readArea(area).references ?? [] : []
    const value = refs.find((r) => r.field === 'values')
    if (value?.items?.length) {
      const at = refs.flatMap((r) => (r.field === 'values' || !r.items?.length ? [] : [{ field: r.field, item: r.items[0] }]))
      f.sort.by = { value: value.items[0], ...(at.length ? { at } : {}) }
    }
  }
  const nf = format(a.numFmtId)
  if (nf !== undefined) f.numberFormat = nf
  return f
}

function dataFieldOf(e: Element, format: (id: string | undefined) => NumberFormat | undefined): PivotDataField {
  const a = attributes(e.attrs)
  const d = take<PivotDataField>({ field: Number(a.fld ?? 0) }, a, 'name')
  if (a.subtotal !== undefined && a.subtotal !== 'sum') d.function = a.subtotal as PivotFunction
  const x14 = attributes(element(e.body ?? '', 'dataField')?.attrs ?? '').pivotShowAs
  const show = x14 ?? a.showDataAs
  if (show !== undefined && show !== 'normal') {
    d.showAs = showAs(show)
    take(d, a, 'baseField:n')
    const item = Number(a.baseItem ?? 0)
    if (a.baseItem !== undefined) d.baseItem = item === PREVIOUS ? 'previous' : item === PREVIOUS + 1 ? 'next' : item
  }
  const nf = format(a.numFmtId)
  if (nf !== undefined) d.numberFormat = nf
  return d
}

/** The opening tags of a PivotTable's filters, which name their field (`fld`): the AutoFilter inside one has `<filter>` tags of its own. */
const PIVOT_FILTER = /<(?:[\w.-]+:)?filter\s((?:"[^"]*"|'[^']*'|[^>"'])*?\bfld=(?:"[^"]*"|'[^']*'|[^>"'])*?)\/?>/g

/** A filter from its opening tag's attributes and its text up to the next filter: the dialog's values, a date filter's as ISO dates. */
function filterOf(attrs: string, body: string, dates: DateSystemOptions): PivotFilter {
  const a = attributes(attrs)
  const f: PivotFilter = { field: Number(a.fld ?? 0), type: a.type as PivotFilterType }
  const top = element(body, 'top10')
  const ta = top && attributes(top.attrs)
  const given = (k: string) => a[k] !== undefined
  // a date of the 1904 system is 1,462 days after the 1900 system's of its number
  const day = (n: number) => (f.type.startsWith('date') ? new Date(serialToUtcMs(n + (dates.dateSystem === 1904 ? 1462 : 0))).toISOString().slice(0, 10) : n)
  const args = given('stringValue1') ? ['stringValue1', 'stringValue2'].filter(given).map((k) => attrText(a[k])) : ta ? [Number(ta.val)] : elements(body, 'customFilter').map((c) => day(Number(attributes(c.attrs).val)))
  if (args.length) f.args = args
  if (ta && !flag(ta.top, true)) f.bottom = true
  if (Number(a.evalOrder) >= 0) f.order = Number(a.evalOrder)
  return take(f, a, 'iMeasureFld=value:n name description')
}

/** The PivotTable's filters: each from its tag to the next one's. */
function filtersOf(xml: string, dates: DateSystemOptions): PivotFilter[] {
  const tags = [...xml.matchAll(PIVOT_FILTER)]
  return tags.map((m, k) => filterOf(m[1], xml.slice(m.index! + m[0].length, tags[k + 1]?.index ?? xml.length), dates))
}

const LINE_TYPES: Record<string, PivotLine['type']> = { default: 'subtotal', grand: 'grand', blank: 'blank', ...ITEM_FUNCTIONS }

function linesOf(xml: string, local: string): PivotLine[] | undefined {
  const e = element(xml, local)
  if (!e) return undefined
  return elements(e.body ?? '', 'i').map((i) => {
    const a = attributes(i.attrs)
    const type = a.t === undefined ? undefined : LINE_TYPES[a.t]
    return take<PivotLine>({ items: values(i.body ?? '') as number[], ...(type ? { type } : {}) }, a, 'r=repeat:p i=value:p')
  })
}

const axisOf = (xml: string, local: string): PivotAxisField[] | undefined => {
  const e = element(xml, local)
  return e && elements(e.body ?? '', 'field').map((f) => axisField(attributes(f.attrs).x))
}

/** A format's style: a dxf's conditional style with the font's name and size, the alignment and protection a cell style holds. */
function formatOf(e: Element, dxfs: Dxf[]): PivotFormat {
  const a = attributes(e.attrs)
  const area = element(e.body ?? '', 'pivotArea')
  const out: PivotFormat = { area: area ? readArea(area) : {} }
  if (a.action === 'blank') out.blank = true
  const d = a.dxfId === undefined ? undefined : dxfs[Number(a.dxfId)]
  if (d) {
    const { numberFormat, font, ...rest } = d.style
    const { name, size, ...cell } = d.cell ?? {}
    const style: CellStyle = { ...rest, ...cell }
    const full = { ...font, ...(name === undefined ? {} : { name }), ...(size === undefined ? {} : { size }) }
    if (Object.keys(full).length) style.font = full
    if (Object.keys(style).length) out.style = style
    if (numberFormat !== undefined) out.numberFormat = numberFormat
  }
  return out
}

/** What the reader of a PivotTable needs of the workbook: number formats by id and the differential formats. */
export interface TableContext extends DateSystemOptions {
  format: (id: string | undefined) => NumberFormat | undefined
  dxfs: Dxf[]
}

/** A PivotTable part as the model's PivotTable over the cache of `cache`; with what it has that the model does not (`chartFormats`). */
export function readPivotTable(xml: string, cache: number, ctx: TableContext): { table: PivotTable; charts: boolean } {
  const root = element(xml, 'pivotTableDefinition')
  const a = attributes(root?.attrs ?? '')
  const la = attributes(element(xml, 'location')?.attrs ?? '')
  const layout = layoutOf(flag(a.compact, true), flag(a.outline))
  const t: PivotTable = {
    name: attrText(a.name),
    cache,
    location: { range: la.ref ?? 'A1', firstHeaderRow: Number(la.firstHeaderRow ?? 0), firstDataRow: Number(la.firstDataRow ?? 0), firstDataColumn: Number(la.firstDataCol ?? 0) },
    fields: elements(element(xml, 'pivotFields')?.body ?? '', 'pivotField').map((e) => fieldOf(e, layout, ctx.format)),
  }
  take(t.location, la, 'rowPageCount=pageRows:p colPageCount=pageColumns:p')
  const rows = axisOf(xml, 'rowFields')
  if (rows) t.rows = rows
  const columns = axisOf(xml, 'colFields')
  if (columns) t.columns = columns
  const pages = element(xml, 'pageFields')
  if (pages)
    t.pages = elements(pages.body ?? '', 'pageField').map((e) => take({ field: 0 }, attributes(e.attrs), 'fld=field:n item:n'))
  const data = element(xml, 'dataFields')
  if (data) t.values = elements(data.body ?? '', 'dataField').map((e) => dataFieldOf(e, ctx.format))
  const filters = filtersOf(xml, { dateSystem: ctx.dateSystem })
  if (filters.length) t.filters = filters
  const grand = take({}, a, 'rowGrandTotals=rows:t colGrandTotals=columns:t')
  if (Object.keys(grand).length) t.grandTotals = grand
  if (layout !== 'compact') t.layout = layout
  if (a.indent !== undefined && Number(a.indent) !== 1) t.indent = Number(a.indent)
  const x14 = attributes(element(xml.slice(xml.lastIndexOf('<extLst')), 'pivotTableDefinition')?.attrs ?? '')
  if (flag(x14.fillDownLabelsDefault)) t.repeatLabels = true
  // the Values row, which Excel shows unless Excel 2010's extension hides it
  if (!flag(x14.hideValuesRow)) t.valuesRow = true
  take(t, a, 'dataOnRows=valuesOnRows:f dataPosition=valuesPosition:n mergeItem=mergeLabels:f showHeaders=fieldHeaders:t pageOverThenDown=pagesAcrossFirst:f subtotalHiddenItems:f asteriskTotals:f customListSort:t multipleFieldFilters:t showDrill:t preserveFormatting:t gridDropZones=classicLayout:f enableDrill:t pageWrap=pagesWrap:p')
  // Autofit column widths on update: `useAutoFormatting`, false by default in the file, which Excel writes true
  if (!flag(a.useAutoFormatting)) t.autofitColumns = false
  const captions = take({}, a, `grandTotalCaption=grandTotal rowHeaderCaption=rowHeader colHeaderCaption=columnHeader${a.dataCaption === 'Values' ? '' : ' dataCaption=values'}`)
  if (Object.keys(captions).length) t.captions = captions
  if (flag(a.showError) || a.errorCaption !== undefined) t.errors = { show: flag(a.showError), ...(a.errorCaption === undefined ? {} : { text: attrText(a.errorCaption) }) }
  if (!flag(a.showMissing, true) || a.missingCaption !== undefined) t.empty = { show: flag(a.showMissing, true), ...(a.missingCaption === undefined ? {} : { text: attrText(a.missingCaption) }) }
  const sa = element(xml, 'pivotTableStyleInfo')
  if (sa) {
    t.style = take({}, attributes(sa.attrs), 'name showRowHeaders=rowHeaders:f showColHeaders=columnHeaders:f showRowStripes=rowStripes:f showColStripes=columnStripes:f showLastColumn=lastColumn:f')
  }
  const alt = take({}, x14, 'altText=title altTextSummary=description')
  if (Object.keys(alt).length) t.altText = alt
  const formats = element(xml, 'formats')
  if (formats) t.formats = elements(formats.body ?? '', 'format').map((e) => formatOf(e, ctx.dxfs))
  const rowLines = linesOf(xml, 'rowItems')
  if (rowLines) t.rowLines = rowLines
  const columnLines = linesOf(xml, 'colItems')
  if (columnLines) t.columnLines = columnLines
  return { table: t, charts: element(xml, 'chartFormats') !== undefined }
}

/** The scopes of a PivotTable's conditional formats by their rules' priorities (`conditionalFormats`). */
export function pivotRuleScopes(xml: string, out: Map<number, PivotRuleScope>) {
  const name = attrText(attributes(element(xml, 'pivotTableDefinition')?.attrs ?? '').name)
  for (const e of elements(element(xml, 'conditionalFormats')?.body ?? '', 'conditionalFormat')) {
    const a = attributes(e.attrs)
    const scope: PivotRuleScope = { pivot: name, areas: elements(element(e.body ?? '', 'pivotAreas')?.body ?? '', 'pivotArea').map(readArea) }
    if (a.scope === 'data' || a.scope === 'field') scope.scope = a.scope
    if (a.type === 'all' || a.type === 'row' || a.type === 'column') scope.type = a.type
    out.set(Number(a.priority ?? 0), scope)
  }
}
