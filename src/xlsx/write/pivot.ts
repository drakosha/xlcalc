/**
 * A PivotTable's part (`pivotTableDefinition`) as Excel 365 writes it: its place, its fields with their items and
 * settings, the fields on each axis, the saved layout (`rowItems`, `colItems`), the filters' and the value fields, the
 * formats of its areas (differential formats of the workbook's styles), the areas of its conditional formats, its style
 * and its filters. Its cells are the sheet's.
 */
import type { ConditionalStyle, NumberFormat, PivotArea, PivotCache, PivotDataField, PivotField, PivotFilter, PivotLayout, PivotLine, PivotTable } from '../../contract.js'
import { valueFieldName } from '../../pivot/names.js'
import { dateToSerial } from '../../date-serial.js'
import { escapeAttr } from '../xml/text.js'
import { HEAD, MAIN } from './xml.js'
import { X14 } from './pivot-cache.js'
import { fieldIndex, pivotAreaXml, SUBTOTAL_ATTR } from './pivot-area.js'

/** What the part takes from the workbook's writing: a number format's id, a style's differential format, the rules of the PivotTable's conditional formats by priority. */
export interface PivotWriting {
  numFmt(f: NumberFormat | undefined): number
  dxf(s: ConditionalStyle): number
  date1904: boolean
  /** The PivotTable's conditional formats: each rule's priority on its sheet and its areas. */
  rules: { priority: number; scope?: 'selection' | 'data' | 'field'; type?: 'none' | 'all' | 'row' | 'column'; areas: PivotArea[] }[]
}

const AXIS = { rows: 'axisRow', columns: 'axisCol', pages: 'axisPage' } as const

/** The value field Show Values As of Excel 2010 (`x14:dataField`), and the others' names in the file. */
const SHOW_X14 = new Set(['percentOfParentRow', 'percentOfParentColumn', 'percentOfParent', 'percentOfRunningTotal', 'rankAscending', 'rankDescending'])
const SHOW_NAME: Record<string, string> = { percentOfColumn: 'percentOfCol', percentOfParentColumn: 'percentOfParentCol' }

/** The base item's previous and next (`xlPrevious`, `xlNext`). */
const BASE_ITEM = { previous: 1048828, next: 1048829 } as const

/** A line's type as `rowItems` writes it. */
const lineType = (t: PivotLine['type']) => (t === undefined ? '' : t === 'subtotal' ? ' t="default"' : t === 'grand' || t === 'blank' ? ` t="${t}"` : ` t="${SUBTOTAL_ATTR[t]}"`)

function linesXml(tag: 'rowItems' | 'colItems', lines: PivotLine[] | undefined): string {
  if (!lines) return ''
  const items = lines.map((l) => {
    const attrs = `${lineType(l.type)}${l.repeat ? ` r="${l.repeat}"` : ''}${l.value ? ` i="${l.value}"` : ''}`
    const xs = l.items.map((v) => (v ? `<x v="${v}"/>` : '<x/>')).join('')
    return xs ? `<i${attrs}>${xs}</i>` : `<i${attrs}/>`
  })
  return `<${tag} count="${items.length}">${items.join('')}</${tag}>`
}

/** Where a field lies: its axis, and whether it is summarized. */
function placeOf(p: PivotTable) {
  const axis = new Map<number, keyof typeof AXIS>()
  for (const f of p.rows ?? []) if (f !== 'values') axis.set(f, 'rows')
  for (const f of p.columns ?? []) if (f !== 'values') axis.set(f, 'columns')
  for (const f of p.pages ?? []) axis.set(f.field, 'pages')
  return { axis, data: new Set((p.values ?? []).map((v) => v.field)) }
}

/** A field's form as attributes: compact is the default, outline is not compact, tabular is neither. */
const formAttrs = (l: PivotLayout) => (l === 'compact' ? '' : l === 'outline' ? ' compact="0"' : ' compact="0" outline="0"')

function fieldXml(f: PivotField, i: number, p: PivotTable, place: ReturnType<typeof placeOf>, w: PivotWriting): string {
  const axis = place.axis.get(i)
  const subs = f.subtotals
  const own = subs !== undefined && subs.length > 0
  const sort = f.sort
  let a = ''
  if (f.name !== undefined) a += ` name="${escapeAttr(f.name)}"`
  if (axis) a += ` axis="${AXIS[axis]}"`
  if (place.data.has(i)) a += ' dataField="1"'
  if (f.numberFormat !== undefined) a += ` numFmtId="${w.numFmt(f.numberFormat)}"`
  a += formAttrs(f.layout ?? p.layout ?? 'compact')
  if (f.subtotalTop === false) a += ' subtotalTop="0"'
  if (f.showEmpty !== true) a += ' showAll="0"'
  if (f.subtotalCaption !== undefined) a += ` subtotalCaption="${escapeAttr(f.subtotalCaption)}"`
  if (f.blankLine) a += ' insertBlankRow="1"'
  if (f.pageBreak) a += ' insertPageBreak="1"'
  if (f.multipleItems) a += ' multipleItemSelectionAllowed="1"'
  if (f.includeNewItems) a += ' includeNewItemsInFilter="1"'
  if (sort) a += ` sortType="${sort.order}"`
  if (subs !== undefined && !own) a += ' defaultSubtotal="0"'
  if (own) a += ' defaultSubtotal="0"' + subs.map((s) => ` ${SUBTOTAL_ATTR[s]}Subtotal="1"`).join('')
  const items = f.items?.map((t) => {
    const x = (t.item === undefined ? '' : ` x="${t.item}"`) + (t.caption === undefined ? '' : ` n="${escapeAttr(t.caption)}"`) + (t.hidden ? ' h="1"' : '') + (t.collapsed ? ' sd="0"' : '') + (t.calculated ? ' f="1"' : '') + (t.missing ? ' m="1"' : '')
    return `<item${x}/>`
  })
  // the subtotals' items close the list: automatic, or one per function
  if (items && subs === undefined) items.push('<item t="default"/>')
  if (items && own) for (const s of subs) items.push(`<item t="${SUBTOTAL_ATTR[s]}"/>`)
  const list = items ? `<items count="${items.length}">${items.join('')}</items>` : ''
  const scope = sort?.by ? `<autoSortScope>${pivotAreaXml({ dataOnly: false, outline: false, fieldPosition: 0, references: [{ field: 'values', items: [sort.by.value], selected: false }, ...(sort.by.at ?? []).map((r) => ({ field: r.field, items: [r.item], selected: false }))] })}</autoSortScope>` : ''
  // Repeat Item Labels of a field unlike the PivotTable's, an Excel 2010 setting
  const fill = f.repeatLabels !== undefined && f.repeatLabels !== (p.repeatLabels ?? false) ? `<extLst><ext uri="{2946ED86-A175-432a-8AC1-64E0C546D7DE}" xmlns:x14="${X14}"><x14:pivotField fillDownLabels="${f.repeatLabels ? 1 : 0}"/></ext></extLst>` : ''
  return list || scope || fill ? `<pivotField${a}>${list}${scope}${fill}</pivotField>` : `<pivotField${a}/>`
}

function dataFieldXml(d: PivotDataField, c: PivotCache, w: PivotWriting): string {
  const show = d.showAs ?? 'normal'
  const x14 = SHOW_X14.has(show)
  const base = d.baseItem === undefined ? 0 : typeof d.baseItem === 'number' ? d.baseItem : BASE_ITEM[d.baseItem]
  const name = d.name ?? valueFieldName(d.function ?? 'sum', c.fields[d.field]?.name ?? '')
  const a =
    ` name="${escapeAttr(name)}" fld="${d.field}"` +
    (d.function && d.function !== 'sum' ? ` subtotal="${d.function}"` : '') +
    (show === 'normal' || x14 ? '' : ` showDataAs="${SHOW_NAME[show] ?? show}"`) +
    ` baseField="${d.baseField ?? 0}" baseItem="${base}"` +
    (d.numberFormat === undefined ? '' : ` numFmtId="${w.numFmt(d.numberFormat)}"`)
  return x14 ? `<dataField${a}><extLst><ext uri="{E15A36E0-9728-4e99-A89B-3F7291B0FE68}" xmlns:x14="${X14}"><x14:dataField pivotShowAs="${SHOW_NAME[show] ?? show}"/></ext></extLst></dataField>` : `<dataField${a}/>`
}

/** The operator of a filter's condition by its type's ending, and the text of a label filter's pattern. */
const OPERATOR: [RegExp, string][] = [[/NotEqual$/, 'notEqual'], [/GreaterThanOrEqual$|NewerThanOrEqual$/, 'greaterThanOrEqual'], [/GreaterThan$|NewerThan$/, 'greaterThan'], [/LessThanOrEqual$|OlderThanOrEqual$/, 'lessThanOrEqual'], [/LessThan$|OlderThan$/, 'lessThan']]
const DYNAMIC = /^(tomorrow|today|yesterday|(next|this|last)(Week|Month|Quarter|Year)|yearToDate|Q[1-4]|M\d+)$/

/** A filter's condition as its `autoFilter` holds it: custom conditions, a Top 10, or a dynamic date. */
function conditionXml(f: PivotFilter, w: PivotWriting): string {
  const t = f.type
  if (t === 'count' || t === 'percent' || t === 'sum') {
    const n = f.args?.[0] ?? 10
    return `<top10${f.bottom ? ' top="0"' : ''}${t === 'percent' ? ' percent="1"' : ''} val="${n}" filterVal="${n}"/>`
  }
  if (DYNAMIC.test(t)) return `<dynamicFilter type="${t}"/>`
  const date = t.startsWith('date')
  const arg = (k: number) => {
    const v = f.args?.[k]
    if (date && typeof v === 'string') return String(dateToSerial(v, { dateSystem: w.date1904 ? 1904 : 1900 }))
    return escapeAttr(String(v ?? ''))
  }
  const one = (op: string, v: string) => `<customFilter${op ? ` operator="${op}"` : ''} val="${v}"/>`
  if (/NotBetween$/.test(t)) return `<customFilters>${one('lessThan', arg(0))}${one('greaterThan', arg(1))}</customFilters>`
  if (/Between$/.test(t)) return `<customFilters and="1">${one('greaterThanOrEqual', arg(0))}${one('lessThanOrEqual', arg(1))}</customFilters>`
  const not = /^caption(Not)/.test(t)
  const pattern = /BeginsWith$/.test(t) ? `${arg(0)}*` : /EndsWith$/.test(t) ? `*${arg(0)}` : /Contains$/.test(t) ? `*${arg(0)}*` : arg(0)
  const op = not && pattern !== arg(0) ? 'notEqual' : (OPERATOR.find(([re]) => re.test(t))?.[1] ?? '')
  return `<customFilters>${one(op, pattern)}</customFilters>`
}

function filterXml(f: PivotFilter, id: number, w: PivotWriting): string {
  const label = f.type.startsWith('caption')
  const date = f.type.startsWith('date')
  const s = (k: number) => (f.args?.[k] === undefined || !(label || date) ? '' : ` stringValue${k + 1}="${escapeAttr(String(f.args[k]))}"`)
  const a =
    ` fld="${f.field}"` +
    (f.value === undefined ? '' : ` iMeasureFld="${f.value}"`) +
    ` type="${f.type}" evalOrder="${f.order ?? -1}" id="${id}"` +
    (f.name === undefined ? '' : ` name="${escapeAttr(f.name)}"`) +
    (f.description === undefined ? '' : ` description="${escapeAttr(f.description)}"`) +
    s(0) +
    s(1)
  return `<filter${a}><autoFilter ref="A1"><filterColumn colId="0">${conditionXml(f, w)}</filterColumn></autoFilter></filter>`
}

/** The PivotTable's part, over its cache `c`. */
export function pivotTableXml(p: PivotTable, c: PivotCache, w: PivotWriting): string {
  const layout = p.layout ?? 'compact'
  const place = placeOf(p)
  const cap = p.captions ?? {}
  const off = (name: string, v: boolean | undefined) => (v === false ? ` ${name}="0"` : '')
  const on = (name: string, v: boolean | undefined) => (v ? ` ${name}="1"` : '')
  const a =
    ` name="${escapeAttr(p.name)}" cacheId="${p.cache}"` +
    (p.valuesOnRows ? ' dataOnRows="1"' : '') +
    (p.valuesPosition === undefined ? '' : ` dataPosition="${p.valuesPosition}"`) +
    ' applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1"' +
    ` dataCaption="${escapeAttr(cap.values ?? 'Values')}"` +
    (cap.grandTotal === undefined ? '' : ` grandTotalCaption="${escapeAttr(cap.grandTotal)}"`) +
    (p.errors?.show ? ` showError="1"${p.errors.text === undefined ? '' : ` errorCaption="${escapeAttr(p.errors.text)}"`}` : '') +
    (p.empty?.show === false ? ' showMissing="0"' : p.empty?.text !== undefined ? ` missingCaption="${escapeAttr(p.empty.text)}"` : '') +
    ' updatedVersion="8" minRefreshableVersion="3"' +
    off('showDrill', p.showDrill) +
    off('enableDrill', p.enableDrill) +
    (p.autofitColumns === false ? '' : ' useAutoFormatting="1"') +
    off('preserveFormatting', p.preserveFormatting) +
    ' itemPrintTitles="1" createdVersion="8"' +
    ` indent="${p.indent ?? 1}"` +
    off('showHeaders', p.fieldHeaders) +
    (layout === 'compact' ? '' : ' compact="0" compactData="0"') +
    (p.classicLayout ? ' gridDropZones="1"' : '') +
    (layout === 'tabular' ? '' : ' outline="1" outlineData="1"') +
    on('mergeItem', p.mergeLabels) +
    on('subtotalHiddenItems', p.subtotalHiddenItems) +
    on('asteriskTotals', p.asteriskTotals) +
    off('customListSort', p.customListSort) +
    off('rowGrandTotals', p.grandTotals?.rows) +
    off('colGrandTotals', p.grandTotals?.columns) +
    (p.pagesWrap ? ` pageWrap="${p.pagesWrap}"` : '') +
    on('pageOverThenDown', p.pagesAcrossFirst) +
    (cap.rowHeader === undefined ? '' : ` rowHeaderCaption="${escapeAttr(cap.rowHeader)}"`) +
    (cap.columnHeader === undefined ? '' : ` colHeaderCaption="${escapeAttr(cap.columnHeader)}"`) +
    ` multipleFieldFilters="${p.multipleFieldFilters ? 1 : 0}"`
  const loc = p.location
  const location = `<location ref="${loc.range}" firstHeaderRow="${loc.firstHeaderRow}" firstDataRow="${loc.firstDataRow}" firstDataCol="${loc.firstDataColumn}"${loc.pageRows ? ` rowPageCount="${loc.pageRows}"` : ''}${loc.pageColumns ? ` colPageCount="${loc.pageColumns}"` : ''}/>`
  const fields = `<pivotFields count="${p.fields.length}">${p.fields.map((f, i) => fieldXml(f, i, p, place, w)).join('')}</pivotFields>`
  const axisFields = (tag: string, list: PivotTable['rows']) => (list?.length ? `<${tag} count="${list.length}">${list.map((f) => `<field x="${fieldIndex(f)}"/>`).join('')}</${tag}>` : '')
  const pages = p.pages?.length ? `<pageFields count="${p.pages.length}">${p.pages.map((g) => `<pageField fld="${g.field}"${g.item === undefined ? '' : ` item="${g.item}"`} hier="-1"/>`).join('')}</pageFields>` : ''
  const data = p.values?.length ? `<dataFields count="${p.values.length}">${p.values.map((d) => dataFieldXml(d, c, w)).join('')}</dataFields>` : ''
  const formats = p.formats?.length
    ? `<formats count="${p.formats.length}">${p.formats
        .map((f) => {
          const style: ConditionalStyle = { ...(f.style ?? {}), ...(f.numberFormat === undefined ? {} : { numberFormat: f.numberFormat }) } as ConditionalStyle
          const dxf = f.blank || !Object.keys(style).length ? '' : ` dxfId="${w.dxf(style)}"`
          return `<format${f.blank ? ' action="blank"' : ''}${dxf}>${pivotAreaXml(f.area)}</format>`
        })
        .join('')}</formats>`
    : ''
  const rules = w.rules.length
    ? `<conditionalFormats count="${w.rules.length}">${w.rules
        .map((r) => `<conditionalFormat${r.scope && r.scope !== 'selection' ? ` scope="${r.scope}"` : ''}${r.type && r.type !== 'none' ? ` type="${r.type}"` : ''} priority="${r.priority}"><pivotAreas count="${r.areas.length}">${r.areas.map((x) => pivotAreaXml(x)).join('')}</pivotAreas></conditionalFormat>`)
        .join('')}</conditionalFormats>`
    : ''
  const st = p.style
  const style = st ? `<pivotTableStyleInfo${st.name === undefined ? '' : ` name="${escapeAttr(st.name)}"`} showRowHeaders="${st.rowHeaders ? 1 : 0}" showColHeaders="${st.columnHeaders ? 1 : 0}" showRowStripes="${st.rowStripes ? 1 : 0}" showColStripes="${st.columnStripes ? 1 : 0}" showLastColumn="${st.lastColumn ? 1 : 0}"/>` : ''
  const filters = p.filters?.length ? `<filters count="${p.filters.length}">${p.filters.map((f, i) => filterXml(f, i + 1, w)).join('')}</filters>` : ''
  // Excel 2010's settings: the Values row hidden (Excel shows it without the flag), Repeat All Item Labels, the alt text
  const x14 = (p.valuesRow ? '' : ' hideValuesRow="1"') + (p.repeatLabels ? ' fillDownLabelsDefault="1"' : '') + (p.altText?.title === undefined ? '' : ` altText="${escapeAttr(p.altText.title)}"`) + (p.altText?.description === undefined ? '' : ` altTextSummary="${escapeAttr(p.altText.description)}"`)
  const alt = x14 ? `<ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="${X14}"><x14:pivotTableDefinition${x14}/></ext>` : ''
  const body = `${location}${fields}${axisFields('rowFields', p.rows)}${linesXml('rowItems', p.rowLines)}${axisFields('colFields', p.columns)}${linesXml('colItems', p.columnLines)}${pages}${data}${formats}${rules}${style}${filters}${alt ? `<extLst>${alt}</extLst>` : ''}`
  return `${HEAD}<pivotTableDefinition xmlns="${MAIN}"${a}>${body}</pivotTableDefinition>`
}
