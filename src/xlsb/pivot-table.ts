import { SUBTOTAL_ATTR } from '../xlsx/pivot-names.js'
import { escapeAttr } from '../xlsx/xml/text.js'
import { rangeAt } from './fields.js'
import { Filters } from './filter.js'
import { NS, type Part } from './parts.js'
import { records, type Payload } from './records.js'

const AGGREGATES = ['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp'] as const
const LINE_TYPES = ['data', 'default', ...AGGREGATES.map((fn) => SUBTOTAL_ATTR[fn]), 'grand', 'blank']
const SHOW_AS = ['normal', 'difference', 'percent', 'percentDiff', 'runTotal', 'percentOfRow', 'percentOfCol',
  'percentOfTotal', 'index', 'percentOfParentRow', 'percentOfParentCol', 'percentOfParent', 'percentOfRunningTotal',
  'rankAscending', 'rankDescending']
const COMPARISONS = ['Equal', 'NotEqual', 'GreaterThan', 'GreaterThanOrEqual', 'LessThan', 'LessThanOrEqual', 'Between', 'NotBetween']
const FILTER_TYPES = ['', 'count', 'percent', 'sum', 'captionEqual', 'captionNotEqual', 'captionBeginsWith',
  'captionNotBeginsWith', 'captionEndsWith', 'captionNotEndsWith', 'captionContains', 'captionNotContains',
  ...COMPARISONS.slice(2).map((kind) => `caption${kind}`), ...COMPARISONS.map((kind) => `value${kind}`),
  'dateEqual', 'dateOlderThan', 'dateNewerThan', 'dateBetween', 'tomorrow', 'today', 'yesterday', 'nextWeek', 'thisWeek',
  'lastWeek', 'nextMonth', 'thisMonth', 'lastMonth', 'nextQuarter', 'thisQuarter', 'lastQuarter', 'nextYear', 'thisYear',
  'lastYear', 'yearToDate', 'Q1', 'Q2', 'Q3', 'Q4', ...Array.from({ length: 12 }, (_, month) => `M${month + 1}`),
  'dateNotEqual', 'dateOlderThanOrEqual', 'dateNewerThanOrEqual', 'dateNotBetween']

const attr = (key: string, value: string | number | undefined): string => value === undefined ? '' : ` ${key}="${escapeAttr(String(value))}"`
const bit = (value: number, index: number): number => (value >>> index) & 1
const optional = (payload: Payload, flag: number): string | undefined => flag ? payload.wide() : undefined

export function pivotTable({ bytes, path, xml, pending }: Part): void {
  const filters = new Filters(xml)
  let lineItems = 0
  for (const { type, payload } of records(bytes, path)) {
    switch (type) {
      case 280: xml.add(view(payload)); break // BrtBeginSXView
      case 315: xml.add('</pivotTableDefinition>'); break
      case 314: xml.add(location(payload)); break
      case 287: xml.add(`<pivotFields count="${payload.u32()}">`); break
      case 288: xml.add('</pivotFields>'); break
      case 285: xml.add(field(payload)); break
      case 286: xml.add('</pivotField>'); break
      case 283: xml.add(`<items count="${payload.u32()}">`); break
      case 284: xml.add('</items>'); break
      case 282: xml.add(item(payload)); break
      case 309: case 311: { // BrtBeginISXVDRws / BrtBeginISXVDCols
        const axis = type === 309 ? 'rowFields' : 'colFields'
        const count = payload.u32()
        if (count * 4 !== payload.remaining) return payload.fail(`invalid pivot axis field count ${count}`)
        xml.add(`<${axis} count="${count}">`)
        for (let i = 0; i < count; i++) xml.add(`<field x="${payload.i32()}"/>`)
        xml.add(`</${axis}>`)
        break
      }
      case 299: xml.add(`<rowItems count="${payload.u32()}">`); break
      case 300: xml.add('</rowItems>'); break
      case 301: xml.add(`<colItems count="${payload.u32()}">`); break
      case 302: xml.add('</colItems>'); break
      case 297: { // BrtBeginSXLI
        const repeat = payload.u16()
        const kind = payload.u8()
        payload.skip(1)
        lineItems = payload.u32()
        const value = payload.i32()
        xml.add(`<i r="${repeat}" i="${value}"${attr('t', LINE_TYPES[kind])}>`)
        break
      }
      case 298: xml.add('</i>'); break
      case 388: // BrtBeginISXVIs
        if (lineItems * 4 !== payload.remaining) return payload.fail(`invalid pivot line item count ${lineItems}`)
        for (let i = 0; i < lineItems; i++) xml.add(`<x v="${payload.u32()}"/>`)
        break
      case 291: xml.add(`<pageFields count="${payload.u32()}">`); break
      case 292: xml.add('</pageFields>'); break
      case 289: {
        const fieldIndex = payload.i32()
        const itemIndex = payload.u32()
        payload.skip(4)
        const flags = payload.u8()
        if (flags & 1) payload.wide()
        if (flags & 2) payload.wide()
        xml.add(`<pageField fld="${fieldIndex}"${itemIndex === 0x1000fe ? '' : attr('item', itemIndex)}/>`)
        break
      }
      case 295: xml.add(`<dataFields count="${payload.u32()}">`); break
      case 296: xml.add('</dataFields>'); break
      case 293: xml.add(dataField(payload)); break
      case 294: xml.add('</dataField>'); break
      case 1044: {
        payload.skip(4)
        const show = SHOW_AS[payload.u32()]
        payload.skip(8)
        if (show) xml.add(`<x14:dataField pivotShowAs="${show}"/>`)
        break
      }
      case 513: xml.add(style(payload)); break
      case 1061: {
        payload.skip(4)
        xml.add(`<x14:pivotField fillDownLabels="${payload.u32()}"/>`)
        payload.skip(4)
        break
      }
      case 1062: xml.add(view14(payload)); break
      case 599: xml.add(`<filters count="${payload.u32()}">`); break
      case 600: xml.add('</filters>'); break
      case 601: xml.add(pivotFilter(payload)); break
      case 602: xml.add('</filter>'); break
      case 303: case 459: case 558: case 1147:
        pending()
        continue
      default:
        filters.take(type, payload)
        continue
    }
    payload.done()
  }
}

function pivotFilter(payload: Payload): string {
  const field = payload.i32()
  payload.skip(4)
  const kind = FILTER_TYPES[payload.u32()]
  if (!kind) return payload.fail('unknown PivotTable filter type')
  payload.skip(8)
  const measure = payload.i32()
  payload.skip(4)
  const flags = payload.u16()
  const name = optional(payload, bit(flags, 0))
  const description = optional(payload, bit(flags, 1))
  const first = optional(payload, bit(flags, 2))
  const last = optional(payload, bit(flags, 3))
  return `<filter fld="${field}" type="${kind}"` + (measure < 0 ? '' : attr('iMeasureFld', measure))
    + attr('name', name) + attr('description', description) + attr('stringValue1', first) + attr('stringValue2', last) + '>'
}

function view(payload: Payload): string {
  payload.skip(2)
  const display = payload.u8()
  const indentAndHeaders = payload.u8()
  const behavior = payload.u8()
  const totals = payload.u8()
  const captions = payload.u8()
  const displayTag = payload.u8()
  const layout = payload.u8()
  const filters = payload.u8()
  payload.skip(2)
  const axis = payload.u8()
  const pageWrap = payload.u8()
  payload.skip(2)
  const dataPosition = payload.i32()
  payload.skip(8)
  const cache = payload.u32()
  const name = payload.wide()
  const dataCaption = optional(payload, bit(captions, 3))
  const grand = optional(payload, bit(captions, 4))
  const error = optional(payload, 1 - bit(layout, 6))
  const missing = optional(payload, 1 - bit(layout, 7))
  for (const flag of [bit(captions, 5), bit(captions, 6), bit(captions, 7), bit(displayTag, 6)]) {
    if (flag) payload.wide()
  }
  const columnHeader = optional(payload, bit(filters, 3))
  const rowHeader = optional(payload, bit(filters, 2))
  return `<pivotTableDefinition ${NS} name="${escapeAttr(name)}" cacheId="${cache}"`
    + ` indent="${indentAndHeaders & 0x7f}" compact="${bit(layout, 0)}" outline="${bit(layout, 1)}"`
    + ` rowGrandTotals="${bit(totals, 5)}" colGrandTotals="${bit(totals, 6)}" mergeItem="${bit(captions, 2)}"`
    + ` useAutoFormatting="${bit(totals, 0)}" showError="${bit(totals, 1)}" showMissing="${bit(totals, 2)}"`
    + ` pageOverThenDown="${bit(totals, 3)}" subtotalHiddenItems="${bit(totals, 4)}" pageWrap="${pageWrap}"`
    + ` showHeaders="${1 - bit(indentAndHeaders, 7)}" showDrill="${1 - bit(display, 4)}" enableDrill="${bit(behavior, 5)}"`
    + ` preserveFormatting="${bit(behavior, 7)}" gridDropZones="${1 - bit(layout, 4)}" dataOnRows="${bit(axis, 0)}"`
    + ` multipleFieldFilters="${1 - bit(filters, 1)}" customListSort="${1 - bit(filters, 6)}"`
    + (dataPosition < 0 ? '' : attr('dataPosition', dataPosition))
    + attr('dataCaption', dataCaption) + attr('grandTotalCaption', grand) + attr('errorCaption', error)
    + attr('missingCaption', missing) + attr('colHeaderCaption', columnHeader) + attr('rowHeaderCaption', rowHeader) + '>'
}

function location(payload: Payload): string {
  const firstRow = payload.u32()
  const lastRow = payload.u32()
  const firstCol = payload.u32()
  const lastCol = payload.u32()
  // The binary header/data coordinates are absolute. SpreadsheetML stores offsets from the range's first cell.
  const ref = rangeAt(payload, firstRow, lastRow, firstCol, lastCol)
  const header = payload.u32() - firstRow
  const data = payload.u32() - firstRow
  const column = payload.u32() - firstCol
  return `<location ref="${ref}" firstHeaderRow="${header}" firstDataRow="${data}" firstDataCol="${column}"`
    + ` rowPageCount="${payload.u32()}" colPageCount="${payload.u32()}"/>`
}

function field(payload: Payload): string {
  payload.skip(1)
  const subtotals = payload.u16()
  const flags = payload.u8()
  const format = payload.u32()
  const display = payload.u32()
  payload.skip(8)
  const name = optional(payload, bit(flags, 5))
  const subtotal = optional(payload, bit(flags, 6))
  if (bit(flags, 3)) payload.wide()
  const custom = AGGREGATES.map((fn, i) => bit(subtotals, i + 1) ? attr(`${SUBTOTAL_ATTR[fn]}Subtotal`, 1) : '').join('')
  const sort = bit(display, 12) ? attr('sortType', bit(display, 13) ? 'ascending' : 'descending') : ''
  return `<pivotField defaultSubtotal="${bit(subtotals, 0)}" compact="${bit(flags, 4)}" outline="${bit(display, 6)}"`
    + ` showAll="${bit(display, 5)}" insertBlankRow="${bit(display, 7)}" subtotalTop="${bit(display, 8)}"`
    + ` insertPageBreak="${bit(display, 11)}" includeNewItemsInFilter="${1 - bit(display, 18)}" multipleItemSelectionAllowed="${bit(display, 19)}"`
    + attr('name', name) + attr('subtotalCaption', subtotal) + (format === 0xffffffff ? '' : attr('numFmtId', format))
    + sort + custom + '>'
}

function item(payload: Payload): string {
  const kind = payload.u8()
  const flags = payload.u16()
  const index = payload.i32()
  const label = optional(payload, bit(flags, 4))
  return `<item h="${bit(flags, 0)}" sd="${1 - bit(flags, 1)}" f="${bit(flags, 2)}" m="${bit(flags, 3)}"`
    + attr('n', label) + attr('t', LINE_TYPES[kind]) + (index < 0 ? '' : attr('x', index)) + '/>'
}

function dataField(payload: Payload): string {
  const field = payload.i32()
  const aggregate = payload.u32()
  const show = payload.u32()
  const baseField = payload.i32()
  const baseItem = payload.u32()
  const format = payload.u32()
  return `<dataField fld="${field}" baseField="${baseField}" baseItem="${baseItem}"`
    + attr('subtotal', AGGREGATES[aggregate]) + attr('showDataAs', SHOW_AS[show])
    + attr('numFmtId', format) + attr('name', optional(payload, payload.u8())) + '>'
}

function style(payload: Payload): string {
  const flags = payload.u16()
  const name = payload.nullableWide()
  return `<pivotTableStyleInfo showLastColumn="${bit(flags, 1)}" showRowStripes="${bit(flags, 2)}" showColStripes="${bit(flags, 3)}"`
    + ` showRowHeaders="${bit(flags, 4)}" showColHeaders="${bit(flags, 5)}"${attr('name', name)}/>`
}

function view14(payload: Payload): string {
  payload.skip(4)
  const flags = payload.u8()
  payload.skip(4)
  const title = payload.nullableWide()
  const description = payload.nullableWide()
  payload.nullableWide()
  return `<extLst><ext><x14:pivotTableDefinition fillDownLabelsDefault="${bit(flags, 0)}" hideValuesRow="${1 - bit(flags, 4)}"`
    + attr('altText', title) + attr('altTextSummary', description) + '/></ext></extLst>'
}
