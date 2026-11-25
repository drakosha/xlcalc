/** BIFF8 PivotTable views over a workbook cache. */
import { parseRange } from '../structure/area.js'
import { escapeAttr } from '../xlsx/xml/text.js'
import type { Payload } from '../xlsb/records.js'
import type { Xml } from '../xlsb/xml.js'
import { PivotFilters } from './pivot-filters.js'
import { PivotConditionalFormats, PivotFormats } from './pivot-formats.js'
import { reference, unicode, type RecordInfo } from './records.js'
import type { Styles } from './styles.js'

const SUBTOTAL = ['default', 'sum', 'countA', 'avg', 'max', 'min', 'product', 'count', 'stdDev', 'stdDevP', 'var', 'varP']
const LINES = ['', 'default', 'sum', 'countA', 'count', 'avg', 'max', 'min', 'product', 'stdDev', 'stdDevP', 'var', 'varP', 'grand', 'blank']
const FUNCTION = ['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp']
const SHOW = ['normal', 'difference', 'percent', 'percentDiff', 'runTotal', 'percentOfRow', 'percentOfCol', 'percentOfTotal', 'index']
const AXIS = ['', 'axisRow', 'axisCol', '', 'axisPage']

interface Field {
  attrs: string
  items: string[]
  compact?: boolean
  outline?: boolean
  includeNew?: boolean
  sortMeasure?: number
}

export interface ViewCache {
  id: number
  fields: ReadonlyMap<string, number>
  legacyViewEnablesRefresh?: true
}

interface Pivot {
  range: string
  firstHeaderRow: number
  firstDataRow: number
  firstDataCol: number
  cache?: ViewCache
  expectFields: number
  rowFields: number
  colFields: number
  name: string
  caption: string
  attrs: string
  fields: Field[]
  compact?: boolean
  outline?: boolean
  indent?: number
  fieldProperties?: Field
  data: string[]
  pages: string[]
  axes: string[]
  lines: string[]
  filters: PivotFilters
  formats: PivotFormats
  conditional: PivotConditionalFormats
  pageRows?: number
  pageColumns?: number
  style: string
  leftOut: boolean
}

export function pivotHeader(payload: Payload) {
  const range = reference(payload)
  const area = parseRange(range)!
  const firstHeaderRow = Math.max(0, payload.u16() - area.r1)
  const firstDataRow = Math.max(0, payload.u16() - area.r1)
  const firstDataCol = Math.max(0, payload.u16() - area.c1)
  const cacheIndex = payload.u16()
  payload.skip(2)
  const valuesAxis = payload.u16()
  const valuesPosition = payload.i16()
  const expectFields = payload.u16()
  const rowFields = payload.u16()
  const colFields = payload.u16()
  payload.skip(8)
  const flags = payload.u16()
  payload.skip(2)
  const size = payload.u16()
  const captionSize = payload.u16()
  const name = unicode(payload, size)
  const caption = unicode(payload, captionSize)
  const attrs = ` useAutoFormatting="${(flags >>> 9) & 1}"`
    + (valuesAxis & 1 ? ` dataOnRows="1"${valuesPosition >= 0 ? ` dataPosition="${valuesPosition}"` : ''}` : '')
    + ` rowGrandTotals="${flags & 1}" colGrandTotals="${(flags >>> 1) & 1}"`
  return { range, firstHeaderRow, firstDataRow, firstDataCol, cacheIndex, expectFields, rowFields, colFields, name, caption, attrs }
}

function view(payload: Payload, caches: readonly (ViewCache | undefined)[]): Pivot {
  const { cacheIndex, ...header } = pivotHeader(payload)
  return { ...header, cache: caches[cacheIndex],
    fields: [], data: [], pages: [], axes: [], lines: [],
    filters: new PivotFilters(), formats: new PivotFormats(), conditional: new PivotConditionalFormats(), style: '', leftOut: false }
}

/** QsiSXTag binds a view or a query table to the extension records that follow it. */
export function queryTag(payload: Payload) {
  payload.skip(4)
  const pivot = payload.u16() === 1
  const flags = payload.u16()
  payload.skip(8)
  return { pivot, flags, name: unicode(payload) }
}

/** Old BIFF files can mark the cache as OLAP only in the corresponding worksheet view. */
export function olapCacheIndexes(records: Iterable<RecordInfo>): number[] {
  const views = new Map<string, number>()
  const indexes: number[] = []
  for (const { type, payload } of records) {
    if (type === 0x00b0) {
      const { name, cacheIndex } = pivotHeader(payload)
      views.set(name, cacheIndex)
    } else if (type === 0x0802 && views.size) {
      try {
        const { pivot, flags, name } = queryTag(payload)
        const index = views.get(name)
        if (pivot && flags & 4 && index !== undefined) indexes.push(index)
      } catch {
        // An unusable tag cannot identify the cache. The sheet still has its saved cells.
      }
    }
  }
  return indexes
}

function field(payload: Payload): Field {
  const axis = payload.u16()
  payload.skip(2)
  const subtotals = payload.u16()
  payload.skip(2)
  const size = payload.u16()
  const name = size === 0xffff ? undefined : unicode(payload, size)
  let attrs = ' showAll="0"'
  const placed = AXIS[axis & 7]
  if (placed) attrs += ` axis="${placed}"`
  if (axis & 8) attrs += ' dataField="1"'
  if (name) attrs += ` name="${escapeAttr(name)}"`
  if (!(subtotals & 1)) attrs += ' defaultSubtotal="0"'
  for (let i = 1; i < SUBTOTAL.length; i++) if (subtotals & (1 << i)) attrs += ` ${SUBTOTAL[i]}Subtotal="1"`
  return { attrs, items: [] }
}

function item(payload: Payload): string {
  const kind = payload.u16()
  const flags = payload.u16()
  const index = payload.i16()
  const length = payload.u16()
  const name = length === 0xffff ? undefined : unicode(payload, length)
  if (kind) return ''
  return `<item${index >= 0 ? ` x="${index}"` : ''}${name ? ` n="${escapeAttr(name)}"` : ''}`
    + (flags & 1 ? ' h="1"' : '') + (flags & 2 ? ' sd="0"' : '') + (flags & 8 ? ' f="1"' : '') + (flags & 16 ? ' m="1"' : '') + '/>'
}

function dataField(payload: Payload): string {
  const index = payload.u16()
  const fn = FUNCTION[payload.u16()] ?? 'sum'
  const show = SHOW[payload.u16()] ?? 'normal'
  const base = payload.u16()
  const item = payload.u16()
  const format = payload.u16()
  const size = payload.u16()
  const name = size === 0xffff ? undefined : unicode(payload, size)
  return `<dataField fld="${index}" subtotal="${fn}" showDataAs="${show}" numFmtId="${format}"`
    + (name ? ` name="${escapeAttr(name)}"` : '')
    + (show === 'normal' ? '' : ` baseField="${base}" baseItem="${item === 0x7ffb ? 1048828 : item === 0x7ffc ? 1048829 : item}"`) + '/>'
}

function lineItems(payload: Payload, width: number): string {
  let xml = ''
  while (payload.remaining) {
    const repeat = payload.u16()
    const kind = LINES[payload.u16() & 0x7fff] ?? ''
    const used = payload.u16()
    const value = (payload.u16() >>> 1) & 255
    const entries: number[] = []
    for (let i = 0; i < width; i++) entries.push(payload.u16())
    const values = kind === 'grand' || kind === 'blank' ? [0] : entries.slice(repeat, used).map((item) => item < 0x7fff ? item : 0)
    xml += `<i${kind ? ` t="${kind}"` : ''}${repeat ? ` r="${repeat}"` : ''}${value ? ` i="${value}"` : ''}>`
      + values.map((item) => `<x v="${item}"/>`).join('') + '</i>'
  }
  return xml
}

function viewOptions(pivot: Pivot, payload: Payload): void {
  payload.skip(2)
  const errorLength = payload.u16()
  const missingLength = payload.u16()
  payload.skip(4)
  pivot.pageRows = payload.u16()
  pivot.pageColumns = payload.u16()
  const flags = payload.u32()
  const wrap = (flags >>> 1) & 255
  payload.skip(6)
  const errors = errorLength === 0xffff ? undefined : unicode(payload, errorLength)
  const missing = missingLength === 0xffff ? undefined : unicode(payload, missingLength)
  if (flags & 1) pivot.attrs += ' pageOverThenDown="1"'
  if (wrap) pivot.attrs += ` pageWrap="${wrap}"`
  if (!(flags & (1 << 17))) pivot.attrs += ' enableDrill="0"'
  if (!(flags & (1 << 19))) pivot.attrs += ' preserveFormatting="0"'
  if (flags & (1 << 20)) pivot.attrs += ' mergeItem="1"'
  if (flags & (1 << 21)) pivot.attrs += ' showError="1"'
  if (!(flags & (1 << 22))) pivot.attrs += ' showMissing="0"'
  if (flags & (1 << 23)) pivot.attrs += ' subtotalHiddenItems="1"'
  if (errors !== undefined) pivot.attrs += ` errorCaption="${escapeAttr(errors)}"`
  if (missing !== undefined) pivot.attrs += ` missingCaption="${escapeAttr(missing)}"`
}

function fieldOptions(field: Field, payload: Payload): void {
  const flags = payload.u32()
  const measure = payload.i16()
  payload.skip(2)
  const format = payload.u16()
  if (flags & 1) field.attrs = field.attrs.replace(' showAll="0"', ' showAll="1"')
  if (format && format !== 0xffff) field.attrs += ` numFmtId="${format}"`
  if (flags & (1 << 9)) {
    field.attrs += ` sortType="${flags & (1 << 10) ? 'ascending' : 'descending'}"`
    if (measure >= 0) field.sortMeasure = measure
  }
  if (flags & (1 << 14)) field.attrs += ' insertPageBreak="1"'
  field.outline = Boolean(flags & (1 << 21))
  if (flags & (1 << 22)) field.attrs += ' insertBlankRow="1"'
  if (!(flags & (1 << 23))) field.attrs += ' subtotalTop="0"'
}

export function pivotViews(source: Generator<RecordInfo>, caches: readonly (ViewCache | undefined)[], styles: Styles): Pivot[] {
  const views: Pivot[] = []
  let pivot: Pivot | undefined
  for (const { type, payload } of source) {
    if (type === 0x00b0) {
      pivot = view(payload, caches)
      views.push(pivot)
    } else if (pivot?.cache !== undefined) {
      switch (type) {
        case 0x00b1: pivot.fields.push(field(payload)); break
        case 0x00b2: pivot.fields.at(-1)?.items.push(item(payload)); break
        case 0x00b4: {
          let axis = ''
          while (payload.remaining) axis += `<field x="${payload.i16()}"/>`
          pivot.axes.push(axis)
          break
        }
        case 0x00b5: {
          const width = pivot.lines.length ? pivot.colFields : pivot.rowFields
          if (payload.remaining % (8 + width * 2) === 0) pivot.lines.push(lineItems(payload, width))
          else {
            pivot.lines.push('')
            pivot.leftOut = true
          }
          break
        }
        case 0x00b6:
          while (payload.remaining) {
            const field = payload.u16()
            const item = payload.u16()
            payload.skip(2)
            pivot.pages.push(`<pageField fld="${field}"${item === 0x7ffd ? '' : ` item="${item}"`}/>`)
          }
          break
        case 0x00c5: pivot.data.push(dataField(payload)); break
        case 0x00f1: viewOptions(pivot, payload); break
        case 0x0802: {
          payload.skip(4)
          if (payload.u16() !== 1) break
          const flags = payload.u16()
          if (flags & 1) pivot.cache.legacyViewEnablesRefresh = true
          // Excel's SpreadsheetML attribute, like the BIFF option, suppresses the asterisk.
          if (payload.u32() & 2) pivot.attrs += ' asteriskTotals="1"'
          break
        }
        case 0x0810: {
          payload.skip(14)
          const grand = unicode(payload)
          if (grand) pivot.attrs += ` grandTotalCaption="${escapeAttr(grand)}"`
          break
        }
        case 0x0100: {
          const field = pivot.fields.at(-1)
          if (field) fieldOptions(field, payload)
          break
        }
        case 0x0864: {
          payload.skip(4)
          const owner = payload.u8()
          const property = payload.u8()
          if (owner === 0x1d) pivot.filters.take(property, payload, pivot.expectFields)
          else if (owner === 0x17) {
            if (property === 0) {
              payload.skip(6)
              const name = unicode(payload)
              const index = pivot.cache.fields.get(name)
              pivot.fieldProperties = index === undefined ? undefined : pivot.fields[index]
            } else if (property === 0x19 && pivot.fieldProperties) {
              const flags = payload.u16()
              pivot.fieldProperties.compact = Boolean(flags & 8)
              pivot.fieldProperties.includeNew = !(flags & 0x20)
            } else if (property === 0xff) pivot.fieldProperties = undefined
          } else if (owner === 0 && property === 0x19) {
            const flags = payload.u16()
            const display = payload.u16()
            pivot.compact = Boolean(flags & 1)
            pivot.outline = Boolean(flags & 2)
            pivot.indent = (display >>> 8) & 127
            if (flags & 128) pivot.attrs += ' multipleFieldFilters="0"'
            if (!(flags & 0x10)) pivot.attrs += ' gridDropZones="1"'
            if (flags & 0x400) pivot.attrs += ' customListSort="0"'
            if (display & 0x10) pivot.attrs += ' showDrill="0"'
            if (display & 0x8000) pivot.attrs += ' showHeaders="0"'
          } else if (owner === 0 && property === 0x1e) {
            payload.skip(6)
            const flags = payload.u16()
            const style = payload.wide(payload.u16())
            const name = style ? ` name="${escapeAttr(style)}"` : ''
            pivot.style = `<pivotTableStyleInfo${name} showRowHeaders="${(flags >>> 4) & 1}" showColHeaders="${(flags >>> 5) & 1}"`
              + ` showRowStripes="${(flags >>> 2) & 1}" showColStripes="${(flags >>> 3) & 1}" showLastColumn="${(flags >>> 1) & 1}"/>`
          } else if (owner === 0 && (property === 0x21 || property === 0x22)) {
            const length = payload.u32()
            payload.skip(2)
            const caption = unicode(payload)
            // Compact captions fit in one SXAddl segment, which holds up to 255 characters.
            if (caption.length !== length || length > 255) pivot.leftOut = true
            else {
              const attr = property === 0x21 ? 'rowHeaderCaption' : 'colHeaderCaption'
              pivot.attrs += ` ${attr}="${escapeAttr(caption)}"`
            }
          } else pivot.conditional.take(owner, property, payload)
          break
        }
        default: pivot.formats.take(type, payload, styles)
      }
    }
  }
  return views
}

export function viewPart(pivot: Pivot, xml: Xml): boolean {
  if (pivot.cache === undefined || pivot.fields.length !== pivot.expectFields || !pivot.name) return false
  const slots = pivot.rowFields ? ['rowFields', 'colFields'] : ['colFields']
  const lines = ['rowItems', 'colItems']
  // Before the layout extension, Excel uses the classic grid and a tabular default. Fields can still use outline.
  const legacy = pivot.compact === undefined
  const compact = pivot.compact ?? false
  const outline = pivot.outline ?? false
  const layout = ` compact="${+compact}" outline="${+outline}" indent="${pivot.indent ?? 0}"`
  xml.add(`<pivotTableDefinition name="${escapeAttr(pivot.name)}" cacheId="${pivot.cache.id}" dataCaption="${escapeAttr(pivot.caption)}"`
    + layout + (legacy ? ' gridDropZones="1"' : '') + pivot.attrs + '>')
  const pageRows = pivot.pageRows ? ` rowPageCount="${pivot.pageRows}"` : ''
  const pageColumns = pivot.pageColumns ? ` colPageCount="${pivot.pageColumns}"` : ''
  xml.add(`<location ref="${pivot.range}" firstHeaderRow="${pivot.firstHeaderRow}" firstDataRow="${pivot.firstDataRow}" firstDataCol="${pivot.firstDataCol}"${pageRows}${pageColumns}/><pivotFields>`)
  for (const field of pivot.fields) {
    const { attrs, items, sortMeasure } = field
    const form = ` compact="${+(field.compact ?? compact)}" outline="${+(field.outline ?? outline)}"`
    const include = field.includeNew ?? legacy
    xml.add(`<pivotField${attrs}${form}${include ? ' includeNewItemsInFilter="1"' : ''}>`)
    if (items.length) {
      xml.add('<items>')
      for (const item of items) xml.add(item)
      xml.add('</items>')
    }
    if (sortMeasure !== undefined) {
      xml.add(`<autoSortScope><pivotArea><references><reference field="-2"><x v="${sortMeasure}"/></reference></references></pivotArea></autoSortScope>`)
    }
    xml.add('</pivotField>')
  }
  xml.add('</pivotFields>')
  for (const [i, axis] of pivot.axes.entries()) if (slots[i]) xml.add(`<${slots[i]}>${axis}</${slots[i]}>`)
  for (const [i, body] of pivot.lines.entries()) if (lines[i] && body) xml.add(`<${lines[i]}>${body}</${lines[i]}>`)
  if (pivot.pages.length) xml.add('<pageFields>' + pivot.pages.join('') + '</pageFields>')
  if (pivot.data.length) xml.add('<dataFields>' + pivot.data.join('') + '</dataFields>')
  const formatLoss = pivot.formats.write(xml)
  const conditionalLoss = pivot.conditional.write(xml)
  const filters = pivot.filters.entries
  if (filters.length) xml.add('<filters>' + filters.join('') + '</filters>')
  pivot.leftOut ||= pivot.filters.incomplete || formatLoss || conditionalLoss
  const style = pivot.style || (legacy ? '<pivotTableStyleInfo showRowHeaders="1" showColHeaders="1" showLastColumn="1"/>' : '')
  xml.add(style + '</pivotTableDefinition>')
  return true
}
