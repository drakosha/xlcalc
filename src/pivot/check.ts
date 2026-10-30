/**
 * The full check of PivotTables, their caches, slicers and timelines as input (`BookInput.pivotCaches`,
 * `SheetInput.pivotTables` and the rest, a conditional format's PivotTable): each object's fields by a table of checks.
 * The core checks what it reads and what refers to what (`input/pivot.ts`); this check is the xlsx subpaths' and
 * `xlcalc/pivot`'s, which registers it for every check of a workbook.
 */
import type { BookInput } from '../contract.js'
import { InputError } from '../input-error.js'
import { checkIsoDate, checkPivotValue } from '../input/pivot.js'
import { array, checkBoolean, checkNumber, checkOneOf, checkString, object, onlyKeys } from '../input/basic.js'
import { checkAddress, checkRange } from '../input/range.js'
import { checkFormat, checkStyle } from '../input/styles.js'

type Check = (v: unknown, p: string) => void

const bool: Check = checkBoolean
const text: Check = (v, p) => void checkString(v, p)
const index: Check = (v, p) => checkNumber(v, p, 0, 2 ** 31, true)
const count: Check = (v, p) => checkNumber(v, p, 0, 2 ** 31, true)
const finite: Check = (v, p) => checkNumber(v, p, -Number.MAX_VALUE, Number.MAX_VALUE)
const oneOf = (values: readonly string[]): Check => (v, p) => checkOneOf(v, p, values)
const list = (c: Check): Check => (v, p) => array(v, p, 'an array').forEach((x, i) => c(x, `${p}[${i}]`))
const range: Check = (v, p) => void checkRange(v, p)

const isoDate: Check = checkIsoDate

/** An object of `fields`, those named in `required` given. */
function shape(fields: Record<string, Check>, required: string[], expected: string): Check {
  const keys = Object.keys(fields)
  return (v, p) => {
    const o = object(v, p, expected)
    onlyKeys(o, p, keys)
    for (const k of required) if (o[k] === undefined) throw new InputError(`${p}.${k}`, `expected ${k}, got undefined`)
    for (const k of keys) if (o[k] !== undefined) fields[k](o[k], `${p}.${k}`)
  }
}

const numberOrDate: Check = (v, p) => (typeof v === 'number' ? finite(v, p) : shape({ date: isoDate }, ['date'], 'a number or { date }')(v, p))
const values = list(checkPivotValue)
const FUNCTIONS = ['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp'] as const
const fn = oneOf(FUNCTIONS)
const axisField: Check = (v, p) => (v === 'values' ? undefined : index(v, p))
const layout = oneOf(['compact', 'outline', 'tabular'])
const shown = shape({ show: bool, text }, ['show'], 'a setting such as { show: true, text: "-" }')
const place = shape({ from: checkAddress, to: checkAddress, fromOffset: list(finite), toOffset: list(finite) }, ['from', 'to'], 'a place such as { from: "E2", to: "G12" }')
const pivotRef = shape({ sheet: text, name: text }, ['sheet', 'name'], 'a PivotTable such as { sheet: "Sheet1", name: "PivotTable1" }')
const LEVELS = ['years', 'quarters', 'months', 'days']
const span = shape({ start: isoDate, end: isoDate }, ['start', 'end'], 'dates such as { start: "2025-01-01T00:00:00", end: "2025-12-31T00:00:00" }')

const SOURCES: Record<string, Check> = {
  range: shape({ type: text, sheet: text, range, book: text }, ['sheet', 'range'], 'a source of cells'),
  table: shape({ type: text, table: text, book: text }, ['table'], 'a table source'),
  name: shape({ type: text, name: text, book: text }, ['name'], 'a named source'),
  consolidation: shape({ type: text, ranges: list(shape({ sheet: text, range, book: text }, ['range'], 'a range such as { sheet: "Data", range: "A1:C9" }')) }, ['ranges'], 'consolidated ranges'),
  external: shape({ type: text, connection: index }, [], 'an external source'),
  dataModel: shape({ type: text, connection: index }, [], 'the data model'),
  scenario: shape({ type: text }, [], 'scenarios'),
}

const source: Check = (v, p) => {
  const o = object(v, p, 'a source such as { type: "range", sheet: "Data", range: "A1:C21" }')
  checkOneOf(o.type, `${p}.type`, Object.keys(SOURCES))
  SOURCES[o.type as string](o, p)
}

const group = shape(
  {
    base: index,
    parent: index,
    range: shape({ by: oneOf(['numbers', 'seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years']), start: numberOrDate, end: numberOrDate, autoStart: bool, autoEnd: bool, interval: finite }, ['by'], 'a grouping such as { by: "months" }'),
    discrete: list(index),
    items: values,
  },
  [],
  'a grouping such as { range: { by: "months" } }',
)

const cacheField = shape(
  {
    name: text,
    numberFormat: checkFormat,
    items: values,
    unused: list(index),
    facts: shape({ types: list(oneOf(['blank', 'number', 'string', 'boolean', 'error', 'date'])), integer: bool, longText: bool, min: numberOrDate, max: numberOrDate }, ['types'], 'facts such as { types: ["number"] }'),
    formula: text,
    group,
  },
  ['name'],
  'a cache field such as { name: "Region" }',
)

/** A cache's records: arrays of values, checked in one loop, as there can be many. */
const records: Check = (v, p) =>
  array(v, p, 'an array of records').forEach((r, i) => {
    const rp = `${p}[${i}]`
    for (const [k, x] of array(r, rp, 'an array of values').entries()) if (typeof x !== 'number' && typeof x !== 'string') checkPivotValue(x, `${rp}[${k}]`)
  })

const cache = shape(
  {
    id: index,
    source,
    fields: list(cacheField),
    records,
    calculatedItems: list(shape({ field: index, item: index, formula: text }, ['field', 'item', 'formula'], 'a calculated item')),
    refreshOnLoad: bool,
    saveData: bool,
    retainItems: count,
    enableRefresh: bool,
    stale: bool,
  },
  ['id', 'source', 'fields'],
  'a pivot cache such as { id: 1, source: { type: "range", sheet: "Data", range: "A1:C21" }, fields: [{ name: "Region" }] }',
)

const reference = shape({ field: axisField, items: list(index), selected: bool, byPosition: bool, relative: bool, subtotals: list(oneOf(['subtotal', ...FUNCTIONS])) }, ['field'], 'a reference such as { field: 0, items: [1] }')

/** A part of a PivotTable that a format or a rule applies to. */
export const checkPivotArea = shape(
  {
    type: oneOf(['normal', 'none', 'data', 'all', 'origin', 'button', 'topEnd', 'topRight']),
    field: axisField,
    axis: oneOf(['rows', 'columns', 'pages', 'values']),
    fieldPosition: index,
    dataOnly: bool,
    labelOnly: bool,
    grandRow: bool,
    grandColumn: bool,
    outline: bool,
    offset: range,
    cacheIndex: bool,
    collapsedLevelsAreSubtotals: bool,
    references: list(reference),
  },
  [],
  'an area such as { type: "data" }',
)

/** A conditional format's PivotTable (`ConditionalBase.pivot`). */
export const checkPivotRuleScope = shape(
  { pivot: text, scope: oneOf(['selection', 'data', 'field']), type: oneOf(['none', 'all', 'row', 'column']), areas: list(checkPivotArea) },
  ['pivot', 'areas'],
  'a PivotTable\'s scope such as { pivot: "PivotTable1", areas: [{ type: "data" }] }',
)

const field = shape(
  {
    name: text,
    items: list(shape({ item: index, caption: text, hidden: bool, collapsed: bool, calculated: bool, missing: bool }, [], 'an item such as { item: 0 }')),
    subtotals: list(fn),
    subtotalTop: bool,
    subtotalCaption: text,
    layout,
    repeatLabels: bool,
    blankLine: bool,
    pageBreak: bool,
    showEmpty: bool,
    sort: shape({ order: oneOf(['ascending', 'descending']), by: shape({ value: index, at: list(shape({ field: index, item: index }, ['field', 'item'], 'an item of a field')) }, ['value'], 'a value field such as { value: 0 }') }, ['order'], 'a sort such as { order: "ascending" }'),
    autoSort: bool,
    includeNewItems: bool,
    multipleItems: bool,
    numberFormat: checkFormat,
  },
  [],
  'a PivotTable field such as {}',
)

const FILTERS = [
  ...['Equal', 'NotEqual', 'BeginsWith', 'NotBeginsWith', 'EndsWith', 'NotEndsWith', 'Contains', 'NotContains', 'GreaterThan', 'GreaterThanOrEqual', 'LessThan', 'LessThanOrEqual', 'Between', 'NotBetween'].map((k) => `caption${k}`),
  ...['Equal', 'NotEqual', 'GreaterThan', 'GreaterThanOrEqual', 'LessThan', 'LessThanOrEqual', 'Between', 'NotBetween'].map((k) => `value${k}`),
  'count', 'percent', 'sum',
  ...['Equal', 'NotEqual', 'OlderThan', 'OlderThanOrEqual', 'NewerThan', 'NewerThanOrEqual', 'Between', 'NotBetween'].map((k) => `date${k}`),
  'tomorrow', 'today', 'yesterday', 'nextWeek', 'thisWeek', 'lastWeek', 'nextMonth', 'thisMonth', 'lastMonth', 'nextQuarter', 'thisQuarter', 'lastQuarter', 'nextYear', 'thisYear', 'lastYear', 'yearToDate',
  'Q1', 'Q2', 'Q3', 'Q4', ...Array.from({ length: 12 }, (_, i) => `M${i + 1}`),
]

const SHOW_AS = ['normal', 'percentOfTotal', 'percentOfColumn', 'percentOfRow', 'percent', 'percentOfParentRow', 'percentOfParentColumn', 'percentOfParent', 'difference', 'percentDiff', 'runTotal', 'percentOfRunningTotal', 'rankAscending', 'rankDescending', 'index']
const line = shape({ type: oneOf(['subtotal', ...FUNCTIONS, 'grand', 'blank']), repeat: count, items: list(index), value: index }, ['items'], 'a line such as { items: [0] }')

const pivot = shape(
  {
    name: text,
    cache: index,
    location: shape({ range, firstHeaderRow: count, firstDataRow: count, firstDataColumn: count, pageRows: count, pageColumns: count }, ['range', 'firstHeaderRow', 'firstDataRow', 'firstDataColumn'], 'a place such as { range: "A3:B8", firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 }'),
    fields: list(field),
    rows: list(axisField),
    columns: list(axisField),
    pages: list(shape({ field: index, item: index }, ['field'], 'a page field such as { field: 0 }')),
    values: list(shape({ field: index, name: text, function: fn, showAs: oneOf(SHOW_AS), baseField: index, baseItem: (v, p) => (v === 'previous' || v === 'next' ? undefined : index(v, p)), numberFormat: checkFormat }, ['field'], 'a value field such as { field: 1 }')),
    valuesOnRows: bool,
    valuesPosition: index,
    filters: list(shape({ field: index, type: oneOf(FILTERS), args: list((v, p) => (typeof v === 'number' ? finite(v, p) : text(v, p))), value: index, bottom: bool, wholeDays: bool, order: index, name: text, description: text }, ['field', 'type'], 'a filter such as { field: 0, type: "captionEqual", args: ["North"] }')),
    grandTotals: shape({ rows: bool, columns: bool }, [], 'grand totals such as { rows: false }'),
    layout,
    indent: count,
    repeatLabels: bool,
    mergeLabels: bool,
    fieldHeaders: bool,
    captions: shape({ grandTotal: text, rowHeader: text, columnHeader: text, values: text }, [], 'captions such as { grandTotal: "Total" }'),
    errors: shown,
    empty: shown,
    pagesAcrossFirst: bool,
    pagesWrap: count,
    subtotalHiddenItems: bool,
    asteriskTotals: bool,
    customListSort: bool,
    multipleFieldFilters: bool,
    showDrill: bool,
    autofitColumns: bool,
    preserveFormatting: bool,
    classicLayout: bool,
    style: shape({ name: text, rowHeaders: bool, columnHeaders: bool, rowStripes: bool, columnStripes: bool, lastColumn: bool }, [], 'a style such as { name: "PivotStyleLight16" }'),
    enableDrill: bool,
    altText: shape({ title: text, description: text }, [], 'alt text such as { title: "Sales" }'),
    formats: list(shape({ area: checkPivotArea, style: (v, p) => void checkStyle(v, p), numberFormat: checkFormat, blank: bool }, ['area'], 'a format such as { area: { type: "data" }, numberFormat: "#,##0" }')),
    rowLines: list(line),
    columnLines: list(line),
  },
  ['name', 'cache', 'location', 'fields'],
  'a PivotTable such as { name: "PivotTable1", cache: 1, location: { range: "A3:B8", firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 }, fields: [] }',
)

const slicerCache = shape(
  {
    name: text,
    field: text,
    pivots: list(pivotRef),
    cache: index,
    table: shape({ name: text, column: text }, ['name', 'column'], 'a table and its column'),
    items: list(shape({ item: index, selected: bool, noData: bool }, ['item'], 'an item such as { item: 0 }')),
    sort: oneOf(['ascending', 'descending']),
    customListSort: bool,
    noData: shape({ hide: bool, mark: bool, last: bool }, [], 'items with no data such as { hide: true }'),
    showMissing: bool,
  },
  ['name', 'field', 'pivots'],
  'a slicer cache such as { name: "Slicer_Region", field: "Region", pivots: [] }',
)

const slicer = shape(
  { name: text, cache: text, caption: text, showCaption: bool, place, style: text, columns: count, rowHeight: count, startItem: index, locked: bool },
  ['name', 'cache'],
  'a slicer such as { name: "Region", cache: "Slicer_Region" }',
)

const timelineCache = shape({ name: text, field: text, pivots: list(pivotRef), cache: index, selection: span, bounds: span }, ['name', 'field', 'pivots', 'cache'], 'a timeline cache such as { name: "NativeTimeline_Date", field: "Date", pivots: [], cache: 1 }')

const timeline = shape(
  { name: text, cache: text, caption: text, place, level: oneOf(LEVELS), selectionLevel: oneOf(LEVELS), showHeader: bool, showSelectionLabel: bool, showLevel: bool, showScrollbar: bool, scrollPosition: isoDate, style: text },
  ['name', 'cache'],
  'a timeline such as { name: "Date", cache: "NativeTimeline_Date" }',
)

/** Every PivotTable, cache, slicer and timeline of the book, and every conditional format's PivotTable, by its shape. */
export function checkPivotShapes(book: BookInput, path = 'book') {
  const each = (v: unknown, p: string, check: Check) => v !== undefined && list(check)(v, p)
  each(book.pivotCaches, `${path}.pivotCaches`, cache)
  each(book.slicerCaches, `${path}.slicerCaches`, slicerCache)
  each(book.timelineCaches, `${path}.timelineCaches`, timelineCache)
  book.sheets.forEach((s, i) => {
    const p = `${path}.sheets[${i}]`
    each(s.pivotTables, `${p}.pivotTables`, pivot)
    each(s.slicers, `${p}.slicers`, slicer)
    each(s.timelines, `${p}.timelines`, timeline)
    s.conditionalFormats?.forEach((c, k) => c.pivot !== undefined && checkPivotRuleScope(c.pivot, `${p}.conditionalFormats[${k}].pivot`))
  })
}
