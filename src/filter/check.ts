/**
 * A filter's criteria and sort checked as input (`AutoFilter.criteria`, `TableInput.filter`, `SheetInput.sort`): each
 * column inside the filter's range and listed once, the operators and kinds of their lists. The workbook holds them and
 * does not apply them, so a value's own bounds are left to the file.
 */
import type { ColumnFilter, DynamicFilterKind, FilterOperator, IconSetName, SortCondition } from '../contract.js'
import { ICON_SETS, checkRuleStyle } from '../cond/check.js'
import { InputError } from '../input-error.js'
import { array, checkBoolean, checkNumber, checkOneOf, checkString, given, object, onlyKeys } from '../input/basic.js'
import { checkRange } from '../input/range.js'

export const FILTER_OPERATORS: readonly FilterOperator[] = ['equal', 'notEqual', 'lessThan', 'lessThanOrEqual', 'greaterThan', 'greaterThanOrEqual']
export const DATE_LEVELS = ['year', 'month', 'day', 'hour', 'minute', 'second'] as const
export const DYNAMIC_KINDS: readonly DynamicFilterKind[] = [
  'null', 'aboveAverage', 'belowAverage', 'tomorrow', 'today', 'yesterday', 'nextWeek', 'thisWeek', 'lastWeek', 'nextMonth', 'thisMonth', 'lastMonth',
  'nextQuarter', 'thisQuarter', 'lastQuarter', 'nextYear', 'thisYear', 'lastYear', 'yearToDate', 'Q1', 'Q2', 'Q3', 'Q4',
  'M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10', 'M11', 'M12',
]
export const SORT_BY: readonly NonNullable<SortCondition['by']>[] = ['value', 'cellColor', 'fontColor', 'icon']
export const SORT_METHODS = ['stroke', 'pinYin'] as const
export const iconSet = (s: string): s is IconSetName => Object.hasOwn(ICON_SETS, s)

const anyNumber = (v: unknown, p: string) => checkNumber(v, p, -Number.MAX_VALUE, Number.MAX_VALUE)
const place = (v: unknown, p: string) => checkNumber(v, p, 0, Number.MAX_SAFE_INTEGER, true)
const icons = (v: unknown, p: string) => checkOneOf(v, p, Object.keys(ICON_SETS))

function checkDate(v: unknown, p: string) {
  const o = object(v, p, 'a date such as { level: "month", year: 2024, month: 3 }')
  onlyKeys(o, p, ['level', ...DATE_LEVELS])
  checkOneOf(o.level, `${p}.level`, DATE_LEVELS)
  anyNumber(o.year, `${p}.year`)
  given(o, p, { month: anyNumber, day: anyNumber, hour: anyNumber, minute: anyNumber, second: anyNumber })
}

function checkCondition(v: unknown, p: string) {
  const o = object(v, p, 'a condition such as { operator: "greaterThan", value: "10" }')
  onlyKeys(o, p, ['operator', 'value'])
  if (o.operator !== undefined) checkOneOf(o.operator, `${p}.operator`, FILTER_OPERATORS)
  checkString(o.value, `${p}.value`)
}

/** Each kind of filter's fields and their checks, besides `type`. */
const KINDS: Record<ColumnFilter['type'], (o: Record<string, unknown>, p: string) => void> = {
  values(o, p) {
    onlyKeys(o, p, ['type', 'values', 'blank', 'dates', 'calendar'])
    if (o.values !== undefined) array(o.values, `${p}.values`, 'an array of texts').forEach((x, i) => checkString(x, `${p}.values[${i}]`))
    if (o.dates !== undefined) array(o.dates, `${p}.dates`, 'an array of dates').forEach((x, i) => checkDate(x, `${p}.dates[${i}]`))
    given(o, p, { blank: checkBoolean, calendar: checkString })
  },
  custom(o, p) {
    onlyKeys(o, p, ['type', 'and', 'conditions'])
    given(o, p, { and: checkBoolean })
    const list = array(o.conditions, `${p}.conditions`, 'an array of one or two conditions')
    if (list.length < 1 || list.length > 2) throw new InputError(`${p}.conditions`, `expected one or two conditions, got ${list.length}`)
    list.forEach((x, i) => checkCondition(x, `${p}.conditions[${i}]`))
  },
  top(o, p) {
    onlyKeys(o, p, ['type', 'bottom', 'percent', 'value', 'filterValue'])
    anyNumber(o.value, `${p}.value`)
    given(o, p, { bottom: checkBoolean, percent: checkBoolean, filterValue: anyNumber })
  },
  dynamic(o, p) {
    onlyKeys(o, p, ['type', 'kind', 'value', 'maxValue', 'valueIso', 'maxValueIso'])
    checkOneOf(o.kind, `${p}.kind`, DYNAMIC_KINDS)
    given(o, p, { value: anyNumber, maxValue: anyNumber, valueIso: checkString, maxValueIso: checkString })
  },
  color(o, p) {
    onlyKeys(o, p, ['type', 'style', 'font'])
    if (o.style === undefined) throw new InputError(`${p}.style`, 'expected the color as a style such as { fill: { color: "#FF0000" } }')
    checkRuleStyle(o.style, `${p}.style`)
    given(o, p, { font: checkBoolean })
  },
  icon(o, p) {
    onlyKeys(o, p, ['type', 'set', 'icon'])
    icons(o.set, `${p}.set`)
    given(o, p, { icon: place })
  },
}

function checkFilter(v: unknown, p: string) {
  const o = object(v, p, 'a filter such as { type: "values", values: ["Apple"] }')
  checkOneOf(o.type, `${p}.type`, Object.keys(KINDS))
  KINDS[o.type as ColumnFilter['type']](o, p)
}

function checkSortCondition(v: unknown, p: string) {
  const o = object(v, p, 'a sort condition such as { range: "C2:C21", descending: true }')
  onlyKeys(o, p, ['range', 'descending', 'by', 'customList', 'style', 'iconSet', 'icon'])
  checkRange(o.range, `${p}.range`, false)
  if (o.by !== undefined) checkOneOf(o.by, `${p}.by`, SORT_BY)
  if (o.style !== undefined) checkRuleStyle(o.style, `${p}.style`)
  given(o, p, { descending: checkBoolean, customList: checkString, iconSet: icons, icon: place })
}

/** A sort as a file keeps it. */
export function checkFilterSort(v: unknown, p: string) {
  const o = object(v, p, 'a sort such as { range: "A2:C21", conditions: [{ range: "C2:C21" }] }')
  onlyKeys(o, p, ['range', 'caseSensitive', 'byColumns', 'method', 'conditions'])
  checkRange(o.range, `${p}.range`, false)
  given(o, p, { caseSensitive: checkBoolean, byColumns: checkBoolean })
  if (o.method !== undefined) checkOneOf(o.method, `${p}.method`, SORT_METHODS)
  if (o.conditions !== undefined) array(o.conditions, `${p}.conditions`, 'an array of sort conditions').forEach((x, i) => checkSortCondition(x, `${p}.conditions[${i}]`))
}

/** The criteria of a filter whose range is `width` columns wide. */
export function checkFilterCriteria(v: unknown, p: string, width: number) {
  const o = object(v, p, 'criteria such as { columns: [{ column: 0, filter: { type: "values", values: ["Apple"] } }] }')
  onlyKeys(o, p, ['columns', 'sort'])
  if (o.sort !== undefined) checkFilterSort(o.sort, `${p}.sort`)
  if (o.columns === undefined) return
  const seen = new Set<number>()
  array(o.columns, `${p}.columns`, 'an array of columns').forEach((c, i) => {
    const q = `${p}.columns[${i}]`
    const f = object(c, q, 'a column such as { column: 0, filter: { type: "values", values: ["Apple"] } }')
    onlyKeys(f, q, ['column', 'hiddenButton', 'showButton', 'filter'])
    checkNumber(f.column, `${q}.column`, 0, width - 1, true)
    if (seen.has(f.column as number)) throw new InputError(`${q}.column`, `column ${f.column} is listed twice (a filter has one entry for each column)`)
    seen.add(f.column as number)
    given(f, q, { hiddenButton: checkBoolean, showButton: checkBoolean, filter: checkFilter })
  })
}
