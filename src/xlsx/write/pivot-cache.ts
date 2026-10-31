/**
 * A pivot cache's parts as Excel 365 writes them (`pivotCacheDefinition`, `pivotCacheRecords`): its source, its fields
 * with the kinds of value each holds and its items, grouping, calculated fields and items, and the rows when it saves
 * them. A cache whose PivotTables changed without a refresh (`stale`) is refreshed by Excel on opening.
 */
import type { PivotCache, PivotCacheField, PivotFieldGroup, PivotValue } from '../../contract.js'
import { ERROR_TEXT } from '../../syntax/ast.js'
import { escapeAttr } from '../xml/text.js'
import { counted, countedIf, give, HEAD, MAIN, REL, X14 } from './xml.js'
import { pivotAreaXml } from './pivot-area.js'

/** A cache value as an item or a record holds it: `m` missing, `n` number, `s` text, `b` boolean, `e` error, `d` date. */
export function valueXml(v: PivotValue, more = ''): string {
  if (v === null) return `<m${more}/>`
  if (typeof v === 'number') return `<n v="${v}"${more}/>`
  if (typeof v === 'string') return `<s v="${escapeAttr(v)}"${more}/>`
  if (typeof v === 'boolean') return `<b v="${v ? 1 : 0}"${more}/>`
  if ('error' in v) return `<e v="${ERROR_TEXT[v.error]}"${more}/>`
  return `<d v="${v.date}"${more}/>`
}

type Kind = 'blank' | 'number' | 'string' | 'boolean' | 'error' | 'date'

const kindOf = (v: PivotValue): Kind => (v === null ? 'blank' : typeof v === 'number' ? 'number' : typeof v === 'string' ? 'string' : typeof v === 'boolean' ? 'boolean' : 'error' in v ? 'error' : 'date')

/** A value's `{ date }` or number for min and max, by its kind. */
const bound = (v: number | { date: string } | undefined, date: boolean) => (v === undefined ? undefined : date ? (typeof v === 'object' ? v.date : undefined) : typeof v === 'number' ? v : undefined)

/**
 * `sharedItems`: the kinds of value from the field's facts, else from its items and rows, as Excel flags them (text,
 * booleans and errors are its strings; a blank or a string makes it semi-mixed; two kinds besides blanks, mixed), the
 * least and greatest number and date, and the items.
 */
function sharedItems(f: PivotCacheField, column: PivotValue[]): string {
  const values = f.items ?? column
  const kinds = new Set<Kind>(f.facts?.types ?? values.map(kindOf))
  const str = kinds.has('string') || kinds.has('boolean') || kinds.has('error')
  const number = kinds.has('number')
  const date = kinds.has('date')
  const numbers = values.filter((v): v is number => typeof v === 'number')
  const dates = values.flatMap((v) => (v !== null && typeof v === 'object' && 'date' in v ? [v.date] : [])).sort()
  const integer = f.facts?.integer ?? (numbers.length > 0 && numbers.every(Number.isInteger))
  const min = bound(f.facts?.min, false) ?? (numbers.length ? Math.min(...numbers) : undefined)
  const max = bound(f.facts?.max, false) ?? (numbers.length ? Math.max(...numbers) : undefined)
  const minDate = bound(f.facts?.min, true) ?? dates[0]
  const maxDate = bound(f.facts?.max, true) ?? dates[dates.length - 1]
  const mixed = [str, number, date].filter(Boolean).length > 1
  const longText = f.facts?.longText ?? values.some((v) => typeof v === 'string' && v.length > 255)
  let a = ''
  if (!str && !kinds.has('blank') && kinds.size) a += ' containsSemiMixedTypes="0"'
  if (date && !str && !number) a += ' containsNonDate="0"'
  if (date) a += ' containsDate="1"'
  if (!str && kinds.size) a += ' containsString="0"'
  if (kinds.has('blank')) a += ' containsBlank="1"'
  if (mixed) a += ' containsMixedTypes="1"'
  if (number) a += ' containsNumber="1"'
  if (number && integer) a += ' containsInteger="1"'
  if (number && min !== undefined) a += ` minValue="${min}" maxValue="${max}"`
  if (date && minDate !== undefined) a += ` minDate="${minDate}" maxDate="${maxDate}"`
  if (longText) a += ' longText="1"'
  if (!f.items) return `<sharedItems${a}/>`
  const unused = new Set(f.unused ?? [])
  return `<sharedItems${a} count="${f.items.length}">${f.items.map((v, i) => valueXml(v, unused.has(i) ? ' u="1"' : '')).join('')}</sharedItems>`
}

/** A number or date of a group's range: `startNum` or `startDate`. */
const rangeValue = (name: 'start' | 'end', v: number | { date: string } | undefined) => (v === undefined ? '' : typeof v === 'number' ? ` ${name}Num="${v}"` : ` ${name}Date="${v.date}"`)

/** A field's `fieldGroup`: its parent and base, the range of dates or numbers, the manual groups, the groups' labels. */
function fieldGroupXml(g: PivotFieldGroup): string {
  let x = ''
  if (g.range) {
    const r = g.range
    x += `<rangePr${give(r, 'autoStart:t autoEnd:t')}${r.by === 'numbers' ? '' : ` groupBy="${r.by}"`}${rangeValue('start', r.start)}${rangeValue('end', r.end)}${give(r, 'groupInterval=interval:n')}/>`
  }
  if (g.discrete) x += counted('discretePr', g.discrete.map((i) => `<x v="${i}"/>`))
  if (g.items) x += counted('groupItems', g.items.map((v) => valueXml(v)))
  return `<fieldGroup${give(g, 'par=parent:n base:n')}>${x}</fieldGroup>`
}

/** The source as `cacheSource`; `link` is the relationship id of another workbook's path, or of the workbook itself for a source whose sheet is gone. */
function sourceXml(c: PivotCache, link: string | undefined): string {
  const s = c.source
  const r = link === undefined ? '' : ` r:id="${link}"`
  if (s.type === 'range') return `<cacheSource type="worksheet"><worksheetSource ref="${s.range}" sheet="${escapeAttr(s.sheet)}"${r}/></cacheSource>`
  if (s.type === 'table' || s.type === 'name') return `<cacheSource type="worksheet"><worksheetSource name="${escapeAttr(s.type === 'table' ? s.table : s.name)}"${r}/></cacheSource>`
  if (s.type === 'consolidation') {
    const sets = s.ranges.map((x) => `<rangeSet ref="${x.range}"${x.sheet === undefined ? '' : ` sheet="${escapeAttr(x.sheet)}"`}/>`)
    return `<cacheSource type="consolidation"><consolidation>${counted('rangeSets', sets)}</consolidation></cacheSource>`
  }
  if (s.type === 'scenario') return '<cacheSource type="scenario"/>'
  return `<cacheSource type="external"${s.connection === undefined ? '' : ` connectionId="${s.connection}"`}/>`
}

/** The fields' values in the rows, by field: what `sharedItems` reads its kinds from when the field lists no items. */
const columns = (c: PivotCache) => c.fields.map((_, i) => (c.fields[i].items ? [] : (c.records ?? []).map((r) => r[i] ?? null)))

/**
 * The cache's definition part. `records` is the relationship id of its rows' part, `link` that of a source in another
 * workbook; `x14` the cache's id among the workbook's slicer sources (`pivotCacheId`), when a slicer filters it.
 */
export function pivotCacheXml(c: PivotCache, numFmt: (f: PivotCacheField) => number, records: string | undefined, link: string | undefined, x14?: number): string {
  const cols = columns(c)
  const fields = c.fields.map((f, i) => {
    const added = f.formula !== undefined || (f.group?.base !== undefined && f.group.base !== i)
    const attrs = ` name="${escapeAttr(f.name)}" numFmtId="${numFmt(f)}"${f.formula === undefined ? '' : ` formula="${escapeAttr(f.formula)}"`}${added ? ' databaseField="0"' : ''}`
    return `<cacheField${attrs}>${f.formula === undefined ? sharedItems(f, cols[i]) : ''}${f.group ? fieldGroupXml(f.group) : ''}</cacheField>`
  })
  const calculated = countedIf('calculatedItems', c.calculatedItems?.map((k) => `<calculatedItem formula="${escapeAttr(k.formula)}">${pivotAreaXml({ cacheIndex: true, outline: false, fieldPosition: 0, references: [{ field: k.field, items: [k.item] }] })}</calculatedItem>`))
  const saved = records !== undefined
  const attrs =
    (saved ? ` r:id="${records}"` : ' saveData="0"') +
    (c.refreshOnLoad || c.stale ? ' refreshOnLoad="1"' : '') +
    give(c, 'enableRefresh:t createdVersion="8" refreshedVersion="8" minRefreshableVersion="3"') +
    (saved ? ` recordCount="${c.records!.length}"` : '') +
    give(c, 'missingItemsLimit=retainItems:n')
  const ext = `<extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="${X14}"><x14:pivotCacheDefinition${x14 === undefined ? '' : ` pivotCacheId="${x14}"`}/></ext></extLst>`
  return `${HEAD}<pivotCacheDefinition xmlns="${MAIN}" xmlns:r="${REL}"${attrs}>${sourceXml(c, link)}${counted('cacheFields', fields)}${calculated}${ext}</pivotCacheDefinition>`
}

/** The rows' part: each row's value per field of the source, an item's index for a field that lists its items. */
export function pivotRecordsXml(c: PivotCache): string {
  const rows = (c.records ?? []).map((r) => `<r>${r.map((v, i) => (c.fields[i]?.items && typeof v === 'number' ? `<x v="${v}"/>` : valueXml(v))).join('')}</r>`)
  return `${HEAD}<pivotCacheRecords xmlns="${MAIN}" xmlns:r="${REL}" count="${rows.length}">${rows.join('')}</pivotCacheRecords>`
}
