/**
 * A pivot cache's parts as the workbook's `PivotCache`: the definition (`pivotCacheDefinition`: its source, fields with
 * their items, facts, groups and formulas, calculated items) and the records it saved (`pivotCacheRecords`). The values
 * of items and records are read by one scan of their tags, as a cache can hold many.
 */
import type { NumberFormat, PivotCache, PivotCacheField, PivotFieldFacts, PivotFieldGroup, PivotGroupBy, PivotSource, PivotValue } from '../../contract.js'
import { attr, attributes, attrsAt, element, flag, has, within } from '../xml/scan.js'
import { decodeEscapes, decodeXml } from '../xml/text.js'
import { ERRORS } from './cells.js'
import { modelDate } from '../../pivot-time.js'

/** The tags of values (`<s v="North"/>`, `<x v="2"/>`, `<m/>`), and with `r` a record's. */
const VALUES = /<(?:[\w.-]+:)?(r|[xnsbedm])((?:\s(?:"[^"]*"|'[^']*'|[^>"'])*?)?)(\/?)>/g

/** An attribute's text, its entities and escapes decoded. */
export const attrText = (v: string | undefined): string => decodeEscapes(decodeXml(v ?? ''))

/**
 * Attributes set on `out` by a table of entries `attr=key:kind` (the key the attribute's name without `=key`): `s` text
 * (the default), `n` a number, `p` a number above 0, `t` a flag true by default, set where false, `f` a flag false by
 * default, set where true. An attribute absent sets nothing.
 */
export function take<T extends object>(out: T, a: Record<string, string>, spec: string): T {
  const o = out as Record<string, unknown>
  for (const e of spec.split(' ')) {
    const [name, kind = 's'] = e.split(':')
    const [attr, key = attr] = name.split('=')
    const v = a[attr]
    if (v === undefined) continue
    if (kind === 's') o[key] = attrText(v)
    else if (kind === 'n' || (kind === 'p' && Number(v) > 0)) o[key] = Number(v)
    else if (kind !== 'p' && flag(v) !== (kind === 't')) o[key] = kind === 'f'
  }
  return out
}

/** The value of a tag of `VALUES` by its name and attributes; for `x`, the item's index. */
function valueOf(k: string, a: string): PivotValue {
  if (k === 'm') return null
  const v = attributes(a).v
  if (k === 'n' || k === 'x') return Number(v ?? 0)
  if (k === 'b') return flag(v)
  if (k === 'e') return { error: ERRORS[v ?? ''] ?? 'NA' }
  if (k === 'd') return { date: modelDate(v ?? '') }
  return attrText(v)
}

/** The values of the tags in `xml`, in order (records' tags skipped); with `unused`, the indexes of those marked `u` (no longer in the source). */
export function values(xml: string, unused?: number[]): PivotValue[] {
  const out: PivotValue[] = []
  VALUES.lastIndex = 0
  for (let m = VALUES.exec(xml); m; m = VALUES.exec(xml))
    if (m[1] !== 'r') {
      if (unused && flag(attributes(m[2]).u)) unused.push(out.length)
      out.push(valueOf(m[1], m[2]))
    }
  return out
}

/** The records part's rows: each `<r>` with its values, an index for `<x>`. */
export function readRecords(xml: string): PivotValue[][] {
  const out: PivotValue[][] = []
  let row: PivotValue[] | undefined
  VALUES.lastIndex = 0
  for (let m = VALUES.exec(xml); m; m = VALUES.exec(xml)) {
    if (m[1] === 'r') out.push((row = []))
    else row?.push(valueOf(m[1], m[2]))
  }
  return out
}

const numberOrDate = (n: string | undefined, d: string | undefined): number | { date: string } | undefined => (d !== undefined ? { date: modelDate(d) } : n !== undefined ? Number(n) : undefined)

/** A field's facts from its `sharedItems` flags (`containsString` and the two others true by default) and its items. */
function factsOf(a: Record<string, string>, items: PivotValue[]): PivotFieldFacts | undefined {
  const types = new Set<PivotFieldFacts['types'][number]>()
  if (flag(a.containsBlank)) types.add('blank')
  if (flag(a.containsString, true) && flag(a.containsSemiMixedTypes, true)) types.add('string')
  if (flag(a.containsNumber)) types.add('number')
  if (flag(a.containsDate)) types.add('date')
  for (const v of items) if (typeof v === 'boolean') types.add('boolean')
  for (const v of items) if (v !== null && typeof v === 'object' && 'error' in v) types.add('error')
  const facts: PivotFieldFacts = { types: [...types] }
  take(facts, a, 'containsInteger=integer:f longText:f')
  const min = numberOrDate(a.minValue, a.minDate)
  const max = numberOrDate(a.maxValue, a.maxDate)
  if (min !== undefined) facts.min = min
  if (max !== undefined) facts.max = max
  return facts
}

const GROUP_BY: Record<string, PivotGroupBy> = { range: 'numbers', seconds: 'seconds', minutes: 'minutes', hours: 'hours', days: 'days', months: 'months', quarters: 'quarters', years: 'years' }

function groupOf(xml: string): PivotFieldGroup {
  const a = attrsAt(xml, 'fieldGroup')
  const g: PivotFieldGroup = {}
  take(g, a, 'base:n par=parent:n')
  const r = element(xml, 'rangePr')
  if (r) {
    const ra = attributes(r.attrs)
    const range: NonNullable<PivotFieldGroup['range']> = { by: GROUP_BY[ra.groupBy ?? 'range'] ?? 'numbers' }
    const start = numberOrDate(ra.startNum, ra.startDate)
    const end = numberOrDate(ra.endNum, ra.endDate)
    if (start !== undefined) range.start = start
    if (end !== undefined) range.end = end
    take(range, ra, 'autoStart:t autoEnd:t')
    if (ra.groupInterval !== undefined && Number(ra.groupInterval) !== 1) range.interval = Number(ra.groupInterval)
    g.range = range
  }
  const d = element(xml, 'discretePr')
  if (d) g.discrete = values(d.body ?? '') as number[]
  const items = element(xml, 'groupItems')
  if (items) g.items = values(items.body ?? '')
  return g
}

function fieldOf(xml: string, attrs: string, format: (id: string | undefined) => NumberFormat | undefined): PivotCacheField {
  const a = attributes(attrs)
  const f: PivotCacheField = { name: attrText(a.name) }
  const nf = format(a.numFmtId)
  if (nf !== undefined) f.numberFormat = nf
  const shared = element(xml, 'sharedItems')
  const unused: number[] = []
  const items = shared?.body ? values(shared.body, shared.body.includes(' u=') ? unused : undefined) : []
  if (items.length) f.items = items
  if (unused.length) f.unused = unused
  if (shared) f.facts = factsOf(attributes(shared.attrs), items)
  if (a.formula !== undefined) f.formula = attrText(a.formula)
  if (has(xml, 'fieldGroup')) f.group = groupOf(xml)
  return f
}

/** Where the cache's data come from; `book` names a link's target by a relationship's id, `connection` tells the data model's connection. */
function sourceOf(xml: string, book: (id: string | undefined) => string | undefined, model: (id: number) => boolean): PivotSource {
  const s = attrsAt(xml, 'cacheSource')
  const linked = (a: Record<string, string>) => {
    const b = book(attr(a, 'id'))
    return b === undefined ? {} : { book: b }
  }
  if (s.type === 'external') {
    const id = s.connectionId === undefined ? undefined : Number(s.connectionId)
    return { type: id !== undefined && model(id) ? 'dataModel' : 'external', ...(id === undefined ? {} : { connection: id }) }
  }
  if (s.type === 'scenario') return { type: 'scenario' }
  if (s.type === 'consolidation') {
    const ranges = within(xml, 'rangeSets', 'rangeSet').map((e) => {
      const a = attributes(e.attrs)
      return { range: attrText(a.ref), ...(a.sheet === undefined ? {} : { sheet: attrText(a.sheet) }), ...linked(a) }
    })
    return { type: 'consolidation', ranges }
  }
  const w = attrsAt(xml, 'worksheetSource')
  if (w.name !== undefined) return { type: 'name', name: attrText(w.name), ...linked(w) }
  return { type: 'range', sheet: attrText(w.sheet), range: attrText(w.ref), ...linked(w) }
}

/** What the reader of a cache needs of the workbook: number formats by id, link targets, which connections are the data model's. */
export interface CacheContext {
  format: (id: string | undefined) => NumberFormat | undefined
  book: (id: string | undefined) => string | undefined
  model: (id: number) => boolean
}

/** A cache definition's part and its records' as the workbook's cache of `id`. */
export function readPivotCache(id: number, xml: string, records: string | undefined, ctx: CacheContext): PivotCache {
  const a = attrsAt(xml, 'pivotCacheDefinition')
  const fields = within(xml, 'cacheFields', 'cacheField').map((e) => fieldOf(e.body ?? '', e.attrs, ctx.format))
  const cache: PivotCache = { id, source: sourceOf(xml, ctx.book, ctx.model), fields }
  if (records !== undefined) cache.records = readRecords(records)
  const calculated = within(xml, 'calculatedItems', 'calculatedItem').flatMap((e) => {
    const r = element(e.body ?? '', 'reference')
    const item = r && values(r.body ?? '')[0]
    return r && typeof item === 'number' ? [{ field: Number(attributes(r.attrs).field ?? 0), item, formula: attrText(attributes(e.attrs).formula) }] : []
  })
  if (calculated.length) cache.calculatedItems = calculated
  return take(cache, a, 'refreshOnLoad:f saveData:t enableRefresh:t missingItemsLimit=retainItems:n')
}
