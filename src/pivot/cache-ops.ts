/**
 * The operations on a pivot cache (`editStructure`): a field grouped (dates by periods, numbers by an interval, items by
 * hand), ungrouped, a group renamed; calculated fields and items added, changed and deleted. Each acts on every
 * PivotTable of the cache, as in Excel, which are laid out again from the cache in one step.
 */
import { caseKey } from '../collate.js'
import type { BookInput, PivotCacheField, PivotFieldGroup, PivotGrouping, PivotGroupBy, PivotTable, PivotValue } from '../contract.js'
import { EPOCH_1904 } from '../dates.js'
import { InputError } from '../input-error.js'
import { checkSheetOf } from '../input.js'
import { array, checkNumber, checkOneOf, checkString, integer, object, onlyKeys } from '../input/basic.js'
import { say, type EngineMessage } from '../locale/messages.js'
import type { Applied } from '../structure/geometry.js'
import { FormulaError } from '../syntax/ast.js'
import { parse } from '../syntax/parse.js'
import { tokenize } from '../syntax/tokens.js'
import { refreshPivots } from './engine.js'
import { placeOf, placed, withFields, withoutFields, type CacheTables } from './fields.js'
import { template } from './formulas.js'
import { rangeLabels } from './group.js'
import { groupedItems } from './items.js'
import type { PivotCacheEdit, PivotCacheOp } from './hook.js'
import { itemKey } from './items.js'
import { cacheDate, dateSerial, itemText, numberOf } from './values.js'

/** A refusal's message, thrown out of an operation. */
class Refused extends Error {
  constructor(readonly m: EngineMessage) {
    super('refused')
  }
}
const refuse = (m: EngineMessage): never => {
  throw new Refused(m)
}

/** What an operation works on: the cache and its PivotTables, the one the operation names among them, the epoch. */
interface Work extends CacheTables {
  own: number
  epoch: number
}

const DATE_PERIODS = ['seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years'] as const
const PERIOD_NAMES: Record<string, string> = { seconds: 'Seconds', minutes: 'Minutes', hours: 'Hours', days: 'Days', months: 'Months', quarters: 'Quarters', years: 'Years' }

/** A field's values, record by record (its items without records). */
function valuesOf(x: CacheTables, f: number): PivotValue[] {
  const field = x.cache.fields[f]
  const records = x.cache.records
  if (!records) return field.items ?? []
  return records.map((r) => {
    const raw = r[f] ?? null
    return field.items && typeof raw === 'number' ? (field.items[raw] ?? null) : raw
  })
}

/** A source's field: neither made by grouping nor calculated. */
const sourceField = (x: CacheTables, f: number) => {
  const g = x.cache.fields[f].group
  return x.cache.fields[f].formula === undefined && (g?.base === undefined || g.base === f)
}

/** The fields a field's dates are grouped into, by periods. */
const periodFields = (x: CacheTables, base: number) => x.cache.fields.flatMap((f, i) => (f.group?.range && f.group.base === base && i !== base ? [i] : []))

/** The field of groups by hand over `base`, or −1. */
const handField = (x: CacheTables, base: number) => x.cache.fields.findIndex((f, i) => i !== base && f.group?.base === base && f.group.discrete !== undefined)

const hasCalculatedItems = (x: CacheTables, f: number) => (x.cache.calculatedItems ?? []).some((c) => c.field === f)

/** The cache field's item at a place of the operation's PivotTable's field. */
function cacheItem(w: Work, f: number, p: unknown, path: string): number {
  const items = w.tables[w.own].fields[f]?.items ?? []
  const at = integer(p, path, 0, Math.max(0, items.length - 1))
  const i = items[at]?.item
  if (i === undefined || items[at].calculated) throw new InputError(path, `expected an item of the field, not a calculated one, got ${JSON.stringify(p)}`)
  return i
}

/** Each PivotTable with `f`'s items dropped, which the next calculation lists again (its items changed meaning). */
function relisted(w: Work, f: number): void {
  w.tables = w.tables.map((t) => {
    const fields = t.fields.slice()
    const { items: _, ...rest } = fields[f]
    fields[f] = rest
    return { ...t, fields, ...(t.pages ? { pages: t.pages.map((p) => (p.field === f ? { field: f } : p)) } : {}) }
  })
}

/** A date of an operation as a cache's (`2024-01-31T00:00:00`), ISO 8601 by day or to the second. */
function dateOf(v: unknown, path: string, epoch: number): number {
  const s = checkString(v, path)
  if (!/^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2})?)?$/.test(s)) throw new InputError(path, `expected an ISO 8601 date (2024-01-31), got ${JSON.stringify(v)}`)
  return dateSerial(s.length === 10 ? `${s}T00:00:00` : s, epoch)
}

function groupDates(w: Work, f: number, g: Record<string, unknown>): void {
  onlyKeys(g, 'op.group', ['by', 'periods', 'days', 'start', 'end'])
  const periods = array(g.periods, 'op.group.periods', 'periods').map((p, i) => (checkOneOf(p, `op.group.periods[${i}]`, DATE_PERIODS), p as Exclude<PivotGroupBy, 'numbers'>))
  if (!periods.length) throw new InputError('op.group.periods', 'expected at least one period')
  const days = g.days === undefined ? 1 : integer(g.days, 'op.group.days', 1, 32767)
  if (days > 1 && (periods.length > 1 || periods[0] !== 'days')) throw new InputError('op.group.days', "a number of days groups by days alone (periods ['days'])")
  const values = valuesOf(w, f)
  if (!values.some((v) => v !== null) || values.some((v) => v !== null && (typeof v !== 'object' || !('date' in v)))) refuse({ key: 'pivotCantGroup' })
  const serials = values.flatMap((v) => (v === null ? [] : [numberOf(v, w.epoch)!]))
  const [min, max] = [Math.min(...serials), Math.max(...serials)]
  const start = g.start === undefined ? min : dateOf(g.start, 'op.group.start', w.epoch)
  // the end of whole days a day after the last, as Excel takes it (the set group: 12/22/2025 ends at 12/23/2025)
  const end = g.end === undefined ? (Number.isInteger(max) ? max + 1 : max) : dateOf(g.end, 'op.group.end', w.epoch)
  ungroupDates(w, f)
  const blank = values.includes(null)
  const order = [...new Set(periods)].sort((a, b) => DATE_PERIODS.indexOf(a) - DATE_PERIODS.indexOf(b))
  const name = w.cache.fields[f].name
  const made = order.map((by): PivotCacheField => {
    const range: NonNullable<PivotFieldGroup['range']> = { by, start: { date: cacheDate(start, w.epoch) }, end: { date: cacheDate(end, w.epoch) } }
    if (g.start !== undefined) range.autoStart = false
    if (g.end !== undefined) range.autoEnd = false
    if (days > 1) range.interval = days
    return { name: `${PERIOD_NAMES[by]} (${name})`, group: { base: f, range, items: rangeLabels(range, blank, false, w.epoch) } }
  })
  const at = withFields(w, made)
  const coarseFirst = made.map((_, i) => at + i).reverse()
  setGroup(w, f, { ...w.cache.fields[f].group, parent: coarseFirst[0] })
  // the PivotTable grouped shows the periods in place of the dates; another of the cache shows them before the dates
  w.tables = w.tables.map((t, k) => {
    const where = placeOf(t, f)
    return where ? placed(t, where, coarseFirst, k === w.own ? 1 : 0) : t
  })
}

function setGroup(w: Work, f: number, group: PivotFieldGroup | undefined): void {
  const fields = w.cache.fields.slice()
  const { group: _, ...rest } = fields[f]
  fields[f] = group && Object.keys(group).length ? { ...rest, group } : rest
  w.cache = { ...w.cache, fields }
}

function groupNumbers(w: Work, f: number, g: Record<string, unknown>): void {
  onlyKeys(g, 'op.group', ['by', 'interval', 'start', 'end'])
  const number = (v: unknown, path: string, min: number) => (checkNumber(v, path, min, Number.MAX_VALUE), v as number)
  const interval = number(g.interval, 'op.group.interval', Number.MIN_VALUE)
  const values = valuesOf(w, f)
  if (!values.some((v) => typeof v === 'number') || values.some((v) => v !== null && typeof v !== 'number')) refuse({ key: 'pivotCantGroup' })
  const numbers = values.filter((v): v is number => typeof v === 'number')
  const start = g.start === undefined ? Math.min(...numbers) : number(g.start, 'op.group.start', -Number.MAX_VALUE)
  const end = g.end === undefined ? Math.max(...numbers) : number(g.end, 'op.group.end', start)
  const range: NonNullable<PivotFieldGroup['range']> = { by: 'numbers', start, end }
  if (g.start !== undefined) range.autoStart = false
  if (g.end !== undefined) range.autoEnd = false
  if (interval !== 1) range.interval = interval
  const { parent } = w.cache.fields[f].group ?? {}
  setGroup(w, f, { ...(parent !== undefined ? { parent } : {}), base: f, range, items: rangeLabels(range, values.includes(null), numbers.every(Number.isInteger), w.epoch) })
  relisted(w, f)
}

/** The groups by hand of a field over `base`: each group's items of the base field, a group of one item of its own label alone. */
function handGroups(w: Work, g: number, base: number): { label: PivotValue; items: number[] }[] {
  const group = w.cache.fields[g].group!
  const out = (group.items ?? []).map((label) => ({ label, items: [] as number[] }))
  group.discrete!.forEach((k, i) => out[k]?.items.push(i))
  return out.filter((x) => x.items.length)
}

/** Groups by hand made again from `groups`: an item of its own first, in the base field's order, then the groups made, in theirs. */
function setHandGroups(w: Work, g: number, base: number, groups: { label: PivotValue; items: number[] }[]): void {
  const baseItems = w.cache.fields[base].items ?? []
  const own = (x: { label: PivotValue; items: number[] }) => x.items.length === 1 && itemKey(x.label) === itemKey(baseItems[x.items[0]] ?? null)
  const sorted = [...groups.filter(own).sort((a, b) => a.items[0] - b.items[0]), ...groups.filter((x) => !own(x))]
  const discrete = new Array<number>(baseItems.length).fill(0)
  sorted.forEach((x, k) => x.items.forEach((i) => (discrete[i] = k)))
  setGroup(w, g, { ...w.cache.fields[g].group, base, discrete, items: sorted.map((x) => x.label) })
  // the groups in the order of the base field's items on each PivotTable, a group at its first item's place (the set
  // group: Group1 of Apple and Banana before Cherry); a field of no items listed again by label
  w.tables = w.tables.map((t) => {
    const order = t.fields[base]?.items
    const fields = t.fields.slice()
    if (!order) {
      const { items: _, ...rest } = fields[g]
      fields[g] = rest
    } else {
      const seen = new Set<number>()
      const items = order.flatMap((x) => (x.item === undefined || seen.has(discrete[x.item]) ? [] : (seen.add(discrete[x.item]), [{ item: discrete[x.item] }])))
      fields[g] = { ...fields[g], items }
    }
    return { ...t, fields, ...(t.pages ? { pages: t.pages.map((p) => (p.field === g ? { field: g } : p)) } : {}) }
  })
}

function groupItems(w: Work, f: number, spec: Record<string, unknown>): void {
  onlyKeys(spec, 'op.group', ['by', 'items'])
  const chosen = [...new Set(array(spec.items, 'op.group.items', 'items').map((p, i) => cacheItem(w, f, p, `op.group.items[${i}]`)))]
  if (!chosen.length) throw new InputError('op.group.items', 'expected at least one item')
  if (!sourceField(w, f) || w.cache.fields[f].group?.range) refuse({ key: 'pivotCantGroup' })
  if (hasCalculatedItems(w, f)) refuse({ key: 'pivotGroupedItems' })
  const baseItems = w.cache.fields[f].items ?? []
  let g = handField(w, f)
  if (g < 0) {
    const taken = new Set(w.cache.fields.map((x) => caseKey(x.name)))
    let n = 2
    while (taken.has(caseKey(`${w.cache.fields[f].name}${n}`))) n++
    g = withFields(w, [{ name: `${w.cache.fields[f].name}${n}`, group: { base: f, discrete: baseItems.map((_, i) => i), items: baseItems.slice() } }])
    setGroup(w, f, { ...w.cache.fields[f].group, parent: g })
    // the field of groups comes before its base field on every PivotTable that shows the base
    w.tables = w.tables.map((t) => (placeOf(t, f) ? placed(t, placeOf(t, f)!, [g]) : t))
  }
  const groups = handGroups(w, g, f).map((x) => ({ ...x, items: x.items.filter((i) => !chosen.includes(i)) }))
  const labels = new Set(groups.map((x) => (typeof x.label === 'string' ? caseKey(x.label) : '')))
  let n = 1
  while (labels.has(caseKey(`Group${n}`))) n++
  setHandGroups(w, g, f, [...groups.filter((x) => x.items.length), { label: `Group${n}`, items: chosen }])
}

/** The fields of a field's dates by periods gone; the dates shown where the first of them was. */
function ungroupDates(w: Work, base: number): boolean {
  const made = periodFields(w, base)
  if (!made.length) return false
  w.tables = w.tables.map((t) => {
    const first = made.map((m) => placeOf(t, m)).find((x) => x !== null)
    return first && !placeOf(t, base) ? placed(t, first, [base]) : t
  })
  const { parent: _, ...rest } = w.cache.fields[base].group ?? {}
  setGroup(w, base, rest)
  withoutFields(w, new Set(made))
  return true
}

function ungroup(w: Work, f: number, items: unknown): void {
  const field = w.cache.fields[f]
  const g = field.group
  if (g?.discrete) {
    const base = g.base!
    const groups = handGroups(w, f, base)
    // the groups chosen, or all of them, back to an item each
    const gone = items === undefined ? null : new Set(array(items, 'op.items', 'items').map((p, i) => cacheItem(w, f, p, `op.items[${i}]`)))
    const kept = gone ? groups.filter((_, k) => !gone.has(k)) : []
    const baseItems = w.cache.fields[base].items ?? []
    const freed = groups.filter((x) => !kept.includes(x)).flatMap((x) => x.items)
    const own = (i: number) => ({ label: baseItems[i] ?? null, items: [i] })
    const next = [...kept.filter((x) => x.items.length > 1 || itemKey(x.label) !== itemKey(baseItems[x.items[0]] ?? null)), ...kept.filter((x) => x.items.length === 1 && itemKey(x.label) === itemKey(baseItems[x.items[0]] ?? null)), ...freed.map(own)]
    if (next.some((x) => x.items.length > 1 || itemKey(x.label) !== itemKey(baseItems[x.items[0]] ?? null))) return setHandGroups(w, f, base, next)
    // no group left: the field of groups goes, its base shown where it was
    w.tables = w.tables.map((t) => (placeOf(t, f) && !placeOf(t, base) ? placed(t, placeOf(t, f)!, [base]) : t))
    const { parent: _, ...rest } = w.cache.fields[base].group ?? {}
    setGroup(w, base, rest)
    return withoutFields(w, new Set([f]))
  }
  if (items !== undefined) throw new InputError('op.items', 'expected no items: only groups made by hand are ungrouped by their items')
  if (g?.range && (g.base === undefined || g.base === f)) {
    const { range: _, items: __, base: ___, ...rest } = g
    setGroup(w, f, rest)
    return relisted(w, f)
  }
  if (!ungroupDates(w, g?.range && g.base !== undefined ? g.base : f)) refuse({ key: 'pivotNotGrouped', field: field.name })
}

function renameGroup(w: Work, f: number, item: unknown, name: unknown): void {
  const g = w.cache.fields[f].group
  if (!g?.items) refuse({ key: 'pivotNotGrouped', field: w.cache.fields[f].name })
  const i = cacheItem(w, f, item, 'op.item')
  const text = checkString(name, 'op.name')
  if (!text.trim()) throw new InputError('op.name', 'expected a name, got an empty text')
  if (g!.items!.some((v, k) => k !== i && caseKey(itemText(v, w.epoch)) === caseKey(text))) refuse({ key: 'pivotItemTaken', item: text })
  setGroup(w, f, { ...g, items: g!.items!.map((v, k) => (k === i ? text : v)) })
}

/** A formula of the operation as the engine reads it, refused when it does not read or names cells. */
function checkedFormula(text: string, names: (field: string | null, name: string) => number | undefined): void {
  const t = template(text, names)
  const plain = t.text.replace(/__pivot\d+_/g, '1')
  try {
    if (tokenize(plain).some((x) => x.t === 'ref' || x.t === 'xref' || x.t === 'table' || x.t === 'xname' || x.t === '{')) refuse({ key: 'pivotFormulaRefs' })
    parse(plain)
  } catch (e) {
    if (e instanceof FormulaError) refuse({ key: 'pivotFormulaBad', formula: `=${text}` })
    throw e
  }
}

const formulaText = (v: unknown, path: string) => checkString(v, path).replace(/^\s*=/, '').trim()

function calculatedField(w: Work, name: unknown, formula: unknown): void {
  const n = checkString(name, 'op.name').trim()
  if (!n) throw new InputError('op.name', 'expected a name, got an empty text')
  const at = w.cache.fields.findIndex((f) => caseKey(f.name) === caseKey(n))
  if (formula === null) {
    if (at < 0 || w.cache.fields[at].formula === undefined) throw new InputError('op.name', `unknown calculated field ${JSON.stringify(n)}`)
    return withoutFields(w, new Set([at]))
  }
  if (at >= 0 && w.cache.fields[at].formula === undefined) refuse({ key: 'pivotNameTaken', name: n })
  const text = formulaText(formula, 'op.formula')
  const byName = new Map(w.cache.fields.map((f, i) => [caseKey(f.name), i]))
  // a field named in its own formula, through others too, would go round
  const self = at >= 0 ? at : -1
  const reaches = (f: number, seen = new Set<number>()): boolean => {
    if (f === self || f === -2) return true
    const ff = w.cache.fields[f]?.formula
    if (ff === undefined || seen.has(f)) return false
    seen.add(f)
    return template(ff, (fl, nm) => (fl === null ? byName.get(caseKey(nm)) : undefined)).args.some((x) => reaches(x, seen))
  }
  const names = (fl: string | null, nm: string) => (fl === null ? (caseKey(nm) === caseKey(n) ? -2 : byName.get(caseKey(nm))) : undefined)
  checkedFormula(text, names)
  if (template(text, names).args.some((x) => reaches(x))) refuse({ key: 'pivotFormulaBad', formula: `=${text}` })
  if (at >= 0) {
    const fields = w.cache.fields.slice()
    fields[at] = { ...fields[at], formula: text }
    w.cache = { ...w.cache, fields }
  } else withFields(w, [{ name: n, formula: text }])
}

const AVERAGES = ['average', 'stdDev', 'stdDevp', 'var', 'varp']
const PLAIN = /^[\p{L}_][\p{L}\p{N}_.]*$/u
const quotedName = (s: string) => (PLAIN.test(s) && !/^[A-Za-z]{1,3}\d+$/.test(s) ? s : `'${s.replace(/'/g, "''")}'`)

function calculatedItem(w: Work, f: number, name: unknown, formula: unknown): void {
  const n = checkString(name, 'op.name').trim()
  if (!n) throw new InputError('op.name', 'expected a name, got an empty text')
  const field = w.cache.fields[f]
  if (!sourceField(w, f) || !field.items) throw new InputError('op.field', 'expected a field of the source with items')
  if (field.group?.range || handField(w, f) >= 0) refuse({ key: 'pivotGroupedItems' })
  const own = (w.cache.calculatedItems ?? []).filter((c) => c.field === f)
  const existing = own.find((c) => caseKey(itemText(field.items![c.item] ?? null, w.epoch)) === caseKey(n))
  if (formula === null) {
    if (!existing) throw new InputError('op.name', `unknown calculated item ${JSON.stringify(n)}`)
    return withoutItem(w, f, existing.item)
  }
  if (w.tables.some((t) => (t.values ?? []).some((v) => AVERAGES.includes(v.function ?? 'sum')))) refuse({ key: 'pivotItemsAverage' })
  const calculated = new Set(own.map((c) => c.item))
  if (!existing && field.items.some((v, i) => !calculated.has(i) && caseKey(itemText(v, w.epoch)) === caseKey(n))) refuse({ key: 'pivotItemTaken', item: n })
  const labels = new Map<string, number>()
  field.items.forEach((v, i) => !calculated.has(i) && !labels.has(caseKey(itemText(v, w.epoch))) && labels.set(caseKey(itemText(v, w.epoch)), i))
  const typed = formulaText(formula, 'op.formula')
  const names = (fl: string | null, nm: string) => (fl === null || caseKey(fl) === caseKey(field.name) ? labels.get(caseKey(nm)) : undefined)
  checkedFormula(typed, names)
  // the file's form: each item by its field, Region[East], Product['Long name']
  const t = template(typed, names)
  const stored = t.text.replace(/__pivot(\d+)_/g, (_, k: string) => `${quotedName(field.name)}[${quotedName(itemText(field.items![t.args[Number(k)]] ?? null, w.epoch))}]`)
  if (existing) {
    w.cache = { ...w.cache, calculatedItems: (w.cache.calculatedItems ?? []).map((c) => (c === existing ? { ...c, formula: stored } : c)) }
    return
  }
  const item = field.items.length
  const fields = w.cache.fields.slice()
  fields[f] = { ...field, items: [...field.items, n] }
  w.cache = { ...w.cache, fields, calculatedItems: [...(w.cache.calculatedItems ?? []), { field: f, item, formula: stored }] }
  w.tables = w.tables.map((t2) => {
    const pf = t2.fields[f]
    if (!pf?.items) return t2
    const tf = t2.fields.slice()
    tf[f] = { ...pf, items: [...pf.items, { item, calculated: true }] }
    return { ...t2, fields: tf }
  })
}

/** A calculated item gone from its field: the items after it move down, on the PivotTables too. */
function withoutItem(w: Work, f: number, k: number): void {
  const down = (i: number) => (i > k ? i - 1 : i)
  const field = w.cache.fields[f]
  const fields = w.cache.fields.slice()
  fields[f] = { ...field, items: field.items!.filter((_, i) => i !== k), ...(field.unused ? { unused: field.unused.filter((i) => i !== k).map(down) } : {}) }
  const records = w.cache.records?.map((r) => (typeof r[f] === 'number' && (r[f] as number) > k ? Object.assign(r.slice(), { [f]: (r[f] as number) - 1 }) : r))
  const calculatedItems = (w.cache.calculatedItems ?? []).flatMap((c) => (c.field !== f ? [c] : c.item === k ? [] : [{ ...c, item: down(c.item) }]))
  w.cache = { ...w.cache, fields, ...(records ? { records } : {}), calculatedItems }
  w.tables = w.tables.map((t) => {
    const pf = t.fields[f]
    const place = pf?.items?.findIndex((x) => x.item === k) ?? -1
    if (!pf?.items) return t
    const tf = t.fields.slice()
    tf[f] = { ...pf, items: pf.items.flatMap((x, p) => (p === place ? [] : [x.item === undefined ? x : { ...x, item: down(x.item) }])) }
    const pages = t.pages?.map((p) => (p.field !== f || p.item === undefined ? p : p.item === place ? { field: f } : { ...p, item: p.item > place ? p.item - 1 : p.item }))
    return { ...t, fields: tf, ...(pages ? { pages } : {}) }
  })
}

function apply(w: Work, o: Record<string, unknown>): void {
  const field = (path = 'op.field') => integer(o.field, path, 0, w.cache.fields.length - 1)
  switch (o.op as PivotCacheOp['op']) {
    case 'groupPivotField': {
      const f = field()
      const g = object(o.group, 'op.group', 'a grouping')
      checkOneOf(g.by, 'op.group.by', ['dates', 'numbers', 'items'])
      const by = g.by as PivotGrouping['by']
      if (by !== 'items' && (!sourceField(w, f) || hasCalculatedItems(w, f))) refuse(hasCalculatedItems(w, f) ? { key: 'pivotGroupedItems' } : { key: 'pivotCantGroup' })
      return by === 'dates' ? groupDates(w, f, g) : by === 'numbers' ? groupNumbers(w, f, g) : groupItems(w, f, g)
    }
    case 'ungroupPivotField':
      return ungroup(w, field(), o.items)
    case 'renamePivotGroup':
      return renameGroup(w, field(), o.item, o.name)
    case 'setPivotCalculatedField':
      return calculatedField(w, o.name, o.formula)
    case 'setPivotCalculatedItem':
      return calculatedItem(w, field(), o.name, o.formula)
  }
}

export const editPivotCache: PivotCacheEdit = (book, op, values, today) => {
  const o = op as unknown as Record<string, unknown>
  const s = book.sheets.indexOf(checkSheetOf(book, o.sheet, 'op.sheet'))
  const name = checkString(o.pivot, 'op.pivot')
  const table = book.sheets[s].pivotTables?.find((t) => caseKey(t.name) === caseKey(name))
  if (!table) throw new InputError('op.pivot', `unknown PivotTable ${JSON.stringify(name)} on ${JSON.stringify(book.sheets[s].name)}`)
  const cache = (book.pivotCaches ?? []).find((c) => c.id === table.cache)!
  const spots = book.sheets.flatMap((sh, i) => (sh.pivotTables ?? []).flatMap((t, k) => (t.cache === cache.id ? [{ s: i, k, t }] : [])))
  const w: Work = { cache, tables: spots.map((x) => x.t), own: spots.findIndex((x) => x.t === table), epoch: book.dateSystem === 1904 ? EPOCH_1904 : 0 }
  try {
    apply(w, o)
  } catch (e) {
    if (e instanceof Refused) return { ok: false, code: 'pivot', message: say(book.locale, e.m) }
    throw e
  }
  const sheets = book.sheets.map((sh, i) => {
    const mine = spots.filter((x) => x.s === i)
    if (!mine.length) return sh
    const list: PivotTable[] = [...sh.pivotTables!]
    for (const x of mine) list[x.k] = w.tables[spots.indexOf(x)]
    return { ...sh, pivotTables: list }
  })
  const next: BookInput = { ...book, sheets, pivotCaches: book.pivotCaches!.map((c) => (c.id === cache.id ? w.cache : c)) }
  // laid out again from the cache, as Excel does; a cache that saved no rows is read from its source
  return refreshPivots(next, [{ sheet: book.sheets[s].name, name: table.name }], values, today, !w.cache.records) as Applied
}
