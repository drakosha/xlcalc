/** A pivot cache's items as one key per value, and a field listed as items; core code as well as `xlcalc/pivot` takes them. */
import { caseKey } from '../collate.js'
import type { PivotCache, PivotCacheField, PivotValue } from '../contract.js'

/**
 * The key two values are one item by: text in any case is one (`North` and `north`), a date is no number, each error
 * its own.
 */
export function itemKey(v: PivotValue): string {
  if (v === null) return 'b'
  switch (typeof v) {
    case 'number':
      return `n${v}`
    case 'string':
      return `s${caseKey(v)}`
    case 'boolean':
      return v ? 'T' : 'F'
  }
  return 'error' in v ? `e${v.error}` : `d${v.date}`
}

/** A field's values listed as items in the order they first come, the records (copies the caller owns) turned to their indexes. */
export function listedItems(field: PivotCacheField, records: PivotValue[][], at: number): PivotCacheField {
  const items: PivotValue[] = []
  const seen = new Map<string, number>()
  for (const r of records) {
    const v = r[at] ?? null
    const k = itemKey(v)
    let i = seen.get(k)
    if (i === undefined) {
      i = items.length
      seen.set(k, i)
      items.push(v)
    }
    r[at] = i
  }
  return { ...field, items }
}

/** The cache with the fields `fields` listed as items, its records copied when one is listed anew; the cache itself when all are. */
export function withListed(cache: PivotCache, fields: Iterable<number>): PivotCache {
  const unlisted = [...fields].filter((f) => cache.fields[f] && !cache.fields[f].items)
  if (!unlisted.length) return cache
  const records = (cache.records ?? []).map((r) => r.slice())
  const out = cache.fields.slice()
  for (const f of unlisted) out[f] = listedItems(out[f], records, f)
  return { ...cache, fields: out, records }
}
