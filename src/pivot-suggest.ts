/**
 * Where Excel's PivotTable Fields list puts a field whose box is checked, by the kinds of value its cache holds alone
 * (no engine): numbers, empty cells among them, go to Values with Sum; any other field to the rows.
 */
import type { PivotCache, PivotCacheField, PivotFieldFacts, PivotFieldSuggestion, PivotValue } from './contract.js'
import { array, integer, object } from './input/basic.js'

type Kind = PivotFieldFacts['types'][number]

const kindOf = (v: PivotValue): Kind =>
  v === null ? 'blank' : typeof v === 'number' ? 'number' : typeof v === 'string' ? 'string' : typeof v === 'boolean' ? 'boolean' : 'error' in v ? 'error' : 'date'

/** The kinds of the field `f`: its facts, else its items', else its rows' in the cache; none when the cache holds no value of it. */
function kindsOf(cache: PivotCache, f: number): Set<Kind> {
  const field = cache.fields[f]
  if (field.facts) return new Set(field.facts.types)
  if (field.items) return new Set(field.items.map(kindOf))
  return new Set((cache.records ?? []).flatMap((r) => (r[f] === undefined ? [] : [kindOf(r[f])])))
}

/**
 * Where checking the box of the field `field` (its index in `cache.fields`) in Excel's PivotTable Fields list puts it: a
 * field of numbers, with empty cells or without, on Values with Sum, as a calculated field; any other field (text,
 * dates, booleans, errors or text among numbers, empty cells alone, a field of groups, a field the cache holds no value
 * of) on the rows.
 */
export function suggestPivotField(cache: PivotCache, field: number): PivotFieldSuggestion {
  const c = object(cache, 'cache', 'a pivot cache such as { id: 1, source: {...}, fields: [{ name: "Amount" }] }')
  const f = integer(field, 'field', 0, array(c.fields, 'cache.fields', 'the fields of the pivot cache').length - 1)
  const at = object(cache.fields[f], `cache.fields[${f}]`, 'a cache field such as { name: "Amount" }') as unknown as PivotCacheField
  if (at.formula !== undefined) return { axis: 'values', function: 'sum' }
  if (at.group) return { axis: 'rows' }
  const kinds = kindsOf(cache, f)
  if (!kinds.has('number') || [...kinds].some((k) => k !== 'number' && k !== 'blank')) return { axis: 'rows' }
  return { axis: 'values', function: 'sum' }
}
