/**
 * What Excel's PivotTable Fields list does with a field, by the kinds of value its cache holds alone (no engine):
 * numbers, empty cells among them, go to Values with Sum when checked; any other field to the rows, and takes Count on
 * Values.
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
 * What Excel's PivotTable Fields list does with the field `field` (its index in `cache.fields`): checking its box puts a
 * field of numbers, with empty cells or without, or a calculated field on Values (`axis`), any other field (text, dates,
 * booleans, errors or text among numbers, empty cells alone, a field of groups, a field the cache holds no value of) on
 * the rows; on Values, checked or dragged there, the first takes Sum, the others Count (`function`).
 */
export function suggestPivotField(cache: PivotCache, field: number): PivotFieldSuggestion {
  const c = object(cache, 'cache', 'a pivot cache such as { id: 1, source: {...}, fields: [{ name: "Amount" }] }')
  const f = integer(field, 'field', 0, array(c.fields, 'cache.fields', 'the fields of the pivot cache').length - 1)
  const at = object(cache.fields[f], `cache.fields[${f}]`, 'a cache field such as { name: "Amount" }') as unknown as PivotCacheField
  const kinds = at.formula !== undefined || at.group ? null : kindsOf(cache, f)
  const numbers = at.formula !== undefined || (kinds !== null && kinds.has('number') && [...kinds].every((k) => k === 'number' || k === 'blank'))
  return numbers ? { axis: 'values', function: 'sum' } : { axis: 'rows', function: 'count' }
}
