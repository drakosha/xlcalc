/** Exact searches of an immutable line: numeric/boolean keys, and text groups ordered by the book's comparator. */
import { spend } from '../../exec/budget.js'
import type { V } from '../../values.js'
import type { Compare, Key } from './search.js'

interface Positions { first: number; last: number }
interface TextPosition extends Positions { value: string }
interface Index {
  scalar?: Map<number | boolean, Positions> | false
  text?: TextPosition[]
  scalarWork: number
  textWork: number
}

// Range arrays are immutable and replaced on invalidation; their indexes disappear with them.
const caches = new WeakMap<readonly V[], WeakMap<Compare, Map<string, Index>>>()
const MIN_LENGTH = 256
const STRIDE = 0xffff

function indexFor(values: readonly V[], stride: number, length: number, compare: Compare): Index {
  let comparators = caches.get(values)
  if (!comparators) caches.set(values, comparators = new WeakMap())
  let lines = comparators.get(compare)
  if (!lines) comparators.set(compare, lines = new Map())
  const id = `${stride}:${length}`
  let index = lines.get(id)
  if (!index) {
    lines.set(id, index = { scalarWork: 0, textWork: 0 })
  }
  return index
}

/** Remember actual scan work: several hits at the start of a long range should not trigger a full index build. */
export function scannedExact(values: readonly V[], stride: number, length: number, key: Key, compare: Compare, work: number): void {
  if (length < MIN_LENGTH || (typeof key === 'number' && !Number.isFinite(key))) return
  const index = indexFor(values, stride, length, compare)
  if (typeof key === 'string') index.textWork += work
  else index.scalarWork += work
}

/** Undefined means the caller scans; -1 is an indexed miss. `stride` uses the stored grid width, not its referenced width. */
export function indexedExact(values: readonly V[], stride: number, length: number, key: Key, compare: Compare, last: boolean): number | undefined {
  if (length < MIN_LENGTH || (typeof key === 'number' && !Number.isFinite(key))) return undefined
  const index = indexFor(values, stride, length, compare)
  let position: Positions | undefined
  if (typeof key === 'string') {
    // Sorting text costs more than a scalar map: wait for several scans' worth of comparisons.
    if (!index.text && index.textWork < length * 4) return undefined
    index.text ??= textIndex(values, stride, length, compare)
    let lo = 0, hi = index.text.length - 1
    while (lo <= hi) {
      const mid = (lo + hi) >>> 1
      const entry = index.text[mid]
      const c = compare(entry.value, key)
      if (c === 0) { position = entry; break }
      if (c < 0) lo = mid + 1
      else hi = mid - 1
    }
  } else {
    if (index.scalar === undefined) {
      if (index.scalarWork < length) return undefined
      const map = new Map<number | boolean, Positions>()
      for (let i = 0; i < length; i++) {
        if ((i & STRIDE) === 0) spend(Math.min(STRIDE + 1, length - i))
        const value = values[i * stride]
        if (value === undefined || (typeof value === 'number' && !Number.isFinite(value))) { index.scalar = false; break }
        if (typeof value !== 'number' && typeof value !== 'boolean') continue
        const found = map.get(value)
        if (found) found.last = i
        else map.set(value, { first: i, last: i })
      }
      index.scalar ??= map
    }
    if (index.scalar === false) return undefined
    position = index.scalar.get(key)
  }
  return position ? (last ? position.last : position.first) : -1
}

/** Collator-equivalent spellings share their first/last positions, without assuming lower-case text is an equality key. */
function textIndex(values: readonly V[], stride: number, length: number, compare: Compare): TextPosition[] {
  const distinct = new Map<string, TextPosition>()
  for (let i = 0; i < length; i++) {
    if ((i & STRIDE) === 0) spend(Math.min(STRIDE + 1, length - i))
    const value = values[i * stride]
    if (typeof value !== 'string') continue
    const found = distinct.get(value)
    if (found) found.last = i
    else distinct.set(value, { value, first: i, last: i })
  }
  const sorted = [...distinct.values()].sort((a, b) => { spend(1); return compare(a.value, b.value) })
  const out: TextPosition[] = []
  for (const entry of sorted) {
    const previous = out[out.length - 1]
    if (previous && compare(previous.value, entry.value) === 0) {
      previous.first = Math.min(previous.first, entry.first)
      previous.last = Math.max(previous.last, entry.last)
    } else out.push(entry)
  }
  return out
}
