/**
 * The selections of criteria a runtime keeps between formulas (`Cx.select`), under their first criteria range: a
 * selection's key (its size, each criteria range and its criterion), the key's hash, what a selection holds and the
 * bytes the cache's budget counts for it. The runtime makes the selections; this module finds and keeps them.
 */
import type { Crit } from '../criteria.js'
import type { RangeCache } from '../range-cache.js'
import type { Rng, V } from '../values.js'

/** Selections kept under their first criteria range, by the hash of their key (`keyHash`), several where hashes meet; their bytes by the budget's estimate. */
export interface Selections {
  byHash: Map<number, Kept | Kept[]>
  bytes: number
}

/**
 * A selection made before, by its key (its size, then each criteria range and its criterion): marked the first time;
 * from the second on, with its positions and the arrays of its ranges' values they were made on.
 */
export interface Kept {
  key: unknown[]
  flats?: WeakRef<readonly V[]>[]
  sel?: readonly number[]
}

/**
 * Bytes of a selection made once and per criteria range of its key; what its positions add, per range (its array's
 * weak reference) and per position. As measured on Node: the 2,841 selections of the synthetic book of 34,000 rows
 * held 2.17 MiB, estimated 2.04.
 */
const MARK_BYTES = 120
const KEY_RANGE_BYTES = 16
const KEPT_BYTES = 96
const WEAK_BYTES = 48
const SELECTED_BYTES = 8

/** Numbers of criteria by the criterion itself, for the hash of a selection's key. */
const critIds = new WeakMap<Crit, number>()
let crits = 0
const critId = (c: Crit) => {
  let id = critIds.get(c)
  if (id === undefined) critIds.set(c, (id = ++crits))
  return id
}

/** A selection's key: its size, then each criteria range and its criterion. */
export function keyOf(ranges: Rng[], list: Crit[], n: number): unknown[] {
  // as long as it is: pushed one by one, an array keeps room for more
  const key = new Array<unknown>(1 + 2 * ranges.length)
  key[0] = n
  for (let i = 0; i < ranges.length; i++) {
    key[1 + 2 * i] = ranges[i]
    key[2 + 2 * i] = list[i]
  }
  return key
}

/** The hash of a selection's key: criteria by their numbers, ranges by their bounds. */
function keyHash(key: unknown[]): number {
  let h = key[0] as number
  const mix = (x: number) => (h = Math.imul(h ^ x, 0x01000193))
  for (let i = 1; i < key.length; i += 2) {
    const g = key[i] as Rng
    mix(critId(key[i + 1] as Crit))
    mix(g.s)
    mix(g.r1)
    mix(g.c1)
    mix(g.r2)
    mix(g.c2)
  }
  return h
}

/** Whether two selections' keys are the same: their sizes and criteria, and ranges of the same bounds. */
function sameKey(a: unknown[], b: unknown[]): boolean {
  if (a.length !== b.length || a[0] !== b[0]) return false
  for (let i = 1; i < a.length; i += 2) if ((a[i] as Rng).key !== (b[i] as Rng).key || a[i + 1] !== b[i + 1]) return false
  return true
}

/** The selection kept for a key, made once or again, and the key's hash. */
export function keptSelection(selections: RangeCache<Selections>, first: Rng, key: unknown[]): [Kept | undefined, number] {
  const hash = keyHash(key)
  const at = selections.get(first.key)?.byHash.get(hash)
  return [Array.isArray(at) ? at.find((k) => sameKey(k.key, key)) : at && sameKey(at.key, key) ? at : undefined, hash]
}

/** A selection just made, kept: marked the first time, with its positions from the second on; its bytes in the budget. */
export function keepSelection(selections: RangeCache<Selections>, first: Rng, kept: Kept | undefined, hash: number, key: unknown[], flats: (readonly V[])[], sel: number[]): void {
  const top = selections.get(first.key) ?? { byHash: new Map(), bytes: 0 }
  const count = flats.length
  if (!kept) {
    const at = top.byHash.get(hash)
    const mark = { key }
    top.byHash.set(hash, at === undefined ? mark : Array.isArray(at) ? [...at, mark] : [at, mark])
    top.bytes += MARK_BYTES + KEY_RANGE_BYTES * count
  } else {
    if (kept.sel) top.bytes -= KEPT_BYTES + WEAK_BYTES * count + SELECTED_BYTES * kept.sel.length
    kept.flats = flats.map((f) => new WeakRef(f))
    // the positions as long as they are: pushed one by one, an array keeps room for many more
    kept.sel = sel.slice()
    top.bytes += KEPT_BYTES + WEAK_BYTES * count + SELECTED_BYTES * sel.length
  }
  selections.set(first.key, top, top.bytes)
}
