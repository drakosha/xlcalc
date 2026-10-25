/**
 * Columns and rows of a sheet's layout, kept by span as xlsx keeps columns: a key is one line (`C`, `5`) or
 * a span of lines (`B:D`, `5:9`); spans do not overlap. Indexes are zero-based and spans inclusive.
 */
import type { LinePatch } from '../contract.js'
import { InputError, what } from '../input-error.js'
import { colToIndex, indexToCol } from '../model.js'
import { LIMIT, type Axis } from '../structure/area.js'

/** A span with its entry. */
export type Line<T> = [lo: number, hi: number, entry: T]

const KEY: Record<Axis, RegExp> = { c: /^([A-Z]{1,3})(?::([A-Z]{1,3}))?$/, r: /^([1-9][0-9]{0,6})(?::([1-9][0-9]{0,6}))?$/ }
const index = (text: string, axis: Axis) => (axis === 'c' ? colToIndex(text) : Number(text) - 1)
const text = (i: number, axis: Axis) => (axis === 'c' ? indexToCol(i) : String(i + 1))

/** The lines of a key, or null when it is not one: malformed, reversed, a span of one line (`C:C`), past the sheet. */
export function parseSpan(key: string, axis: Axis): [number, number] | null {
  const m = KEY[axis].exec(key)
  if (!m) return null
  const lo = index(m[1], axis)
  const hi = m[2] === undefined ? lo : index(m[2], axis)
  return hi < LIMIT[axis] && (m[2] === undefined || lo < hi) ? [lo, hi] : null
}

/** What a key of each axis is, for a message. */
export const SPAN_KEY = { c: 'letters such as "C" or a span such as "B:D"', r: 'a row number such as "5" or a span such as "5:9"' } as const

/** The lines of a key of `columns` or `rows` (`C`, `B:D`; `5`, `5:9`), else `InputError` naming `path`. */
export function checkSpan(key: unknown, path: string, axis: Axis): [number, number] {
  const sp = typeof key === 'string' ? parseSpan(key, axis) : null
  if (!sp) throw new InputError(path, `expected ${SPAN_KEY[axis]} within the sheet, got ${what(key)}`)
  return sp
}

/** The key of lines [lo, hi]: one line, or a span. */
export const spanKey = (lo: number, hi: number, axis: Axis) => (lo === hi ? text(lo, axis) : `${text(lo, axis)}:${text(hi, axis)}`)

/** The entries of a map in order (keys are taken as valid). */
export function lines<T>(map: Record<string, T> | undefined, axis: Axis): Line<T>[] {
  if (!map) return []
  return Object.entries(map)
    .map(([k, v]): Line<T> => [...parseSpan(k, axis)!, v])
    .sort((a, b) => a[0] - b[0])
}

/** The map of entries, keyed by span. */
export function toMap<T>(ls: Line<T>[], axis: Axis): Record<string, T> {
  const out: Record<string, T> = {}
  for (const [lo, hi, v] of [...ls].sort((a, b) => a[0] - b[0])) out[spanKey(lo, hi, axis)] = v
  return out
}

/** The entry of line `i` among entries. */
export function entryAt<T>(ls: Line<T>[], i: number): T | undefined {
  const e = ls[firstEnding(ls, i)]
  return e && e[0] <= i ? e[2] : undefined
}

/** The index of the first of entries `ls` (in order) that ends at or after line `i`; their length for none. */
export function firstEnding(ls: Line<unknown>[], i: number): number {
  // the entries do not overlap, so their ends are in order too
  let [a, b] = [0, ls.length]
  while (a < b) {
    const m = (a + b) >> 1
    if (ls[m][1] < i) a = m + 1
    else b = m
  }
  return a
}

/** The entry of line `i` of a map. */
export const lineAt = <T>(map: Record<string, T> | undefined, axis: Axis, i: number): T | undefined => entryAt(lines(map, axis), i)

/** A line's entry hidden as Excel takes it: marked hidden, or of size 0 (a row's its height, a column's its width). */
export const isHidden = (e: { hidden?: boolean; height?: number; width?: number }): boolean => e.hidden === true || e.height === 0 || e.width === 0

/** The entry with the patch applied; undefined when nothing is left. */
function patched<T extends object>(entry: T | undefined, patch: LinePatch<T>): T | undefined {
  const out: Record<string, unknown> = { ...entry }
  for (const [k, v] of Object.entries(patch)) {
    if (v === null) delete out[k]
    else if (v !== undefined) out[k] = v
  }
  return Object.keys(out).length ? (out as T) : undefined
}

/** The map with `patch` applied to every line of [lo, hi]: spans it cuts are split, entries left empty are dropped. */
export const patchLines = <T extends object>(map: Record<string, T> | undefined, axis: Axis, lo: number, hi: number, patch: LinePatch<T>): Record<string, T> =>
  toMap(patchIn(lines(map, axis), lo, hi, patch), axis)

/**
 * Entries `ls` (in order) with `patch` applied to every line of [lo, hi] as `patchLines` applies it, changed in place
 * (their tuples are replaced, never changed): many patches of one map parse and sort it once.
 */
export function patchIn<T extends object>(ls: Line<T>[], lo: number, hi: number, patch: LinePatch<T>): Line<T>[] {
  const i = firstEnding(ls, lo)
  let j = i
  const out: Line<T>[] = []
  let next = lo
  const put = (a: number, b: number, v: T | undefined) => v && out.push([a, b, v])
  for (; ls[j]?.[0] <= hi; j++) {
    const [a, b, v] = ls[j]
    // a gap before this span inside [lo, hi] gets the patch alone
    if (a > next) put(next, a - 1, patched(undefined, patch))
    if (a < lo) out.push([a, lo - 1, v])
    put(Math.max(a, lo), Math.min(b, hi), patched(v, patch))
    if (b > hi) out.push([hi + 1, b, v])
    next = b + 1
  }
  if (next <= hi) put(next, hi, patched(undefined, patch))
  // a long run of new entries is not spread into the arguments of one call
  if (out.length < 4096) ls.splice(i, j - i, ...out)
  else for (const x of out.concat(ls.splice(i).slice(j - i))) ls.push(x)
  return ls
}

/** The entries of `ls` clipped to [lo, hi], moved by `d`. */
export const slice = <T>(ls: Line<T>[], lo: number, hi: number, d = 0): Line<T>[] =>
  ls.filter(([a, b]) => b >= lo && a <= hi).map(([a, b, v]): Line<T> => [Math.max(a, lo) + d, Math.min(b, hi) + d, v])

/** The entries of `ls` without lines [lo, hi]: spans it cuts are split. */
export const cut = <T>(ls: Line<T>[], lo: number, hi: number): Line<T>[] => [...slice(ls, 0, lo - 1), ...slice(ls, hi + 1, Number.MAX_SAFE_INTEGER)]
