/** XLOOKUP and XMATCH, this engine's own. */
import type { Arg, Cx, FnDef, Grid } from '../../args.js'
import { Arr } from '../../arrays/value.js'
import { wildRe } from '../../criteria.js'
import { E, isErr, Rng, sameType, type V } from '../../values.js'
import type { Compare } from './search.js'

const MATCH_MODES = [0, -1, 1, 2]
const SEARCH_MODES = [1, -1, 2, -2]

/** A blank key finds the first (from the end: the last) blank cell, past the occupied part too, as in Excel. */
function blank(look: Grid, searchMode: number): number {
  const n = look.v.length
  const size = look.rows * look.cols
  if (searchMode < 0 && size > n) return size - 1
  for (let j = 0; j < n; j++) {
    const i = searchMode > 0 ? j : n - 1 - j
    if (look.v[i] === null) return i
  }
  return searchMode > 0 && size > n ? n : -1
}

/** Match modes 0, −1, 1, 2 (wildcards, `~` escaping any character); search modes ±1, ±2 (binary search runs linearly: the same answer on sorted data). */
function xfind(k: V, look: Grid, matchMode: number, searchMode: number, compare: Compare): number {
  if (k === null) return blank(look, searchMode)
  const wild = matchMode === 2 && typeof k === 'string' ? wildRe(k, true) : null
  const n = look.v.length
  let pos = -1
  let best: V = null
  for (let j = 0; j < n; j++) {
    const i = searchMode > 0 ? j : n - 1 - j
    const v = look.v[i]
    if (v === null || isErr(v) || !sameType(v, k)) continue
    if (wild ? wild.test(v as string) : compare(v, k) === 0) {
      pos = i
      break
    }
    if (matchMode === -1 && compare(v, k) < 0 && (pos < 0 || compare(v, best) > 0)) {
      pos = i
      best = v
    } else if (matchMode === 1 && compare(v, k) > 0 && (pos < 0 || compare(v, best) < 0)) {
      pos = i
      best = v
    }
  }
  return pos
}

/** Row (column) `pos` of the return array: a reference of a reference (one cell too: SUM skips text there), else an array. */
function line(ret: Rng | Arr, pos: number, vertical: boolean): Arg {
  if (ret instanceof Arr) {
    const n = vertical ? ret.cols : ret.rows
    const v = Array.from({ length: n }, (_, i) => (vertical ? ret.at(pos, i) : ret.at(i, pos)))
    return n === 1 ? v[0] : vertical ? new Arr(1, n, 1, n, v) : new Arr(n, 1, n, 1, v)
  }
  return vertical ? new Rng(ret.s, ret.r1 + pos, ret.c1, ret.r1 + pos, ret.c2) : new Rng(ret.s, ret.r1, ret.c1 + pos, ret.r2, ret.c1 + pos)
}

/** The answer is a row (column) of the return array: several cells when it is wider (taller) than one. */
function xlookup(cx: Cx, key: V, look: Grid, ret: V | Rng | Arr, notFound?: Arg, matchMode = 0, searchMode = 1): Arg {
  if (isErr(key)) return key
  if (isErr(ret)) return ret
  if (look.cols !== 1 && look.rows !== 1) return E.VALUE
  const vertical = look.cols === 1
  const r = ret instanceof Rng || ret instanceof Arr ? ret : new Arr(1, 1, 1, 1, [ret])
  const rows = r instanceof Rng ? r.r2 - r.r1 + 1 : r.rows
  const cols = r instanceof Rng ? r.c2 - r.c1 + 1 : r.cols
  if (vertical ? rows !== look.rows : cols !== look.cols) return E.VALUE
  if (!MATCH_MODES.includes(matchMode) || !SEARCH_MODES.includes(searchMode)) return E.VALUE
  const pos = xfind(key, look, matchMode, searchMode, cx.rules.compare)
  if (pos < 0) return notFound === undefined ? E.NA : notFound
  return line(r, pos, vertical)
}

function xmatch(cx: Cx, key: V, look: Grid, matchMode = 0, searchMode = 1): V {
  if (isErr(key)) return key
  if ((look.cols !== 1 && look.rows !== 1) || !MATCH_MODES.includes(matchMode) || !SEARCH_MODES.includes(searchMode)) return E.VALUE
  const pos = xfind(key, look, matchMode, searchMode, cx.rules.compare)
  return pos < 0 ? E.NA : pos + 1
}

export const functions: Record<string, FnDef> = {
  // if_not_found is answered as it is, a reference too
  XLOOKUP: { args: 'a G R P? n? n?', array: true, cx: true, fn: xlookup },
  XMATCH: { args: 'a G n? n?', cx: true, fn: xmatch },
}
