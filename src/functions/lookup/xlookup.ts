/** XLOOKUP and XMATCH, this engine's own. */
import { gridAt, type FnDef, type Grid } from '../../args.js'
import { wildRe } from '../../criteria.js'
import { compareValues as compare, E, isErr, sameType, type V } from '../../values.js'

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
function xfind(k: V, look: Grid, matchMode: number, searchMode: number): number {
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

function xlookup(key: V, look: Grid, ret: Grid, notFound?: V, matchMode = 0, searchMode = 1): V {
  if (isErr(key)) return key
  if (look.cols !== 1 && look.rows !== 1) return E.VALUE
  const vertical = look.cols === 1
  if (vertical ? ret.rows !== look.rows || ret.cols !== 1 : ret.cols !== look.cols || ret.rows !== 1) return E.VALUE // an array result: no spilling
  if (!MATCH_MODES.includes(matchMode) || !SEARCH_MODES.includes(searchMode)) return E.VALUE
  const pos = xfind(key, look, matchMode, searchMode)
  if (pos < 0) return notFound === undefined ? E.NA : notFound
  return gridAt(ret, pos)
}

function xmatch(key: V, look: Grid, matchMode = 0, searchMode = 1): V {
  if (isErr(key)) return key
  if ((look.cols !== 1 && look.rows !== 1) || !MATCH_MODES.includes(matchMode) || !SEARCH_MODES.includes(searchMode)) return E.VALUE
  const pos = xfind(key, look, matchMode, searchMode)
  return pos < 0 ? E.NA : pos + 1
}

export const functions: Record<string, FnDef> = {
  XLOOKUP: { args: 'a G G a? n? n?', fn: xlookup },
  XMATCH: { args: 'a G n? n?', fn: xmatch },
}
