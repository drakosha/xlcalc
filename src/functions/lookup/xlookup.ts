/** XLOOKUP and XMATCH, this engine's own. */
import { gridAt, type FnDef, type Grid } from '../../args.js'
import { matches, parseCrit, type Crit } from '../../criteria.js'
import { compareValues as compare, E, isErr, sameType, type V } from '../../values.js'
import type { Key } from './search.js'

const MATCH_MODES = [0, -1, 1, 2]
const SEARCH_MODES = [1, -1, 2, -2]

/** Match modes 0, −1, 1, 2 (wildcards); search modes ±1, ±2 (binary search runs linearly: the same answer on sorted data). */
function xfind(k: Key, look: Grid, matchMode: number, searchMode: number): number {
  const wild: Crit | null = matchMode === 2 && typeof k === 'string' ? parseCrit('=' + k) : null
  const n = look.v.length
  let pos = -1
  let best: V = null
  for (let j = 0; j < n; j++) {
    const i = searchMode > 0 ? j : n - 1 - j
    const v = look.v[i]
    if (v === null || isErr(v) || !sameType(v, k)) continue
    if (wild ? matches(v, wild) : compare(v, k) === 0) {
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
  const pos = xfind(key === null ? 0 : key, look, matchMode, searchMode)
  if (pos < 0) return notFound === undefined ? E.NA : notFound
  return gridAt(ret, pos)
}

function xmatch(key: V, look: Grid, matchMode = 0, searchMode = 1): V {
  if (isErr(key)) return key
  if ((look.cols !== 1 && look.rows !== 1) || !MATCH_MODES.includes(matchMode) || !SEARCH_MODES.includes(searchMode)) return E.VALUE
  const pos = xfind(key === null ? 0 : key, look, matchMode, searchMode)
  return pos < 0 ? E.NA : pos + 1
}

export const functions: Record<string, FnDef> = {
  XLOOKUP: { args: 'a G G a? n? n?', fn: xlookup },
  XMATCH: { args: 'a G n? n?', fn: xmatch },
}
