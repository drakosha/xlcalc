/** XLOOKUP and XMATCH, this engine's own. */
import type { Arg, Cx, FnDef, Grid } from '../../args.js'
import { Arr, asArray } from '../../arrays/value.js'
import { E, isErr, Rng, type CErr, type V } from '../../values.js'
import { regexTest } from '../../regex-level.js'
import { exactMatch, type Compare, type Key } from './search.js'

const MATCH_MODES = [0, -1, 1, 2, 3]
const SEARCH_MODES = [1, -1, 2, -2]

/**
 * Match mode 3, a regular expression found anywhere in a text, from the end with search mode -1 (Excel for Mac;
 * Excel for Windows 16.0.20430 has no such mode, #VALUE!): JavaScript's syntax with PCRE2's commonest constructs, or
 * Excel's with a level of `xlcalc/regex` (`regex-level.ts`). A pattern that is none, a binary search, and Excel's match
 * limit reached are #VALUE!. The key and the values are read as text (42 finds 42, a blank is empty text).
 */
function regex(cx: Cx, k: Key | null, look: Grid, searchMode: number): number | CErr {
  if (Math.abs(searchMode) === 2) return E.VALUE
  const test = regexTest(k === null ? '' : (cx.rules.toText(k) as string))
  if (!test) return E.VALUE
  const n = look.v.length
  const size = look.rows * look.cols
  // a blank is empty text, past the occupied part too (`^$` finds the first blank)
  const blanks = test('')
  if (blanks === 'limit') return E.VALUE
  if (blanks && searchMode < 0 && size > n) return size - 1
  for (let j = 0; j < n; j++) {
    const i = searchMode > 0 ? j : n - 1 - j
    const v = look.v[i]
    if (isErr(v)) continue
    const hit = test(v === null ? '' : typeof v === 'string' ? v : (cx.rules.toText(v) as string))
    if (hit === 'limit') return E.VALUE
    if (hit) return i
  }
  return blanks && searchMode > 0 && size > n ? n : -1
}

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

/** A value of the data that match modes −1 and 1 and the binary search compare with the key: no blank and no error. */
const present = (v: V) => v !== null && !isErr(v)

/**
 * Excel's binary search (search mode 2, −2 on descending data) over every position of the reference: it compares at
 * the middle of the positions left, both ends included, across types (numbers < text < booleans), a blank or an error
 * greater than any key; an equal value is the match, the first of its run (the last on descending data), else match
 * mode −1 takes the next smaller value and 1 the next larger, as far as the search came (`XLOOKUP(12,{"a","b"},{2,3},,1,2)`
 * is 2, and past a whole column's last value the next larger is its last row).
 */
function binary(k: Key, look: Grid, matchMode: number, desc: boolean, compare: Compare): number {
  const n = look.rows * look.cols
  const order = (i: number) => {
    const v = i < look.v.length ? look.v[i] : null
    return present(v) ? compare(v, k) : 1
  }
  let lo = 0
  let hi = n - 1
  let found = -1
  while (lo <= hi) {
    const mid = (lo + hi) >> 1
    const c = order(mid)
    if (c === 0) {
      found = mid
      if (desc) lo = mid + 1
      else hi = mid - 1
    } else if (c < 0 !== desc) lo = mid + 1
    else hi = mid - 1
  }
  if (found >= 0 || matchMode === 0) return found
  // the search ends between hi and lo: the smaller value at hi ascending, at lo descending
  const j = (matchMode === -1) !== desc ? hi : lo
  return j >= 0 && j < n ? j : -1
}

/** Match modes 0, −1, 1, 2 (wildcards, `~` escaping any character; 3 is `regex`); search modes ±1, ±2 (binary, which takes no wildcards). */
function xfind(k: Key | null, look: Grid, matchMode: number, searchMode: number, compare: Compare): number | CErr {
  if (k === null) return blank(look, searchMode)
  if (Math.abs(searchMode) === 2) return matchMode === 2 ? E.VALUE : binary(k, look, matchMode, searchMode < 0, compare)
  if (matchMode === 0 || matchMode === 2)
    return exactMatch({ at: (i) => look.v[i], n: look.rows * look.cols, filled: look.v.length, values: look.v, stride: 1 }, k, matchMode === 2, compare, searchMode < 0)
  // the next smaller or larger value across types, numbers < text < booleans, as the binary search orders them
  const n = look.v.length
  let pos = -1
  let best: V = null
  for (let j = 0; j < n; j++) {
    const i = searchMode > 0 ? j : n - 1 - j
    const v = look.v[i]
    if (!present(v)) continue
    // one comparison with the key: below it is a candidate for −1, above it for 1
    const c = compare(v, k)
    if (c === 0) {
      pos = i
      break
    }
    if (matchMode === -1 ? c < 0 && (pos < 0 || compare(v, best) > 0) : c > 0 && (pos < 0 || compare(v, best) < 0)) {
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
  const r = ret instanceof Rng ? ret : asArray(ret)
  const rows = r instanceof Rng ? r.r2 - r.r1 + 1 : r.rows
  const cols = r instanceof Rng ? r.c2 - r.c1 + 1 : r.cols
  // one value is a row and a column at once: it looks up the return array's one row, else its one column (`XLOOKUP(0,0,A1:A6)` is A1:A6)
  const vertical = look.rows === 1 && look.cols === 1 ? rows === 1 : look.cols === 1
  if (vertical ? rows !== look.rows : cols !== look.cols) return E.VALUE
  if (!MATCH_MODES.includes(matchMode) || !SEARCH_MODES.includes(searchMode)) return E.VALUE
  const pos = matchMode === 3 ? regex(cx, key, look, searchMode) : xfind(key, look, matchMode, searchMode, cx.rules.compare)
  if (isErr(pos)) return pos
  if (pos < 0) return notFound === undefined ? E.NA : notFound
  return line(r, pos, vertical)
}

function xmatch(cx: Cx, key: V, look: Grid, matchMode = 0, searchMode = 1): V {
  if (isErr(key)) return key
  if ((look.cols !== 1 && look.rows !== 1) || !MATCH_MODES.includes(matchMode) || !SEARCH_MODES.includes(searchMode)) return E.VALUE
  const pos = matchMode === 3 ? regex(cx, key, look, searchMode) : xfind(key, look, matchMode, searchMode, cx.rules.compare)
  return isErr(pos) ? pos : pos < 0 ? E.NA : pos + 1
}

export const functions: Record<string, FnDef> = {
  // if_not_found is answered as it is, a reference too
  XLOOKUP: { args: 'a G R P? n? n?', array: true, cx: true, fn: xlookup },
  XMATCH: { args: 'a G n? n?', cx: true, fn: xmatch },
}
