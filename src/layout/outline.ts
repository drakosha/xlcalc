/**
 * Excel's outline over one axis of a sheet, as Excel for Windows keeps it (the outline sets, `docs/outline.md`): the
 * groups drawn from the lines' outline levels, the collapsed marks Excel keeps on summary lines, and the lines its
 * commands change. Indexes are zero-based. Engine-free: the `./a1` subpath gives the groups, the writer the marks.
 *
 * A summary line follows lines deeper than it (below or right; before them with summaries above or left); its detail is
 * the run of those deeper lines. Excel marks a summary line collapsed whenever its whole detail is hidden, however it got
 * hidden, and drops the mark as soon as a line of it shows, but for one thing: a summary line inside a group whose
 * lines are all hidden keeps the mark it had, which is how an inner group stays collapsed while an outer one closes and
 * opens again. A group touching the sheet's edge has no summary line; its mark is on its line at the edge.
 */
import type { LinePatch } from '../contract.js'
import { LIMIT, type Axis } from '../structure/area.js'
import { entryAt, lines, patchLines, type Line } from './spans.js'

/** What the outline reads of a line: its level, hidden state (a size of 0 hides it too) and mark. */
export interface OutlineEntry {
  outlineLevel?: number
  hidden?: boolean
  collapsed?: boolean
  height?: number
  width?: number
}

type Map<T extends OutlineEntry = OutlineEntry> = Record<string, T>

/** Excel's deepest outline level. */
export const MAX_OUTLINE = 7

/** A group: lines `lo`..`hi` at `level` or deeper; its summary line (null past the sheet's edge); collapsed as Excel reads it. */
export interface Group {
  lo: number
  hi: number
  level: number
  summary: number | null
  collapsed: boolean
}

/** One axis of a sheet: its entries in order, where summaries are (`after` the detail: below, right) and its last line. */
class Outline {
  readonly ls: Line<OutlineEntry>[]
  readonly last: number
  constructor(
    readonly map: Map | undefined,
    readonly axis: Axis,
    readonly after: boolean,
  ) {
    this.ls = lines(map, axis)
    this.last = LIMIT[axis] - 1
  }
  level(i: number): number {
    return i < 0 || i > this.last ? 0 : (entryAt(this.ls, i)?.outlineLevel ?? 0)
  }
  hidden(i: number): boolean {
    const e = entryAt(this.ls, i)
    return !!e && (e.hidden === true || e.height === 0 || e.width === 0)
  }
  mark(i: number): boolean {
    return entryAt(this.ls, i)?.collapsed === true
  }
  /** The end of the run from `i` towards `dir` of lines at `min` or deeper (min ≥ 1): `i - dir` when `i` is not such a line. */
  end(i: number, min: number, dir: 1 | -1): number {
    let p = i
    for (;;) {
      if (p < 0 || p > this.last) return p - dir
      const k = this.index(p)
      const e = k < 0 ? undefined : this.ls[k]
      if (!e || p > e[1] || (e[2].outlineLevel ?? 0) < min) return p - dir
      p = (dir > 0 ? e[1] : e[0]) + dir
    }
  }
  /** The index of the last entry starting at or before `i`, -1 for none. */
  index(i: number): number {
    let [a, b] = [0, this.ls.length - 1]
    if (b < 0 || this.ls[0][0] > i) return -1
    while (a < b) {
      const m = (a + b + 1) >> 1
      if (this.ls[m][0] <= i) a = m
      else b = m - 1
    }
    return a
  }
  /** Whether every line of `lo`..`hi` is hidden. */
  allHidden(lo: number, hi: number): boolean {
    let p = lo
    for (let k = Math.max(0, this.index(lo)); p <= hi; k++) {
      const e = this.ls[k]
      if (!e || e[0] > p || e[1] < p) return false
      if (!(e[2].hidden === true || e[2].height === 0 || e[2].width === 0)) return false
      p = e[1] + 1
    }
    return true
  }
  /** Whether line `s` is a summary line: the line next to it on the detail's side is deeper. */
  isSummary(s: number): boolean {
    return s >= 0 && s <= this.last && this.level(this.after ? s - 1 : s + 1) > this.level(s)
  }
  /** The detail of summary line `s`: the run of lines deeper than it next to it. */
  detail(s: number): [number, number] {
    const min = this.level(s) + 1
    return this.after ? [this.end(s - 1, min, -1), s - 1] : [s + 1, this.end(s + 1, min, 1)]
  }
  /** The group of `level` holding line `i` (at that level or deeper). */
  run(i: number, level: number): [number, number] {
    return [this.end(i, level, -1), this.end(i, level, 1)]
  }
  /** The summary lines, each once, in order: next to the entries deeper than the line past them. */
  summaries(): number[] {
    const out = new Set<number>()
    for (const [lo, hi, e] of this.ls) {
      if (!e.outlineLevel) continue
      const s = this.after ? hi + 1 : lo - 1
      if (this.isSummary(s)) out.add(s)
    }
    return [...out].sort((a, b) => a - b)
  }
  /** The run of grouped lines at the sheet's edge past which a summary would be, with the line holding its mark; null when none touches it. */
  edge(): { lo: number; hi: number; mark: number } | null {
    const m = this.after ? this.last : 0
    if (!this.level(m)) return null
    const [lo, hi] = this.after ? [this.end(m, 1, -1), m] : [m, this.end(m, 1, 1)]
    return this.isSummary(m) ? null : { lo, hi, mark: m }
  }
  /** The marks Excel keeps (see the module's comment), by line: true or false for each summary line and the edge's mark line. */
  marks(): globalThis.Map<number, boolean> {
    const out = new globalThis.Map<number, boolean>()
    for (const s of this.summaries()) {
      const [lo, hi] = this.detail(s)
      if (!this.allHidden(lo, hi)) out.set(s, false)
      else if (this.level(s) > 0 && this.allHidden(...this.run(s, this.level(s)))) out.set(s, this.mark(s))
      else out.set(s, true)
    }
    const e = this.edge()
    if (e) out.set(e.mark, this.allHidden(e.lo, e.hi))
    return out
  }
}

/** The map with `patch` over lines `lo`..`hi`. */
const patched = <T extends OutlineEntry>(map: Map<T> | undefined, axis: Axis, lo: number, hi: number, patch: LinePatch<OutlineEntry>): Map<T> => patchLines(map, axis, lo, hi, patch as LinePatch<T>)

/** The map with the marks Excel keeps: each summary line marked as `Outline.marks` says, other lines as they were. */
export function normalized<T extends OutlineEntry>(map: Map<T> | undefined, axis: Axis, after: boolean): Map<T> | undefined {
  const o = new Outline(map, axis, after)
  let out = map
  for (const [i, mark] of o.marks()) if (mark !== o.mark(i)) out = patched(out, axis, i, i, { collapsed: mark ? true : null })
  return out
}

/**
 * The groups of an axis as Excel draws them: for each level, the longest runs of lines at that level or deeper, by their
 * first line, the outer before the inner where two begin together. A group is collapsed when Excel keeps the mark on its
 * summary line (`Outline.marks`), or for a group at the sheet's edge, when the run there is all hidden.
 */
export function groupsOf(map: Map | undefined, axis: Axis, after: boolean): Group[] {
  const o = new Outline(map, axis, after)
  const marks = o.marks()
  const out: Group[] = []
  for (let level = 1; level <= MAX_OUTLINE; level++) {
    let open: [number, number] | null = null
    const close = () => {
      if (!open) return
      const [lo, hi] = open
      const s = after ? hi + 1 : lo - 1
      const inSheet = s >= 0 && s <= o.last
      const e = inSheet ? null : o.edge()
      out.push({ lo, hi, level, summary: inSheet ? s : null, collapsed: inSheet ? marks.get(s) === true : !!e && o.allHidden(e.lo, e.hi) })
    }
    for (const [lo, hi, e] of o.ls) {
      if ((e.outlineLevel ?? 0) < level) continue
      if (open && lo === open[1] + 1) open[1] = hi
      else {
        close()
        open = [lo, hi]
      }
    }
    close()
  }
  return out.sort((a, b) => a.lo - b.lo || a.level - b.level)
}

/** Why Excel refuses an outline command (`StructureRefusal`). */
export type OutlineRefusal = 'outline_deep' | 'not_grouped' | 'detail'

/** An outline command on one axis: Group or Ungroup of lines, Show or Hide Detail at a line, a level button. */
export type OutlineCommand = { kind: 'group' | 'ungroup'; lo: number; hi: number } | { kind: 'show' | 'hide'; at: number } | { kind: 'levels'; level: number }

/** The pieces of lines `lo`..`hi`: each with the level of its lines. */
function pieces(o: Outline, lo: number, hi: number): [number, number, number][] {
  const out: [number, number, number][] = []
  let p = lo
  for (let k = Math.max(0, o.index(lo)); p <= hi; k++) {
    const e = o.ls[k]
    if (!e || e[0] > hi) {
      out.push([p, hi, 0])
      break
    }
    if (e[1] < p) continue
    if (e[0] > p) out.push([p, e[0] - 1, 0])
    out.push([Math.max(p, e[0]), Math.min(hi, e[1]), e[2].outlineLevel ?? 0])
    p = e[1] + 1
  }
  return out
}

/** Lines shown: not hidden, and a size of 0 dropped, as Excel gives a line it shows its default size. */
function shown<T extends OutlineEntry>(map: Map<T> | undefined, axis: Axis, lo: number, hi: number): Map<T> {
  let out = patched(map, axis, lo, hi, { hidden: null })
  for (const [a, b, e] of lines(out, axis)) if (b >= lo && a <= hi && (e.height === 0 || e.width === 0)) out = patched(out, axis, Math.max(a, lo), Math.min(b, hi), { height: null, width: null })
  return out
}

/**
 * The axis after Excel's command, its marks as Excel keeps them, or why Excel refuses it. Group puts each line one level
 * deeper, but those at the deepest already, and is refused when every line is; Ungroup takes a level off the lines that
 * have one, refused when none has. Show and Hide Detail act at a summary line on its detail, at another line on the group
 * of its level that holds it, through that group's summary line: refused at a line in no group, at a group with no summary
 * line, and when the detail is already shown (no mark) or hidden (marked). Showing the detail shows its lines but those of
 * the collapsed groups inside it. A level button k hides the lines of level k and deeper and shows the other grouped lines.
 */
export function outlineCommand<T extends OutlineEntry>(map: Map<T> | undefined, axis: Axis, after: boolean, c: OutlineCommand): Map<T> | undefined | OutlineRefusal {
  const start = normalized(map, axis, after)
  const o = new Outline(start, axis, after)
  let out = start
  if ('lo' in c) {
    const ps = pieces(o, c.lo, c.hi)
    if (c.kind === 'group' ? ps.every(([, , l]) => l >= MAX_OUTLINE) : ps.every(([, , l]) => l === 0)) return c.kind === 'group' ? 'outline_deep' : 'not_grouped'
    for (const [a, b, l] of ps) {
      const next = c.kind === 'group' ? Math.min(l + 1, MAX_OUTLINE) : l - 1
      if (next !== l && next >= 0) out = patched(out, axis, a, b, { outlineLevel: next || null })
    }
  } else if (c.kind === 'levels') {
    // columns grouped up to the sheet's edge, past which their summary column would be, stay as they are (the set rules-held)
    const edge = axis === 'c' ? o.edge() : null
    for (const [lo, hi, e] of o.ls) {
      const l = e.outlineLevel ?? 0
      for (const [a, b] of edge ? [[lo, Math.min(hi, edge.lo - 1)], [Math.max(lo, edge.hi + 1), hi]] : [[lo, hi]])
        if (l && a <= b) out = l >= c.level ? patched(out, axis, a, b, { hidden: true }) : shown(out, axis, a, b)
    }
  } else {
    const at = c.at
    const level = o.level(at)
    let s = at
    if (!o.isSummary(at)) {
      if (!level) return 'not_grouped'
      const [lo, hi] = o.run(at, level)
      s = after ? hi + 1 : lo - 1
      if (s < 0 || s > o.last) return 'not_grouped'
    }
    const marked = o.mark(s)
    if (marked === (c.kind === 'hide')) return 'detail'
    const [lo, hi] = o.detail(s)
    if (c.kind === 'hide') out = patched(out, axis, lo, hi, { hidden: true })
    else {
      // the lines of the collapsed groups inside stay hidden: their summary lines keep their marks
      const holes = o.summaries().filter((x) => x !== s && x >= lo && x <= hi && o.mark(x)).map((x) => o.detail(x))
      let p = lo
      for (const [a, b] of holes.sort((x, y) => x[0] - y[0])) {
        if (a > p) out = shown(out, axis, p, a - 1)
        p = Math.max(p, b + 1)
      }
      if (p <= hi) out = shown(out, axis, p, hi)
      out = patched(out, axis, s, s, { collapsed: null })
    }
  }
  return normalized(out, axis, after)
}
