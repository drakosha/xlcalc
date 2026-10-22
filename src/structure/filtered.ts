/**
 * Rows on a filtered sheet, where Excel acts on the visible rows (the struct set `hidden`): deleting rows deletes
 * the visible ones among them and inserting rows inserts as many rows as are visible among them, at the first; a
 * copy leaves out the hidden rows of its source (`clip.ts`) and pastes into each run of visible rows of its target
 * as into a target alone. A selection of hidden rows only is taken whole. Every hidden row of a filtered sheet
 * counts, whatever hid it.
 */
import type { BookInput, CellInput } from '../contract.js'
import { say } from '../locale/messages.js'
import { cornersText, hasCell, type Rect } from './area.js'
import { clipped, CLIP } from './clip.js'
import { copyRange } from './copy.js'
import { placer, type Applied, type Geometry } from './geometry.js'
import { placedCells, shift, shiftTransform } from './shift.js'
import { arraysOf, gridOf } from './grid.js'
import { rewriteCells, sheetIndex, sheetNames, sheetOf, type Site, type Transform } from './rewrite.js'
import { deleteSpan } from './transform.js'
import { madeFrom } from '../looks/sheet.js'
import type { ValuesIn } from './values.js'
import { visibleRows, visibleRuns } from './visible.js'

/** Applies the steps in turn, each to the book the last one gave; the geometry is theirs one after another. */
function chain(book: BookInput, steps: ((b: BookInput) => Applied)[]): Applied {
  let out = book
  const geometry: Geometry = []
  for (const step of steps) {
    const r = step(out)
    if (!r.ok) return r
    out = r.book
    geometry.push(...r.geometry)
  }
  return { ok: true, book: out, geometry }
}

/** `n` rows inserted before `at` (zero-based) on `sheet`: on a filtered sheet as many as are visible among them. */
export function insertRows(book: BookInput, sheet: string, at: number, n: number): Applied {
  const visible = visibleRows(book, sheet, at, at + n - 1)
  return shift(book, sheet, 'r', at, visible?.length || n)
}

/** Rows `at` … `at + n - 1` deleted from `sheet`: on a filtered sheet the visible ones, each run of them from the bottom up. */
export function deleteRows(book: BookInput, sheet: string, at: number, n: number): Applied {
  const visible = visibleRuns(book, sheet, at, at + n - 1)
  if (!visible?.length) return shift(book, sheet, 'r', at, -n)
  const runs = visible.reverse()
  const steps = (b: BookInput) => chain(b, runs.map(([lo, hi]) => (x: BookInput) => shift(x, sheet, 'r', lo, lo - hi - 1)))
  // a table's rows deleted run by run change its cells (`keepFirstRow`, `tableShift`): the cells go through each step
  return runs.length > 1 && !sheetOf(book, sheet).tables?.length ? movedOnce(book, sheet, runs, steps) : steps(book)
}

/** The transforms one after another, as each step's formulas are rewritten by its own: a reference lost stays lost. */
function composed(ts: Transform[]): Transform {
  return {
    ref(site) {
      let out: Site | undefined
      for (const t of ts) {
        const o = t.ref(out ?? site)
        if (o === 'lost' || o === 'gone') return o
        out = o ?? out
      }
      return out
    },
  }
}

/**
 * The runs of rows deleted (`steps` of `runs`, descending) with the cells moved once: the steps run over the book without
 * cells but the sheet's array formulas, which their refusals read, and every sheet's cells take the steps' transforms
 * one after another, where a step each rewrote and placed them all (O(runs × book)).
 */
function movedOnce(book: BookInput, sheet: string, runs: [number, number][], steps: (b: BookInput) => Applied): Applied {
  const ti = sheetIndex(book, sheet)
  const bare = book.sheets.map((s, i) => ({ ...s, cells: i === ti ? Object.fromEntries(arraysOf(s).map(([addr]) => [addr, s.cells[addr]])) : {} }))
  const r = steps({ ...book, sheets: bare })
  if (!r.ok) return r
  const name = book.sheets[ti].name
  const t = composed(runs.map(([lo, hi]) => shiftTransform(name, 'r', lo, hi - lo + 1, deleteSpan)))
  const names = sheetNames(book)
  const sheets = r.book.sheets.map((s, i) => {
    const old = book.sheets[i]
    const cells = i === ti ? placedCells(old, ti, r.geometry, t, names) : rewriteCells(old.cells, old.name, old.name, t, names)
    if (cells === old.cells) return s === bare[i] ? old : { ...s, cells }
    madeFrom([old.cells], cells)
    return { ...s, cells }
  })
  return { ...r, book: { ...r.book, sheets } }
}

/**
 * The runs of visible rows a paste into `target` of sheet `to` fills one by one, or null where it fills the target whole:
 * one cell, or no hidden row of a filtered sheet.
 */
export function pasteRuns(book: BookInput, to: string, target: Rect): [number, number][] | null {
  const one = target.r1 === target.r2 && target.c1 === target.c2
  const visible = one ? null : visibleRuns(book, to, target.r1, target.r2)
  return visible?.length ? visible : null
}

/**
 * Copy and paste over filtered sheets. A source over hidden rows of a filtered sheet is copied as its visible rows,
 * packed and as values (`values` gives them). A target of several cells over hidden rows of a filtered sheet takes
 * the copy in each run of its visible rows, as a target alone; a copy that left rows out cannot be pasted over
 * several runs (Excel refuses a selection of several areas pasted onto several areas).
 */
export function copyRows(book: BookInput, from: string, src: Rect, to: string, target: Rect, withMeta: boolean, values: ValuesIn): Applied {
  const clip = clipped(book, from, src, values)
  const [source, at, rect] = clip ? [clip.book, CLIP, clip.src] : [book, from, src]
  const parts = pasteRuns(book, to, target)
  if (!parts) return copyRange(book, at, rect, to, target, withMeta, source)
  if (clip && parts.length > 1)
    return { ok: false, code: 'filtered', message: say(book.locale, { key: 'filteredPaste', sheet: to, area: cornersText(target) }) }
  const steps = (b: BookInput) => chain(b, parts.map(([lo, hi]) => (x: BookInput) => copyRange(x, at, rect, to, { ...target, r1: lo, r2: hi }, withMeta, source)))
  return parts.length > 1 ? pastedOnce(book, to, target, steps) : steps(book)
}

/**
 * The pastes into the runs of `target` (`steps`) over the sheet `to` with only its cells of the target and its array
 * formulas, which a paste reads, where each paste walked all its cells (O(runs × sheet)): its other cells come back
 * after, in its order before the pasted ones, as the pastes leave them.
 */
function pastedOnce(book: BookInput, to: string, target: Rect, steps: (b: BookInput) => Applied): Applied {
  const ti = sheetIndex(book, to)
  const sheet = book.sheets[ti]
  const g = gridOf(sheet)
  const arrays = new Set(arraysOf(sheet).map(([addr]) => addr))
  const read: Record<string, CellInput> = {}
  for (let i = 0; i < g.addr.length; i++) if (hasCell(target, g.row[i], g.col[i]) || arrays.has(g.addr[i])) read[g.addr[i]] = g.cell[i]
  const r = steps({ ...book, sheets: book.sheets.map((s, i) => (i === ti ? { ...s, cells: read } : s)) })
  if (!r.ok) return r
  const [pasted, place, p] = [r.book.sheets[ti], placer(r.geometry), { s: ti, r: 0, c: 0 }]
  const cells: Record<string, CellInput> = {}
  for (let i = 0; i < g.addr.length; i++) {
    p.r = g.row[i]
    p.c = g.col[i]
    // the pastes clear only cells of the target
    if (!hasCell(target, p.r, p.c) || place(p)) cells[g.addr[i]] = g.cell[i]
  }
  for (const a in pasted.cells) if (!(a in cells)) cells[a] = pasted.cells[a]
  return { ...r, book: { ...r.book, sheets: r.book.sheets.map((s, i) => (i === ti ? { ...pasted, cells } : s)) } }
}
