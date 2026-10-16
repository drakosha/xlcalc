/**
 * The texts a session's searches read (`findCells`), kept with its state: for each sheet searched or read
 * (`readCells`), the cells a search reads (its cells and those its arrays spill into) by their places, row by row, and
 * each kind's text of them, read when a search first needs it. A call that changes the book has the cells its answer
 * lists read again when a search next needs them, and the cells written since (the sheet's log of writes,
 * `cell-writes.ts`) put in their places (a place whose cell is gone stays, with no text); a change of a sheet's layout drops the sheet's texts, a change of the styles all of them,
 * a move of cells the places too, and the texts conditional formats show are read again after any change of a value,
 * which they may read. A state built anew (a structure operation, new names, a call that stopped) starts with none.
 */
import { writesMark, writesSince } from '../cell-writes.js'
import { keyParts, type EngineState } from '../engine.js'
import { paced, type Steps } from '../exec/budget.js'
import { formatAddress, MAX_COLS, MAX_ROWS, readAddress } from '../model.js'
import { lowerBound } from '../sorted.js'
import { parseRange } from '../structure/area.js'
import { sheetKey } from '../syntax.js'
import type { Outcome } from './answer.js'

/** What a search reads (`FindCellsOptions.in`), in the cells' own case or in any case (`folded`). */
export type TextKind = `${'text' | 'edit' | 'formula' | 'value'}${'' | ' folded'}`

/** The texts of a sheet's cells that searches read. */
export interface SheetTexts {
  /** The places `r * MAX_COLS + c` of the cells, ascending. */
  at: Float64Array
  /** Each kind's texts by the cells' indices in `at`; undefined where none has been read since the cell changed. */
  texts: Map<TextKind, (string | undefined)[]>
  /** The cells by columns: the indices of `at` in the order of their places `c * MAX_ROWS + r`, and those places; made when a search by columns needs them. */
  columns?: { ix: Uint32Array; at: Float64Array }
  /** The places of the cells changed since the texts were read. */
  changed: Set<number>
  /** The sheet the places are of, and how far its log of writes was read (`cell-writes.ts`). */
  sheet: object
  mark: [number, number]
}

const kept = new WeakMap<EngineState, Map<number, SheetTexts>>()

const same = (v: number) => v

/** The cells of sheet `s` a search reads: its cells, and those its arrays spill into. The steps yield between cells. */
function* sheetCells(state: EngineState, s: number): Steps<SheetTexts> {
  const sheet = state.input.sheets[s]
  const mark = writesMark(sheet)
  const { cells } = sheet
  const out: number[] = []
  const p = { r: 0, c: 0 }
  for (const addr in cells) {
    if (paced(1)) yield
    readAddress(addr, p)
    out.push(p.r * MAX_COLS + p.c)
  }
  for (const x of state.spills.on.values()) {
    if (x.s !== s) continue
    for (let r = x.r1; r <= x.r2; r++)
      for (let c = x.c1; c <= x.c2; c++) {
        if (paced(1)) yield
        if (!(formatAddress(r, c) in cells)) out.push(r * MAX_COLS + c)
      }
  }
  return { at: Float64Array.from(out).sort(), texts: new Map(), changed: new Set(), sheet, mark }
}

/** The changed cells' texts left to read again, and the cells new to the sheet put in their places. */
function refresh(x: SheetTexts) {
  if (!x.changed.size) return
  const fresh: number[] = []
  for (const m of x.changed) {
    const i = lowerBound(x.at, m, same)
    if (x.at[i] === m) for (const t of x.texts.values()) t[i] = undefined
    else fresh.push(m)
  }
  x.changed.clear()
  if (!fresh.length) return
  fresh.sort((a, b) => a - b)
  const old = x.at
  const at = new Float64Array(old.length + fresh.length)
  // each old cell's new index
  const moved = new Uint32Array(old.length)
  for (let i = 0, j = 0, k = 0; k < at.length; k++) {
    if (j < fresh.length && (i === old.length || fresh[j] < old[i])) at[k] = fresh[j++]
    else {
      moved[i] = k
      at[k] = old[i++]
    }
  }
  for (const [kind, t] of x.texts) {
    const u = new Array<string | undefined>(at.length)
    for (let i = 0; i < old.length; i++) u[moved[i]] = t[i]
    x.texts.set(kind, u)
  }
  x.at = at
  x.columns = undefined
}

/**
 * The texts of sheet `s` kept with the state, its cells read first when it has none (or its sheet is another object, or
 * its log of writes started again); the cells written since put in their places, and the changed cells' texts left to
 * read again.
 */
export function* textsOf(state: EngineState, s: number): Steps<SheetTexts> {
  let sheets = kept.get(state)
  if (!sheets) kept.set(state, (sheets = new Map()))
  let x = sheets.get(s)
  const written = x && x.sheet === state.input.sheets[s] ? writesSince(x.sheet, x.mark) : null
  if (x && written) {
    for (const m of written) x.changed.add(m)
    x.mark = writesMark(x.sheet)
    refresh(x)
  } else sheets.set(s, (x = yield* sheetCells(state, s)))
  return x
}

/** The kinds of texts a sheet keeps at most, the most recently searched: two kinds in both cases, Look in Formulas and Values. */
const KINDS = 4

/** The texts of kind `kind` of a sheet's cells, by their indices; the kind searched longest ago dropped past `KINDS`. */
export function kindOf(x: SheetTexts, kind: TextKind): (string | undefined)[] {
  const t = x.texts.get(kind) ?? new Array<string | undefined>(x.at.length)
  x.texts.delete(kind)
  x.texts.set(kind, t)
  for (const k of x.texts.keys()) if (x.texts.size > KINDS) x.texts.delete(k)
  return t
}

/** A sheet's cells by columns. */
export function byColumns(x: SheetTexts): NonNullable<SheetTexts['columns']> {
  if (!x.columns) {
    const keys = Float64Array.from(x.at, (m) => (m % MAX_COLS) * MAX_ROWS + Math.floor(m / MAX_COLS))
    const ix = Uint32Array.from(keys.keys()).sort((a, b) => keys[a] - keys[b])
    x.columns = { ix, at: Float64Array.from(ix, (i) => keys[i]) }
  }
  return x.columns
}

/** A call that changed the book, by its answer (`Outcome`): what the texts kept with its state no longer hold. */
export function textsChanged(state: EngineState, o: Outcome): void {
  const sheets = kept.get(state)
  if (!sheets?.size) return
  if (o.moved.length) return sheets.clear()
  // a sheet's layout and the styles change what its cells show, not where its cells are
  for (const x of o.styles ? sheets.values() : []) x.texts.clear()
  for (const name of o.layout) sheets.get(state.sheetIdx.get(sheetKey(name))!)?.texts.clear()
  // what conditional formats show reads values anywhere
  const shown = new Set(o.rules.map((x) => state.sheetIdx.get(sheetKey(x.sheet))))
  for (const [i, x] of sheets)
    if (shown.has(i) || (o.valued && state.input.sheets[i].conditionalFormats?.length)) {
      x.texts.delete('text')
      x.texts.delete('text folded')
    }
  for (const k of o.keys) {
    const [i, r, c] = keyParts(k)
    const x = sheets.get(i)
    if (!x) continue
    x.changed.add(r * MAX_COLS + c)
    // each cell of the area of an array of the old kind opens with its formula
    const area = state.input.sheets[i].cells[formatAddress(r, c)]?.array
    const a = area === undefined ? null : parseRange(area)
    if (a) for (let ar = a.r1; ar <= a.r2; ar++) for (let ac = a.c1; ac <= a.c2; ac++) x.changed.add(ar * MAX_COLS + ac)
  }
  // a sheet most of whose cells changed is read again whole
  for (const [i, x] of sheets) if (x.changed.size > x.at.length / 2) sheets.delete(i)
}
