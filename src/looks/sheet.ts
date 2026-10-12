/**
 * The looks of a sheet of a book (`SheetInput.looks`) as a layer (`layer.ts`): read once per list, as a list never
 * changes (a change gives the sheet a new one); the look of a cell without an entry; a sheet as the engine takes it;
 * the writes of a session's cells, and a sheet's looks put back into entries for an operation that reads entries alone.
 */
import type { CellInput, LookRange, SheetInput } from '../contract.js'
import type { Look } from '../layout/lines.js'
import { formatAddress, MAX_COLS, readAddress } from '../model.js'
import type { Rect } from '../structure/area.js'
import { cellEdits, edited, emptyLayer, LookTable, NONE, type Layer, type RowsEdit } from './layer.js'
import { readRanges, writeRanges } from './ranges.js'

/** The layer of each list read; the lists written in the one form (`writeRanges`) are its keys from the start. */
const LAYERS = new WeakMap<readonly LookRange[], Layer>()
const WRITTEN = new WeakSet<readonly LookRange[]>()

/**
 * The cells taken (`looksTaken`): no entry of a look alone, no look under an entry with content. The writes of a session
 * (`putInput`) keep them so, and an operation that gives a book again gives new cells, which are taken anew.
 */
const TAKEN = new WeakSet<object>()

/** Looks of a session's cells written and not yet in the sheet's list, by row · MAX_COLS + column (`putInput`). */
const PENDING = new WeakMap<object, Map<number, Look | undefined>>()

/** The layer of a list, read once. */
function layerOfList(list: readonly LookRange[]): Layer {
  let layer = LAYERS.get(list)
  if (!layer) LAYERS.set(list, (layer = readRanges(list)))
  return layer
}

/** A layer's list in the one form, known as written in it (the layer read already). */
export function listOf(layer: Layer): LookRange[] {
  const list = writeRanges(layer)
  LAYERS.set(list, layer)
  WRITTEN.add(list)
  return list
}

/** The layer of a sheet's looks (its cells' writes in it), undefined without looks. */
export function layerOf(sheet: Pick<SheetInput, 'looks'>): Layer | undefined {
  settle(sheet)
  return sheet.looks?.length ? layerOfList(sheet.looks) : undefined
}

/** The look of a cell from the sheet's looks: what a cell without an entry shows as its own. */
export const lookIn = (sheet: Pick<SheetInput, 'looks'>, r: number, c: number): Look | undefined => layerOf(sheet)?.lookAt(r, c)

/** A cell's entry, else its look from the sheet's looks as an entry of a look alone (a frozen object), else undefined. */
export const inputIn = (sheet: Pick<SheetInput, 'cells' | 'looks'>, addr: string, r: number, c: number): CellInput | undefined =>
  sheet.cells[addr] ?? (sheet.looks || PENDING.has(sheet) ? lookIn(sheet, r, c) : undefined)

/** Whether an entry holds a look alone: a style or a number format and no other field (a field of undefined is none); an empty entry holds none. */
export function lookAlone(cell: CellInput): boolean {
  for (const k in cell) if (k !== 'style' && k !== 'numberFormat' && (cell as Record<string, unknown>)[k] !== undefined) return false
  return cell.style !== undefined || cell.numberFormat !== undefined
}

/** The sheet with `layer` as its looks: a new object, without the field for no looks. */
const withLayer = (sheet: SheetInput, cells: Record<string, CellInput>, layer: Layer): SheetInput => withLooks({ ...sheet, cells }, layer.empty ? undefined : listOf(layer))

/**
 * A sheet (checked) as the engine takes it: its entries of a look alone moved into its looks, over the rectangles it has,
 * no look under its other entries, the list in the one form. `alone`: the addresses of those entries when the caller
 * found them already. The sheet itself when it has neither, or its looks are in the one form and nothing moves.
 */
export function looksTaken(sheet: SheetInput, alone?: string[]): SheetInput {
  const cells = sheet.cells
  if (TAKEN.has(cells) && (!sheet.looks || WRITTEN.has(sheet.looks))) return sheet
  if (!alone) {
    alone = []
    for (const a in cells) if (lookAlone(cells[a])) alone.push(a)
  }
  if (!alone.length && !sheet.looks) {
    TAKEN.add(cells)
    return sheet
  }
  const base = sheet.looks?.length ? layerOfList(sheet.looks) : emptyLayer()
  const looks = new LookTable(base.looks)
  const p = { r: 0, c: 0 }
  const moved = new Set(alone)
  const edits: [number, number, number][] = []
  for (const a of alone) {
    readAddress(a, p)
    edits.push([p.r, p.c, looks.of(cells[a])])
  }
  if (!base.empty)
    for (const a in cells) {
      if (moved.has(a)) continue
      readAddress(a, p)
      if (base.idAt(p.r, p.c) !== NONE) edits.push([p.r, p.c, NONE])
    }
  const layer = edited(base, cellEdits(edits), looks)
  const kept = moved.size ? keptCells(cells, moved) : cells
  TAKEN.add(kept)
  if (layer === base && sheet.looks && WRITTEN.has(sheet.looks) && kept === cells) return sheet
  return withLayer(sheet, kept, layer)
}

/** The cells but those of `gone`. */
function keptCells(cells: Record<string, CellInput>, gone: Set<string>): Record<string, CellInput> {
  const out: Record<string, CellInput> = {}
  for (const a in cells) if (!gone.has(a)) out[a] = cells[a]
  return out
}

/**
 * Writes a cell's input on a sheet of a session's book: content to the sheet's entries (with no look under it), a look
 * alone or none to its looks, which take the writes when they are next read (`layerOf`) or settled.
 */
export function putInput(sheet: SheetInput, addr: string, r: number, c: number, input: CellInput | null): void {
  const alone = input === null || lookAlone(input)
  if (alone) delete sheet.cells[addr]
  else sheet.cells[addr] = input
  let pending = PENDING.get(sheet)
  const k = r * MAX_COLS + c
  // content over a cell: no look under it, which matters only where the layer has one
  if (!alone && !pending?.has(k) && !(sheet.looks?.length && layerOfList(sheet.looks).idAt(r, c) !== NONE)) return
  if (!pending) PENDING.set(sheet, (pending = new Map()))
  pending.set(k, alone && input ? input : undefined)
}

/** The cells' writes (`putInput`) put into the sheet's list of looks. */
export function settle(sheet: Pick<SheetInput, 'looks'>): void {
  const pending = PENDING.get(sheet)
  if (!pending) return
  PENDING.delete(sheet)
  const base = sheet.looks?.length ? layerOfList(sheet.looks) : emptyLayer()
  const looks = new LookTable(base.looks)
  const edits: [number, number, number][] = []
  for (const [k, look] of pending) edits.push([Math.floor(k / MAX_COLS), k % MAX_COLS, looks.of(look)])
  const layer = edited(base, cellEdits(edits), looks)
  if (layer === base) return
  if (layer.empty) delete sheet.looks
  else sheet.looks = listOf(layer)
}

/** Rectangles as edits of rows that clear them: each band of rows with the columns the rectangles over it cover. */
function clearing(xs: Rect[]): RowsEdit[] {
  const cuts = [...new Set(xs.flatMap((x) => [x.r1, x.r2 + 1]))].sort((a, b) => a - b)
  const out: RowsEdit[] = []
  for (let k = 0; k + 1 < cuts.length; k++) {
    const [lo, hi] = [cuts[k], cuts[k + 1] - 1]
    const cols = xs.filter((x) => x.r1 <= lo && x.r2 >= hi).map((x): [number, number] => [x.c1, x.c2]).sort((a, b) => a[0] - b[0])
    const top: number[] = []
    for (const [c1, c2] of cols) {
      const n = top.length
      if (n && top[n - 2] >= c1 - 1) top[n - 2] = Math.max(top[n - 2], c2)
      else top.push(c1, c2, NONE)
    }
    if (top.length) out.push({ lo, hi, top })
  }
  return out
}

/**
 * The sheet with the looks of its cells without an entry in rectangles `xs` put back as entries of a look alone, and its
 * looks without them: what an operation that reads and moves entries alone takes (a structure operation). The book the
 * operation gives is taken again (`looksTaken`), so its looks are in the one form; the sheet itself when nothing moves.
 */
export function expanded(sheet: SheetInput, xs: Rect[]): SheetInput {
  const layer = layerOf(sheet)
  if (!layer || !xs.length) return sheet
  const cells = { ...sheet.cells }
  let n = 0
  for (const x of xs)
    for (const piece of layer.pieces(x)) {
      const look = layer.looks.list[piece.look]
      for (let r = piece.r1; r <= piece.r2; r++)
        for (let c = piece.c1; c <= piece.c2; c++) {
          const a = formatAddress(r, c)
          if (cells[a]) continue
          cells[a] = { ...look }
          n++
        }
    }
  if (!n) return sheet
  return withLayer(sheet, cells, edited(layer, clearing(xs)))
}

/**
 * The lines of a layer as the input hash writes them: each pattern once, when a block first uses it (`W`, its runs with
 * their looks by `lookText`), and each block of rows with its pattern's number in the order they came (`B`).
 */
export function* layerLines(layer: Layer, lookText: (look: Look) => string): Generator<string> {
  const seen = new Map<number, number>()
  for (let b = 0; b < layer.blocks; b++) {
    const k = layer.at[b]
    let n = seen.get(k)
    if (n === undefined) {
      seen.set(k, (n = seen.size))
      const p = layer.patterns[k]
      let line = `W ${n}`
      for (let i = 0; i < p.length; i += 3) line += ` ${p[i]}-${p[i + 1]} ${lookText(layer.looks.list[p[i + 2]])}`
      yield `${line}\n`
    }
    yield `B ${layer.r1[b]}-${layer.r2[b]} ${n}\n`
  }
}

/** Cells' inputs written on the sheets of a book (`putInput`), each sheet then settled: what a call that stopped puts back. */
export function putInputs(sheets: SheetInput[], index: (name: string) => number, edits: { sheet: string; cell: string; input: CellInput | null }[]): void {
  const p = { r: 0, c: 0 }
  const touched = new Set<SheetInput>()
  for (const { sheet, cell, input } of edits) {
    const s = sheets[index(sheet)]
    readAddress(cell, p)
    putInput(s, cell, p.r, p.c, input)
    touched.add(s)
  }
  for (const s of touched) settle(s)
}

/** The sheet's looks without those of cells `at` (`[row, column]`): a new list, or the sheet's own when none of them has a look. */
export function looksWithout(sheet: Pick<SheetInput, 'looks'>, at: [number, number][]): LookRange[] | undefined {
  const layer = layerOf(sheet)
  if (!layer) return sheet.looks
  const next = edited(layer, cellEdits(at.map(([r, c]): [number, number, number] => [r, c, NONE])))
  return next === layer ? sheet.looks : next.empty ? undefined : listOf(next)
}

/** A sheet with the list of looks given: a new object, without the field for none. */
export function withLooks(sheet: SheetInput, looks: LookRange[] | undefined): SheetInput {
  const out = { ...sheet, looks }
  if (!looks) delete out.looks
  return out
}
