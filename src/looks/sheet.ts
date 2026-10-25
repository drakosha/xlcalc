/**
 * The looks of a sheet of a book (`SheetInput.looks`) as a layer (`layer.ts`): read once per list, as a list never
 * changes (a change gives the sheet a new one); the look of a cell without an entry; a sheet as the engine takes it;
 * the writes of a session's cells, and a sheet's looks put back into entries for an operation that reads entries alone.
 */
import type { CellInput, LookRange, SheetInput } from '../contract.js'
import type { Look } from '../layout/lines.js'
import { wrote } from '../cell-writes.js'
import { formatAddress, MAX_COLS, readAddress } from '../model.js'
import type { Rect } from '../structure/area.js'
import { cellEdits, edited, emptyLayer, LookTable, NONE, type Layer } from './layer.js'
import { clearing } from './move.js'
import { readRanges, writeRanges } from './ranges.js'

/** The layer of each list read; the lists written in the one form (`writeRanges`) are its keys from the start. */
const LAYERS = new WeakMap<readonly LookRange[], Layer>()
const WRITTEN = new WeakSet<readonly LookRange[]>()

/**
 * The cells taken (`looksTaken`): no entry of a look alone, no look under an entry with content. The writes of a session
 * (`putInput`) keep them so, and an operation that gives a book again gives new cells, which are taken anew.
 */
const TAKEN = new WeakSet<object>()
/** Cells a step made from cells taken (`TAKEN`), by the addresses of their entries of a look alone: `looksTaken` reads no other entry of them. */
const MADE = new WeakMap<object, string[]>()

/** Looks of a session's cells written and not yet in the sheet's list, by row · MAX_COLS + column (`putInput`). */
const PENDING = new WeakMap<object, Map<number, Look | undefined>>()

/** The layer a sheet holds whose list is not made yet (`setLayer`): an edit changes the layer, the list is made when read. */
const CURRENT = new WeakMap<object, Layer>()

/** The layer of a list, read once. */
function layerOfList(list: readonly LookRange[]): Layer {
  let layer = LAYERS.get(list)
  if (!layer) LAYERS.set(list, (layer = readRanges(list)))
  return layer
}

/** The sheet's list of looks at last, read or set: a field of its own, the layer no more held beside it. */
function made(sheet: object, list: LookRange[] | undefined): LookRange[] | undefined {
  CURRENT.delete(sheet)
  if (list) Object.defineProperty(sheet, 'looks', { value: list, writable: true, configurable: true, enumerable: true })
  else delete (sheet as SheetInput).looks
  return list
}

/**
 * Gives a sheet `layer` as its looks: the list in the one form is made when `looks` is first read (a book given back,
 * a copy, a host), so a session's edit of a cell changes the layer alone; no field for an empty layer.
 */
export function setLayer(sheet: SheetInput, layer: Layer): void {
  if (layer.empty) {
    made(sheet, undefined)
    return
  }
  CURRENT.set(sheet, layer)
  Object.defineProperty(sheet, 'looks', {
    configurable: true,
    enumerable: true,
    get(this: SheetInput) {
      return made(this, listOf(layer))
    },
    set(this: SheetInput, list: LookRange[] | undefined) {
      made(this, list)
    },
  })
}

/** The layer of a sheet's looks as it is, without its cells' writes not yet settled; undefined without looks. */
function heldLayer(sheet: Pick<SheetInput, 'looks'>): Layer | undefined {
  return CURRENT.get(sheet) ?? (sheet.looks?.length ? layerOfList(sheet.looks) : undefined)
}

/** Whether a sheet has looks, or writes of looks to settle: without making its list. */
export const hasLooks = (sheet: Pick<SheetInput, 'looks'>): boolean => CURRENT.has(sheet) || PENDING.has(sheet) || !!sheet.looks?.length

/** A copy of a sheet without its looks (their list is not made for it). */
function withoutLooks(sheet: SheetInput): SheetInput {
  const out: Record<string, unknown> = {}
  for (const k of Object.keys(sheet)) if (k !== 'looks') out[k] = (sheet as unknown as Record<string, unknown>)[k]
  return out as unknown as SheetInput
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
  return heldLayer(sheet)
}

/** The look of a cell from the sheet's looks: what a cell without an entry shows as its own. */
export const lookIn = (sheet: Pick<SheetInput, 'looks'>, r: number, c: number): Look | undefined => layerOf(sheet)?.lookAt(r, c)

/** A cell's entry, else its look from the sheet's looks (which is undefined without them) as an entry of a look alone (a frozen object), else undefined. */
export const inputIn = (sheet: Pick<SheetInput, 'cells' | 'looks'>, addr: string, r: number, c: number): CellInput | undefined =>
  sheet.cells[addr] ?? lookIn(sheet, r, c)

/**
 * Marks `cells` as made of the entries of cells taken (`from`, each moved or its formula rewritten, which makes no entry
 * a look alone) and the entries `added`: its entries of a look alone are those of `added` (`looksTaken`).
 */
export function madeFrom(from: readonly object[], cells: object, added: Record<string, CellInput | null> = {}): void {
  if (from.every((x) => TAKEN.has(x))) MADE.set(cells, Object.keys(added).filter((a) => added[a] && lookAlone(added[a]!)))
}

/** Whether an entry (a cell, a row, a column) shows a look of its own: a style or a number format. */
export const hasLook = (e: Pick<Look, 'style' | 'numberFormat'>): boolean => e.style !== undefined || e.numberFormat !== undefined

/** Whether an entry holds a look alone: a style or a number format and no other field (a field of undefined is none); an empty entry holds none. */
export function lookAlone(cell: CellInput): boolean {
  for (const k in cell) if (k !== 'style' && k !== 'numberFormat' && (cell as Record<string, unknown>)[k] !== undefined) return false
  return hasLook(cell)
}

/** The sheet with `cells` and `layer` as its looks: a new object, without the field for no looks. */
export function withLayer(sheet: SheetInput, cells: Record<string, CellInput>, layer: Layer): SheetInput {
  const out = { ...withoutLooks(sheet), cells }
  setLayer(out, layer)
  return out
}

/**
 * A sheet (checked) as the engine takes it: its entries of a look alone moved into its looks, over the rectangles it has,
 * no look under its other entries, the list in the one form. `alone`: the addresses of those entries when the caller
 * found them already. The sheet itself when it has neither, or its looks are in the one form and nothing moves.
 */
export function looksTaken(sheet: SheetInput, alone?: string[]): SheetInput {
  const cells = sheet.cells
  // looks in the one form: a layer held, or a list written in it
  const formed = CURRENT.has(sheet) || !sheet.looks || WRITTEN.has(sheet.looks)
  if (TAKEN.has(cells) && formed) return sheet
  alone ??= MADE.get(cells)
  if (!alone) {
    alone = []
    for (const a in cells) if (lookAlone(cells[a])) alone.push(a)
  }
  if (!alone.length && !hasLooks(sheet)) {
    TAKEN.add(cells)
    return sheet
  }
  const base = layerOf(sheet) ?? emptyLayer()
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
  if (layer === base && formed && kept === cells) return sheet
  return withLayer(sheet, kept, layer)
}

/** The cells but those of `gone`. */
function keptCells(cells: Record<string, CellInput>, gone: Set<string>): Record<string, CellInput> {
  const out: Record<string, CellInput> = {}
  for (const a in cells) if (!gone.has(a)) out[a] = cells[a]
  return out
}

/**
 * Writes a cell's input on a sheet of a session's book, the one way a cell of a live book is written: content to the
 * sheet's entries (with no look under it), a look alone or none to its looks, which take the writes when they are next
 * read (`layerOf`) or settled; the cell's place noted in the sheet's log of writes (`cell-writes.ts`).
 */
export function putInput(sheet: SheetInput, addr: string, r: number, c: number, input: CellInput | null): void {
  wrote(sheet, r * MAX_COLS + c)
  const alone = input === null || lookAlone(input)
  if (alone) delete sheet.cells[addr]
  else sheet.cells[addr] = input
  let pending = PENDING.get(sheet)
  const k = r * MAX_COLS + c
  // content over a cell: no look under it, which matters only where the layer has one
  if (!alone && !pending?.has(k) && (heldLayer(sheet)?.idAt(r, c) ?? NONE) === NONE) return
  if (!pending) PENDING.set(sheet, (pending = new Map()))
  pending.set(k, alone && input ? input : undefined)
}

/** The cells' writes (`putInput`) put into the sheet's list of looks. */
export function settle(sheet: Pick<SheetInput, 'looks'>): void {
  const pending = PENDING.get(sheet)
  if (!pending) return
  PENDING.delete(sheet)
  const base = heldLayer(sheet) ?? emptyLayer()
  const looks = new LookTable(base.looks)
  const edits: [number, number, number][] = []
  for (const [k, look] of pending) edits.push([Math.floor(k / MAX_COLS), k % MAX_COLS, looks.of(look)])
  const layer = edited(base, cellEdits(edits), looks)
  if (layer !== base) setLayer(sheet as SheetInput, layer)
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

/** Sheets given back the looks they had (`layers`, by index, as `layerOf` read them), where they have others now. */
export function putLayers(sheets: SheetInput[], layers: (Layer | undefined)[]): void {
  sheets.forEach((x, i) => {
    if (layerOf(x) !== layers[i]) setLayer(x, layers[i] ?? emptyLayer())
  })
}

/** The sheet without the looks of cells `at` (`[row, column]`): a new sheet, or the sheet itself when none of them has a look. */
export function withoutLooksAt(sheet: SheetInput, at: [number, number][]): SheetInput {
  const layer = layerOf(sheet)
  const next = layer && edited(layer, cellEdits(at.map(([r, c]): [number, number, number] => [r, c, NONE])))
  return !next || next === layer ? sheet : withLayer(sheet, sheet.cells, next)
}

/** The sheet `sheet` with the looks of `live`, the same sheet as it is now (a layout made before its cells were written). */
export function withLooksOf(sheet: SheetInput, live: SheetInput): SheetInput {
  const layer = layerOf(live)
  return layerOf(sheet) === layer ? sheet : withLayer(sheet, sheet.cells, layer ?? emptyLayer())
}
