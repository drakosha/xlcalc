/**
 * Fields of a workbook the model holds for a host, through Excel for Windows (`bench/windows/run.ts fields <set>`): a
 * set names its cases, each a workbook written by `writeXlsx`, or by `writeKept` over a file `writeXlsx` wrote after
 * layout edits in a session and the sheets' views a host set on the workbook it saves; Excel opens the file, reads
 * each sheet's tab colour, gridlines and headings and the calculation option through its object model, and saves it. `test/fields-win.test.ts` checks that Excel read what the workbook holds and that the
 * reader reads it back from the file Excel saved.
 */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { BookInput, CellEdit, CellPatch, LayoutEdit, SheetView, SplitPanes } from '../../../src/contract.js'
import { withField } from '../../../src/looks/sheet.js'
import { createSession } from '../../../src/index.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { parseAddress } from '../../../src/model.js'
import { openKept, writeKept } from '../../../src/xlsx-keep.js'

const ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '../../..')

export interface FieldsCase {
  id: string
  book: BookInput
  /** Layout edits a session makes over the book read by `openKept`, and views a host sets by sheet (null removes one), written over its file by `writeKept`. */
  kept?: { layout?: LayoutEdit[]; views?: Record<string, SheetView | null>; calculation?: BookInput['calculation'] | null; cells?: CellEdit[]; patches?: CellPatch[] }
}

/** What Excel read of each sheet: its tab colour (`#RRGGBB`, null for none), whether it shows gridlines and headings. */
export type SheetFields = [name: string, tab: string | null, gridlines: boolean, headings: boolean]

/** A note as Excel read it: its sheet and cell, and whether it is shown all the time. */
export type NoteShown = [sheet: string, cell: string, shown: boolean]

/** A cell whose text is turned: its sheet and cell, and the rotation. */
export type Turned = [sheet: string, cell: string, rotation: number | 'vertical']

/** A sheet's window as Excel read it: its name, frozen or not, split or not, each pane's first row and column (from 1) in Excel's order, the active pane's number. */
export type SheetPanes = [name: string, frozen: boolean, split: boolean, panes: [row: number, column: number][], active: number]

/** A sheet's window as Excel read it, and where its object model puts the lines, in points (`SplitHorizontal`, `SplitVertical`). */
export interface PanesAnswer {
  panes: SheetPanes
  points: [x: number, y: number]
}

/** A sheet's window as both Excels' steps read it: name, frozen, split, the lines in points, each pane's first row and column, the active pane. */
export type RawPanes = [name: string, frozen: boolean, split: boolean, x: number, y: number, panes: [number, number][], active: number]

export const panesAnswer = ([name, frozen, split, x, y, panes, active]: RawPanes): PanesAnswer => ({ panes: [name, frozen, split, panes, active], points: [x, y] })

/** What Excel read: the sheets' fields, the calculation option (`automatic`, `manual`, `autoNoTable`) since the set `book`, the notes since `notes`, the turned cells since `rotation`, the windows since `split`. */
export interface FieldsAnswer {
  sheets: SheetFields[]
  calculation?: string
  notes?: NoteShown[]
  turned?: Turned[]
  panes?: PanesAnswer[]
}

export interface FieldsSet {
  buildFieldsCases(): FieldsCase[]
}

/** The answers of a set and the folder of the files Excel saved: Excel for Windows's (`win/fields`) or for Mac's (`fields`). */
export const fieldsFile = (set: string, dir = 'win/fields') => join(ROOT, 'test/excel', dir, `${set}.json`)
export const savedDir = (set: string, dir = 'win/fields') => join(ROOT, 'test/excel', dir, set)

/** A book with the views and the calculation option a host set on it. */
function hostSet(b: BookInput, kept: NonNullable<FieldsCase['kept']>): BookInput {
  const views = kept.views ?? {}
  const out = { ...b, sheets: b.sheets.map((s) => (s.name in views ? withField(s, 'view', views[s.name] ?? undefined) : s)) }
  return kept.calculation === undefined ? out : withField(out, 'calculation', kept.calculation ?? undefined)
}

/** The book a case's file holds: its book, after its layout edits and views for a kept case. */
export function caseBook(c: FieldsCase): BookInput {
  if (!c.kept) return c.book
  const s = createSession(c.book)
  s.setLayout(c.kept.layout ?? [])
  if (c.kept.cells) s.setCells(c.kept.cells)
  if (c.kept.patches) s.patchCells(c.kept.patches)
  const out = hostSet(s.getBook(), c.kept)
  s.close()
  return out
}

/** A case's file: written by `writeXlsx`, or by `writeKept` over that file after its layout edits. */
export async function fieldsFileOf(c: FieldsCase): Promise<Uint8Array> {
  const data = (await writeXlsx({ book: c.book })).data
  if (!c.kept) return data
  const read = await openKept(data)
  const s = createSession(read.book, { cached: read.cached })
  s.setLayout(c.kept.layout ?? [])
  if (c.kept.cells) s.setCells(c.kept.cells)
  if (c.kept.patches) s.patchCells(c.kept.patches)
  const snap = s.getSnapshot()
  const out = (await writeKept(data, { ...snap, book: hostSet(snap.book, c.kept) })).data
  s.close()
  return out
}

/** The notes of a book, each shown all the time or not, in Excel's order: by sheet, by row and column. */
export const bookNotes = (b: BookInput): NoteShown[] =>
  b.sheets.flatMap((s) =>
    Object.entries(s.cells)
      .flatMap(([a, c]): NoteShown[] => (c.note ? [[s.name, a, c.note.shown === true]] : []))
      .sort((x, y) => { const [p, q] = [parseAddress(x[1]), parseAddress(y[1])]; return p.row - q.row || p.col - q.col }),
  )

/** The cells of a book whose text is turned, by their styles, in Excel's order: by sheet, by row and column. */
export const bookTurned = (b: BookInput): Turned[] =>
  b.sheets.flatMap((s) =>
    Object.entries(s.cells)
      .flatMap(([a, c]): Turned[] => {
        const r = c.style === undefined ? undefined : b.styles?.[c.style]?.alignment?.rotation
        return r ? [[s.name, a, r]] : []
      })
      .sort((x, y) => { const [p, q] = [parseAddress(x[1]), parseAddress(y[1])]; return p.row - q.row || p.col - q.col }),
  )

/** The calculation option of a book, as Excel names it. */
export const bookCalculation = (b: BookInput): string => b.calculation ?? 'automatic'

/** The panes in Excel's order (top left, top right, bottom left, bottom right) of the lines a split has. */
const PANES = ['topLeft', 'topRight', 'bottomLeft', 'bottomRight'] as const

/**
 * The window of each sheet of a book as Excel reads it: one pane at the view's scroll, or a pane on each side of each
 * line of its split, those after a line scrolled on its axis as the split's top left cell; frozen panes by that alone.
 */
export const bookPanes = (b: BookInput): SheetPanes[] =>
  b.sheets.map((s): SheetPanes => {
    if (s.frozen) return [s.name, true, false, [], 0]
    const at = (a: string | undefined) => (a ? parseAddress(a) : { row: 0, col: 0 })
    const [v, split] = [at(s.view?.topLeft), s.view?.split]
    if (!split) return [s.name, false, false, [[v.row + 1, v.col + 1]], 1]
    const t = at(split.topLeft)
    const names = PANES.filter((p) => (split.x || !p.endsWith('Right')) && (split.y || !p.startsWith('bottom')))
    const panes = names.map((p): [number, number] => [(p.startsWith('bottom') ? t.row : v.row) + 1, (p.endsWith('Right') ? t.col : v.col) + 1])
    return [s.name, false, true, panes, names.indexOf(split.activePane ?? 'topLeft') + 1]
  })

/** What Excel read of a window, as `bookPanes` gives it: of frozen panes only that they are. */
export const seenPanes = (p: SheetPanes): SheetPanes => (p[1] ? [p[0], true, false, [], 0] : p)

/** The split of each sheet's window where frozen panes do not take its place. */
export const splitOf = (b: BookInput) => b.sheets.map((s) => (s.frozen ? undefined : s.view?.split))

/** A split's lines where an Excel saves them: on a pixel of its screen (`pixel` points), found by `round`. */
export function onScreen(s: SplitPanes, pixel: number, round: (n: number) => number): SplitPanes {
  const out = { ...s }
  for (const axis of ['x', 'y'] as const) if (out[axis] !== undefined) out[axis] = round(out[axis] / pixel) * pixel
  return out
}

/** What the sheets of a book hold of the fields, as Excel reads them. */
export const bookFields = (b: BookInput): SheetFields[] => b.sheets.map((s) => [s.name, s.tabColor ?? null, s.view?.gridlines !== false, s.view?.headings !== false])
