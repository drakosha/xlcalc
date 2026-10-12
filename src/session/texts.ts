/**
 * Texts typed into cells (`Session.setCellText`, `Session.setCellTexts`): each read as `text.ts` reads text typed into its
 * cell, in order, so a cell given twice is read by what the earlier text wrote into it and keeps the later; the first
 * text Excel would not take stops them all. Text pasted from the clipboard (`Session.pasteText`): its rectangle's fields,
 * each read as pasted into its cell. A text entered into a selection (`Session.fillCellText`): read once in the active
 * cell, which every cell copies. With the texts written: the width each column takes, the largest a text typed alone
 * into one of its cells gives it, and the cells whose values their rules of data validation do not take, the first of
 * which refuses the texts where the call asks it to (`refuseInvalid`).
 */
import type { CellRef, CellReplace, CellTextRefusal, CellValidationAlert, CellValidationRefusal, ColumnWidth, PasteTextRefusal, PreviewCell, Refusal } from '../contract.js'
import type { EngineState } from '../engine.js'
import { clipboardRows } from '../entry/clipboard.js'
import { spend } from '../exec/budget.js'
import { array, checkAddress, checkSheet, checkString, InputError, object, onlyKeys } from '../input.js'
import { say } from '../locale/messages.js'
import { formatAddress, indexToCol, MAX_COLS, MAX_ROWS, parseAddress } from '../model.js'
import { cellCount, hasCell, rectText, type Rect } from '../structure/area.js'
import { formulaCopier } from '../structure/copy.js'
import { arrayRefusal, cellText, columnWidening, headerRefusal, readText, textRefusal, validationAlert, withTyped } from './text.js'
import { cellValue } from './looks.js'
import { MAX_CELLS } from './styles.js'

/** Text typed into the cell at row `r`, column `c` of sheet `i`. */
export interface Typing {
  i: number
  r: number
  c: number
  text: string
}

/** A cell typed into and its edit, the input its texts give it. */
export interface Typed {
  i: number
  r: number
  c: number
  edit: CellReplace
}

/** `Session.setCellTexts`' entries checked: the cells and their texts. */
export function checkTextEntries(v: unknown, sheetIndex: (name: string) => number | undefined, path = 'entries'): Typing[] {
  return array(v, path, 'an array of texts such as [{ sheet: "Sheet1", cell: "A1", text: "15.03.2026" }]').map((e, k) => {
    const p = `${path}[${k}]`
    const o = object(e, p, 'a text such as { sheet: "Sheet1", cell: "A1", text: "15.03.2026" }')
    onlyKeys(o, p, ['sheet', 'cell', 'text'])
    const i = checkSheet(o.sheet, sheetIndex, `${p}.sheet`)
    const { row, col } = parseAddress(checkAddress(o.cell, `${p}.cell`))
    return { i, r: row, c: col, text: checkString(o.text, `${p}.text`) }
  })
}

/** Cells with their texts read, or the first one Excel would not take, why and its cell. */
export type TypedTexts = { ok: true; cells: Typed[] } | { ok: false; refusal: Refusal<CellTextRefusal>; ref: CellRef }

/**
 * The cells after the texts are typed in order at the moment `now`, each cell once, in the order first given; or the
 * first text Excel stops, and its cell. Pasted (`pasted`), each is read as pasted and the cells are taken together
 * (`takenTogether`).
 */
export function typedTexts(s: EngineState, typings: Typing[], now: string, pasted = false): TypedTexts {
  const cells = new Map<string, Typed>()
  for (const { i, r, c, text } of typings) {
    spend(1)
    const ref = { sheet: s.sheetNames[i], cell: formatAddress(r, c) }
    const k = `${i}:${ref.cell}`
    const earlier = cells.get(k)
    const typed = cellText(s, i, r, c, earlier ? (earlier.edit.input ?? undefined) : s.input.sheets[i].cells[ref.cell], text, now, pasted)
    if (!typed.ok) return { ok: false, refusal: typed, ref }
    const edit = { ...ref, input: typed.input }
    const refused = pasted ? null : textRefusal(s, edit)
    if (refused) return { ok: false, refusal: refused, ref }
    cells.set(k, { i, r, c, edit })
  }
  return pasted ? takenTogether(s, [...cells.values()]) : { ok: true, cells: [...cells.values()] }
}

/**
 * Cells written together, a paste or an entry into a selection: refused where they change part of an array formula (an
 * area they cover whole they replace) or put a formula into a table's header cell, as Excel refuses them, with the cell.
 */
function takenTogether(s: EngineState, cells: Typed[]): TypedTexts {
  const part = arrayRefusal(s, cells.map((x) => x.edit))
  if (part) {
    const { sheet, cell, ...refusal } = part
    return { ok: false, refusal, ref: { sheet, cell } }
  }
  for (const { edit } of cells) {
    const refused = headerRefusal(s, edit)
    if (refused) return { ok: false, refusal: refused, ref: { sheet: edit.sheet, cell: edit.cell } }
  }
  return { ok: true, cells }
}

/**
 * Text pasted from the clipboard at row `r`, column `c` of sheet `i` at the moment `now`: the cells of its rectangle
 * (`clipboardRows`) from there, each field read as pasted into its cell, and the range they fill; or why Excel pastes
 * nothing, with the cell, and a rectangle past the sheet's edge.
 */
export function pastedText(s: EngineState, i: number, r: number, c: number, text: string, now: string): { ok: true; cells: Typed[]; range: string } | { ok: false; refusal: PasteTextRefusal } {
  const rows = clipboardRows(text)
  const x = { r1: r, c1: c, r2: r + rows.length - 1, c2: c + rows[0].length - 1 }
  const sheet = s.sheetNames[i]
  const cell = formatAddress(r, c)
  if (x.r2 >= MAX_ROWS || x.c2 >= MAX_COLS) return { ok: false, refusal: { ok: false, code: 'off_sheet', message: say(s.input.locale, { key: 'copyOffSheet', sheet, cell }), sheet, cell } }
  const typed = typedTexts(s, rows.flatMap((row, dr) => row.map((field, dc) => ({ i, r: r + dr, c: c + dc, text: field }))), now, true)
  // read as pasted, a text Excel does not take as a formula is text: what refuses a paste is its cells together
  return typed.ok ? { ...typed, range: rectText(x) } : { ok: false, refusal: { ...typed.refusal, ...typed.ref } as PasteTextRefusal }
}

/** `FillTextOptions.active` checked: the cell of the range `x` the text is typed into, its top-left cell by default. */
export function checkActive(v: unknown, x: Rect, path = 'opts.active'): { r: number; c: number } {
  if (v === undefined) return { r: x.r1, c: x.c1 }
  const { row, col } = parseAddress(checkAddress(v, path))
  if (!hasCell(x, row, col)) throw new InputError(path, `expected a cell of the range ${rectText(x)}, got ${JSON.stringify(v)}`)
  return { r: row, c: col }
}

/**
 * `text` entered into every cell of the range `x` of sheet `i` from the active cell (`a`) at the moment `now`, as Excel's
 * Control+Return enters it (the set `ui/fill-held`): read once in the active cell (`readText`), each cell takes what the
 * active cell holds after it, a formula moved as copied from the active cell (`formulaCopier`), and its number format,
 * where it has one; a cell keeps its own format where the active cell shows General. The cells row by row, or why Excel
 * stops the entry, with the cell.
 */
export function filledText(s: EngineState, i: number, x: Rect, a: { r: number; c: number }, text: string, now: string): TypedTexts {
  if (cellCount(x) > MAX_CELLS) throw new InputError('range', `a range of ${cellCount(x)} cells: an entry goes into ${MAX_CELLS} cells at most`)
  const cells = s.input.sheets[i].cells
  const sheet = s.sheetNames[i]
  const typed = readText(s, i, a.r, a.c, cells[formatAddress(a.r, a.c)], text, now)
  if (!typed.ok) return { ok: false, refusal: typed as Refusal<CellTextRefusal>, ref: { sheet, cell: formatAddress(a.r, a.c) } }
  const formula = typed.input?.formula
  const copy = formula === undefined ? null : formulaCopier(s.input, s.tables)
  const from = { sheet, row: a.r, col: a.c }
  const out: Typed[] = []
  for (let r = x.r1; r <= x.r2; r++)
    for (let c = x.c1; c <= x.c2; c++) {
      spend(1)
      const cell = formatAddress(r, c)
      const input = copy && (r !== a.r || c !== a.c) ? { ...typed.input, formula: copy(formula!, from, { sheet, row: r, col: c }) } : typed.input
      out.push({ i, r, c, edit: { sheet, cell, input: withTyped(s, i, r, c, cells[cell], input) } })
    }
  return takenTogether(s, out)
}

/** The width each column of the cells takes for the values they hold now, the largest of its cells' (`columnWidening`); none for a column that keeps its width. */
export function textWidths(s: EngineState, cells: Typed[]): ColumnWidth[] {
  const out = new Map<string, ColumnWidth>()
  for (const { i, r, c } of cells) {
    const width = columnWidening(s, i, r, c)
    if (width === undefined) continue
    const x = out.get(`${i}:${c}`)
    if (!x) out.set(`${i}:${c}`, { sheet: s.sheetNames[i], column: indexToCol(c), width })
    else if (width > x.width) x.width = width
  }
  return [...out.values()]
}

/** The cells whose values their rules of data validation do not take now, with the rules' alerts. */
export const textAlerts = (s: EngineState, cells: Typed[]): CellValidationAlert[] =>
  cells.flatMap(({ i, r, c, edit }) => {
    const alert = validationAlert(s, i, r, c)
    return alert ? [{ sheet: edit.sheet, cell: edit.cell, ...alert }] : []
  })

/** A cell's alert as the refusal of its value (`CellTextOptions.refuseInvalid`): the cell, the rule, its style and message. */
export const alertRefusal = ({ valid: _, ...alert }: CellValidationAlert): CellValidationRefusal => ({ ok: false, code: 'validation', ...alert })

const STRICTNESS = { stop: 2, warning: 1, information: 0 }

/**
 * Texts refused for the strictest rule their values break, the first cell of that style: a host that answers a warning
 * with Yes calls again without the option, which would write a value a later cell's stop does not take.
 */
export const strictestRefusal = (alerts: CellValidationAlert[]): CellValidationRefusal | null =>
  alerts.length ? alertRefusal(alerts.reduce((a, b) => (STRICTNESS[b.severity] > STRICTNESS[a.severity] ? b : a))) : null

/** An answer's widths and alerts, those there are. */
export const textsFound = (widths: ColumnWidth[], alerts: CellValidationAlert[]) => ({ ...(widths.length ? { widths } : {}), ...(alerts.length ? { validations: alerts } : {}) })

/** The cells of a preview: each one's input and the value it computes now. */
export const previewCells = (s: EngineState, cells: Typed[]): PreviewCell[] => cells.map(({ i, r, c, edit }) => ({ sheet: edit.sheet, cell: edit.cell, input: edit.input, value: cellValue(s, i, r, c) }))
