/**
 * Texts typed into cells (`Session.setCellText`, `Session.setCellTexts`): each read as `text.ts` reads text typed into its
 * cell, in order, so a cell given twice is read by what the earlier text wrote into it and keeps the later; the first
 * text Excel would not take stops them all. With the texts written: the width each column takes, the largest a text
 * typed alone into one of its cells gives it, and the cells whose values their rules of data validation do not take.
 */
import type { CellRef, CellReplace, CellTextRefusal, CellValidationAlert, ColumnWidth, Refusal } from '../contract.js'
import type { EngineState } from '../engine.js'
import { array, checkAddress, checkSheet, checkString, object, onlyKeys } from '../input.js'
import { formatAddress, indexToCol, parseAddress } from '../model.js'
import { cellText, columnWidening, textRefusal, validationAlert } from './text.js'

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

/** The cells after the texts are typed in order at the moment `now`, each cell once, in the order first given; or the first text Excel stops, and its cell. */
export function typedTexts(s: EngineState, typings: Typing[], now: string): { ok: true; cells: Typed[] } | { ok: false; refusal: Refusal<CellTextRefusal>; ref: CellRef } {
  const cells = new Map<string, Typed>()
  for (const { i, r, c, text } of typings) {
    const ref = { sheet: s.sheetNames[i], cell: formatAddress(r, c) }
    const k = `${i}:${ref.cell}`
    const earlier = cells.get(k)
    const typed = cellText(s, i, r, c, earlier ? (earlier.edit.input ?? undefined) : s.input.sheets[i].cells[ref.cell], text, now)
    if (!typed.ok) return { ok: false, refusal: typed, ref }
    const edit = { ...ref, input: typed.input }
    const refused = textRefusal(s, edit)
    if (refused) return { ok: false, refusal: refused, ref }
    cells.set(k, { i, r, c, edit })
  }
  return { ok: true, cells: [...cells.values()] }
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

/** An answer's widths and alerts, those there are. */
export const textsFound = (widths: ColumnWidth[], alerts: CellValidationAlert[]) => ({ ...(widths.length ? { widths } : {}), ...(alerts.length ? { validations: alerts } : {}) })
