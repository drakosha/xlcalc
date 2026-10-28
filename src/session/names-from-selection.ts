/**
 * Excel's Create Names from Selection (`Session.getNamesFromSelection`, the set `bench/coverage/select`): the labels of
 * the chosen sides of a range name the cells of their column or row past the labels, and with a row and a column of
 * labels, the label in their corner names all of those cells. A label's text, its spaces at the ends left out, is made
 * a name: each character a name does not take becomes `_`, `_` goes before a first character a name cannot start with
 * and before an R1C1 reference (`_R2C3`), and after a cell's address, `R`, `C` or a boolean (`Q1_`, `TRUE_`). A text
 * label and a date, by the text it shows, make names; a number, a time of day, a boolean and an error make none, nor a
 * label without text (the set `names-held`). Each name refers to its cells absolutely, with its sheet.
 */
import type { NameInput, SelectionLabels } from '../contract.js'
import type { EngineState } from '../engine.js'
import { InputError } from '../input-error.js'
import { checkBoolean, object, onlyKeys } from '../input/basic.js'
import { localeBooleans } from '../locale/registry.js'
import { formatAddress } from '../model.js'
import { shownAs } from '../numfmt/shown-as.js'
import { sheetLines, type SheetLines } from '../layout/lines.js'
import { cellValue } from '../result.js'
import type { Rect } from '../structure/area.js'
import { caseKey } from '../collate.js'
import { reservedName } from '../syntax/check.js'
import { quoteSheet } from '../syntax/print.js'
import { IDENT_FIRST, IDENT_NEXT } from '../syntax/scan.js'

const SIDES = ['top', 'left', 'bottom', 'right'] as const

/** The sides of labels, checked: at least one. */
export function checkLabels(v: unknown): SelectionLabels {
  const o = object(v, 'labels', 'the sides of labels such as { top: true, left: true }')
  onlyKeys(o, 'labels', [...SIDES])
  for (const k of SIDES) if (o[k] !== undefined) checkBoolean(o[k], `labels.${k}`)
  if (!SIDES.some((k) => o[k])) throw new InputError('labels', 'expected at least one side of labels: top, left, bottom or right')
  return o as SelectionLabels
}

const FIRST = new RegExp(`^${IDENT_FIRST}`, 'u')
const NEXT = new RegExp(IDENT_NEXT, 'u')

/** An R1C1 reference with a number (`R2C3`, `R2`, `C3`), which Excel puts `_` before. */
const R1C1 = /^[Rr]\d+([Cc]\d*)?$|^[Rr]?[Cc]\d+$/

/** A label's text made a name as Excel makes it; undefined for a label without text. */
export function labelName(label: string, booleans: readonly string[]): string | undefined {
  const text = label.replace(/^ +| +$/g, '')
  if (!text) return undefined
  let name = [...text].map((ch) => (NEXT.test(ch) ? ch : '_')).join('')
  if (!FIRST.test(name) || R1C1.test(name)) name = `_${name}`
  else if (reservedName(name, booleans)) name = `${name}_`
  return [...name].slice(0, 255).join('')
}

/** A cell's value as a label's text: text as it is, a date as it shows; none for a number, a time of day, a boolean or an error. */
function labelText(state: EngineState, s: number, r: number, c: number, formats: SheetLines): string {
  const v = cellValue(state, s, r, c)
  if (typeof v === 'string') return v
  if (typeof v !== 'number') return ''
  const format = formats.format(r, c)
  return shownAs(format) === 'date' ? state.ctx.rules.display(v, format, Infinity).text : ''
}

/** An absolute reference to a rectangle of the sheet, with the sheet. */
function absolute(sheet: string, x: Rect): string {
  const cell = (r: number, c: number) => formatAddress(r, c).replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2')
  return `${quoteSheet(sheet)}!${cell(x.r1, x.c1)}${x.r1 === x.r2 && x.c1 === x.c2 ? '' : `:${cell(x.r2, x.c2)}`}`
}

export function namesFromSelection(state: EngineState, s: number, x: Rect, labels: SelectionLabels): NameInput[] {
  const sheet = state.input.sheets[s].name
  const booleans = localeBooleans(state.input.locale)
  // the cells the labels name: the range past the rows and columns of labels
  const inner: Rect = { r1: x.r1 + (labels.top ? 1 : 0), r2: x.r2 - (labels.bottom ? 1 : 0), c1: x.c1 + (labels.left ? 1 : 0), c2: x.c2 - (labels.right ? 1 : 0) }
  if (inner.r1 > inner.r2 || inner.c1 > inner.c2) return []
  const formats = sheetLines(state.input.sheets[s], state)
  const out = new Map<string, NameInput>()
  const add = (r: number, c: number, cells: Rect) => {
    const name = labelName(labelText(state, s, r, c, formats), booleans)
    if (name === undefined) return
    out.delete(caseKey(name))
    out.set(caseKey(name), { name, ref: absolute(sheet, cells) })
  }
  for (const [side, row] of [['top', x.r1], ['bottom', x.r2]] as const) if (labels[side]) for (let c = inner.c1; c <= inner.c2; c++) add(row, c, { ...inner, c1: c, c2: c })
  for (const [side, col] of [['left', x.c1], ['right', x.c2]] as const) if (labels[side]) for (let r = inner.r1; r <= inner.r2; r++) add(r, col, { ...inner, r1: r, r2: r })
  for (const [v, h, r, c] of [['top', 'left', x.r1, x.c1], ['top', 'right', x.r1, x.c2], ['bottom', 'left', x.r2, x.c1], ['bottom', 'right', x.r2, x.c2]] as const) if (labels[v] && labels[h]) add(r, c, inner)
  return [...out.values()]
}
