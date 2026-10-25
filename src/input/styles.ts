/**
 * The checks of input's looks: number formats, the index of a style of the book, and styles with their fonts, fills,
 * borders and alignments, as a cell's style or a conditional format's. The module needs nothing of the engine.
 */
import type { AlignmentStyle, BorderLine, CellStyle } from '../contract.js'
import { InputError, what } from '../input-error.js'
import { isBuiltin } from '../numfmt/builtin.js'
import { array, checkBoolean, checkNumber, checkOneOf, checkString, isObject, object, onlyKeys } from './basic.js'

/** A number format: a code, or the number of a built-in format (0 to 49). */
export function checkFormat(v: unknown, path: string) {
  if (typeof v !== 'string' && !(typeof v === 'number' && isBuiltin(v))) throw new InputError(path, `expected a format code or the number of a built-in format, 0 to 49, got ${what(v)}`)
}

/** A style's fields where a cell takes its number: where such a style goes. */
const STYLE_HINT = " (a cell refers to one of the workbook's styles by its index; a session's patchCells takes a style by its fields)"

/** A style number of a book with `styles` styles. */
export function checkStyleIndex(v: unknown, path: string, styles: number) {
  if (typeof v !== 'number' || !Number.isInteger(v) || v < 0 || v >= styles) {
    const hint = isObject(v) ? STYLE_HINT : ''
    throw new InputError(path, (styles ? `expected the index of a style of the workbook, 0 to ${styles - 1}, got ${what(v)}` : `the workbook has no styles, got ${what(v)}`) + hint)
  }
}

const COLOR = /^#[0-9A-Fa-f]{6}$/

export function checkColor(v: unknown, path: string) {
  if (typeof v !== 'string' || !COLOR.test(v)) throw new InputError(path, `expected a color such as "#FF0000", got ${what(v)}`)
}

/** The lines of a border (`BorderLine`), as xlsx names them; the xlsx reader takes these and names any other lost. */
export const LINES: readonly BorderLine['style'][] = ['thin', 'medium', 'thick', 'dotted', 'dashed', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']

/** A cell's alignments (`AlignmentStyle`), as xlsx names them. */
export const HORIZONTAL: readonly NonNullable<AlignmentStyle['horizontal']>[] = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed']
export const VERTICAL: readonly NonNullable<AlignmentStyle['vertical']>[] = ['top', 'center', 'bottom', 'justify', 'distributed']

/** Whether a field of a style is given: an undefined one never is, as no field of input is; in a patch a null one neither (`checkStyle`). */
export type Given = (o: Record<string, unknown>, k: string) => boolean

/** The rule of `Given`: a field is given unless undefined, or null in a patch, a field removed; `givenIn(patch)` makes the predicate. */
export const givenIn = (patch: boolean): Given => (o, k) => o[k] !== undefined && !(patch && o[k] === null)

/** A font's fields: a cell's (`named`) has a name and a size too, which a conditional style (a dxf) has not. */
export function checkFont(v: unknown, p: string, given: Given, named: boolean) {
  const f = object(v, p, 'a font such as { bold: true }')
  onlyKeys(f, p, [...(named ? ['name', 'size'] : []), 'bold', 'italic', 'underline', 'strike', 'color'])
  if (given(f, 'name')) checkString(f.name, `${p}.name`)
  if (given(f, 'size')) checkNumber(f.size, `${p}.size`, 1, 409)
  for (const k of ['bold', 'italic', 'strike']) if (given(f, k)) checkBoolean(f[k], `${p}.${k}`)
  if (given(f, 'underline')) checkOneOf(f.underline, `${p}.underline`, ['single', 'double'])
  if (given(f, 'color')) checkColor(f.color, `${p}.color`)
}

/** A solid fill; `example` is the colour its message shows. */
export function checkFill(v: unknown, p: string, example: string) {
  const f = object(v, p, `a fill such as { color: "${example}" }`)
  onlyKeys(f, p, ['color'])
  checkColor(f.color, `${p}.color`)
}

/** A border's sides; `side` says which of those it names are given. */
export function checkBorder(v: unknown, p: string, side: Given) {
  const b = object(v, p, 'a border such as { bottom: { style: "thin" } }')
  onlyKeys(b, p, ['top', 'right', 'bottom', 'left'])
  for (const [k, l] of Object.entries(b)) {
    if (!side(b, k)) continue
    const lp = `${p}.${k}`
    const line = object(l, lp, 'a line such as { style: "thin", color: "#000000" }')
    onlyKeys(line, lp, ['style', 'color'])
    checkOneOf(line.style, `${lp}.style`, LINES)
    if (line.color !== undefined) checkColor(line.color, `${lp}.color`)
  }
}

/** A style, or with `patch` the fields of one to change (`StylePatch`), where null removes a field. */
export function checkStyle(v: unknown, p: string, patch = false): CellStyle {
  const s = object(v, p, 'a style such as { font: { bold: true } }')
  onlyKeys(s, p, ['font', 'fill', 'border', 'alignment'])
  const given: Given = givenIn(patch)
  if (given(s, 'font')) checkFont(s.font, `${p}.font`, given, true)
  if (given(s, 'fill')) checkFill(s.fill, `${p}.fill`, '#FFC000')
  if (given(s, 'border')) checkBorder(s.border, `${p}.border`, given)
  if (given(s, 'alignment')) {
    const a = object(s.alignment, `${p}.alignment`, 'an alignment such as { horizontal: "center" }')
    onlyKeys(a, `${p}.alignment`, ['horizontal', 'vertical', 'wrap', 'indent'])
    if (given(a, 'horizontal')) checkOneOf(a.horizontal, `${p}.alignment.horizontal`, HORIZONTAL)
    if (given(a, 'vertical')) checkOneOf(a.vertical, `${p}.alignment.vertical`, VERTICAL)
    if (given(a, 'wrap')) checkBoolean(a.wrap, `${p}.alignment.wrap`)
    if (given(a, 'indent')) checkNumber(a.indent, `${p}.alignment.indent`, 0, 250, true)
  }
  return s as CellStyle
}

export function checkStyles(v: unknown, path: string): CellStyle[] {
  return array(v, path, 'an array of styles').map((s, i) => checkStyle(s, `${path}[${i}]`))
}
