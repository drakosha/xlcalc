/**
 * The ribbon's Number Format box of Excel for Windows over the codes of the built-in money formats written as a book's
 * own codes (a `numFmt` of its own, not the built-in's number): 5 to 8 and 41 to 44 by their codes in en-US and in
 * ru-RU as `builtinFormatCode` gives them, and en-US's with the dollar quoted, as Excel keeps the built-in's code.
 * The box of a built-in by number is in the sets `ribbon` and `ribbon-accounting`. Predicted by `numberFormatCategory`
 * before Excel ran it.
 */
import { builtinFormatCode } from '../../../src/builtin-format.js'
import { ruRU } from '../../../src/locales/ru-RU.js'
import { collector, type UiCase } from './cases.js'

const IDS = [5, 6, 7, 8, 41, 42, 43, 44]
const en = IDS.map((id) => builtinFormatCode(id)!)
const quoted = [5, 6, 7, 8, 42, 44].map((id) => builtinFormatCode(id)!.replace(/\$/g, '"$"'))
const ru = IDS.map((id) => builtinFormatCode(id, ruRU)!)

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const cells = Object.fromEntries([...en, ...quoted, ...ru].map((numberFormat, i) => [`B${i + 1}`, { value: -1234.5, numberFormat }]))
  add('code', { cells, act: { do: 'ribbon', cells: Object.keys(cells) } })
  return out
}
