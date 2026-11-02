/**
 * The ribbon's Number Format box of Excel for Windows in ru-RU over the codes a file may give ru-RU's built-in short
 * date (14, 27 to 31, 36), date and time (22) and fractions (12, 13) as a book's own codes (a `numFmt` of its own):
 * in en-US syntax with `/` for the date's separator and the fraction's space quoted, as Excel reads them back for the
 * built-ins, with ru-RU's dots and the space bare, and en-US's codes; the built-ins by number beside them. Predicted by
 * `numberFormatCategory` before Excel ran it.
 */
import { collector, type UiCase } from './cases.js'

const IDS = [12, 13, 14, 22]
/** Each a code of the book's own even where the writer would store a built-in's number for it (`own`). */
const CODES = [
  'dd/mm/yyyy', '#" "?/?', '#" "??/??', 'dd/mm/yyyy h:mm',
  'dd.mm.yyyy', '# ?/?', '# ??/??', 'dd.mm.yyyy h:mm',
  'm/d/yyyy', 'm/d/yyyy h:mm', 'dd/mm/yy', 'd/m/yyyy',
]

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const cellsOf = (formats: (string | number)[]) => Object.fromEntries(formats.map((numberFormat, i) => [`B${i + 1}`, { value: 45688.75, numberFormat }]))
  const ids = cellsOf(IDS)
  add('builtin', { cells: ids, act: { do: 'ribbon', cells: Object.keys(ids) } })
  const codes = cellsOf(CODES)
  add('code', { cells: codes, act: { do: 'ribbon', cells: Object.keys(codes) }, own: Object.keys(codes) })
  return out
}
