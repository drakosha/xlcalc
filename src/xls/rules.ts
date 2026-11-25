import { MAX_COLS, MAX_ROWS } from '../model.js'
import { cornerOf, parseRange, type Rect } from '../structure/area.js'
import { rangeText } from '../validation/rules.js'
import { escapeAttr, escapeText } from '../xlsx/xml/text.js'
import { formulaText } from '../xlsb/formula.js'
import type { Payload } from '../xlsb/records.js'
import { validationSettings } from '../xlsb/validation.js'
import { parsedFormula } from './formula.js'
import { reference, unicode } from './records.js'
import type { Sheet, Workbook } from './workbook.js'

function boundedFormula(payload: Payload) {
  const size = payload.u16()
  payload.skip(2)
  return parsedFormula(payload.section(size), size)
}

export function validation(payload: Payload, book: Workbook, sheet: Sheet): string | undefined {
  const flags = payload.u32()
  const fields = ['promptTitle', 'errorTitle', 'prompt', 'error']
  let attrs = validationSettings(flags)
  for (const field of fields) {
    const value = unicode(payload)
    if (value && value !== '\0') attrs += ` ${field}="${escapeAttr(value)}"`
  }
  const formulas = [boundedFormula(payload), boundedFormula(payload)]
  const count = payload.u16()
  const refs: Rect[] = []
  for (let i = 0; i < count; i++) {
    const rect = parseRange(reference(payload))!
    // Validation covering all rows or columns of the binary sheet grows with the sheet in Excel.
    if (rect.r1 === 0 && rect.r2 === 65535) rect.r2 = MAX_ROWS - 1
    if (rect.c1 === 0 && rect.c2 === 255) rect.c2 = MAX_COLS - 1
    refs.push(rect)
  }
  if (!refs.length) return undefined
  const corner = cornerOf(refs)
  let body = ''
  for (const [i, formula] of formulas.entries()) {
    if (!formula.tokens.remaining) continue
    const decoded = formulaText(formula, book.context, corner.r, corner.c, true, sheet.name)
    if (!decoded || !('text' in decoded)) return undefined
    // An explicit list has NUL separators in BIFF and commas in SpreadsheetML.
    const explicit = Boolean(flags & 0x80)
    if (explicit && decoded.text.includes(',')) body += '<list/>'
    const text = explicit ? decoded.text.replaceAll('\0', ',') : decoded.text
    body += `<formula${i + 1}>${escapeText(text)}</formula${i + 1}>`
  }
  return `<dataValidation sqref="${refs.map((ref) => rangeText([ref])).join(' ')}"${attrs}>${body}</dataValidation>`
}
