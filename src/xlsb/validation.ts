import { escapeAttr, escapeText } from '../xlsx/xml/text.js'
import { formula, formulaText } from './formula.js'
import { future, rangeOrigin, ranges } from './future.js'
import type { Part } from './parts.js'
import type { Payload } from './records.js'
import type { Xml } from './xml.js'

const TYPES = ['none', 'whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom']
const OPERATORS = ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual']
const STYLES = ['stop', 'warning', 'information']
const IME = ['noControl', 'on', 'off', 'disabled', 'hiragana', 'fullKatakana', 'halfKatakana', 'fullAlpha', 'halfAlpha', 'fullHangul', 'halfHangul']

export function validation(part: Part, type: number, payload: Payload, extensions: Xml): boolean {
  const { xml, context, pending } = part
  switch (type) {
    case 573:
      payload.skip(14)
      xml.add(`<dataValidations count="${payload.u32()}">`)
      break
    case 574: xml.add('</dataValidations>'); break
    case 64: {
      const flags = payload.u32()
      const refs = ranges(payload)
      if (!refs.length) payload.fail('data validation has no ranges')
      const attributes = validationAttrs(flags, payload)
      const corner = rangeOrigin(refs)
      xml.add(`<dataValidation sqref="${refs.join(' ')}"${attributes}>`)
      for (const tag of ['formula1', 'formula2']) {
        const parsed = formula(payload)
        if (!parsed.tokens.remaining) {
          parsed.extra.done()
          continue
        }
        const decoded = formulaText(parsed, context, corner.row, corner.col, true)
        if (!decoded || !('text' in decoded)) {
          pending()
          continue
        }
        const list = ((flags & 15) === 3 && /^".*"$/.test(decoded.text)) ? decoded.text.replace(/\0/g, ',') : decoded.text
        xml.add(`<${tag}>${escapeText(list)}</${tag}>`)
      }
      xml.add('</dataValidation>')
      break
    }
    case 1054:
      payload.skip(18)
      extensions.add(`<ext><x14:dataValidations count="${payload.u32()}">`)
      break
    case 1154:
      extensions.add('</x14:dataValidations></ext>')
      break
    case 1053: {
      const header = future(payload)
      const flags = payload.u32()
      const attrs = validationAttrs(flags, payload)
      if (!header.ranges.length) payload.fail('extended data validation has no ranges')
      const corner = rangeOrigin(header.ranges)
      extensions.add(`<x14:dataValidation${attrs}>`)
      for (let i = 0; i < header.formulas.length; i++) {
        const value = formulaText(header.formulas[i], context, corner.row, corner.col, true)
        if (!value || !('text' in value)) pending()
        else extensions.add(`<x14:formula${i + 1}><xm:f>${escapeText(value.text)}</xm:f></x14:formula${i + 1}>`)
      }
      extensions.add(`<xm:sqref>${header.ranges.join(' ')}</xm:sqref></x14:dataValidation>`)
      break
    }
    default: return false
  }
  payload.done()
  return true
}

function validationAttrs(flags: number, payload: Payload): string {
  const type = TYPES[flags & 15]
  const style = STYLES[(flags >>> 4) & 7]
  const blank = (flags >>> 8) & 1
  const dropdown = (flags >>> 9) & 1
  const ime = IME[(flags >>> 10) & 0xff]
  const showInput = (flags >>> 18) & 1
  const showError = (flags >>> 19) & 1
  const operator = OPERATORS[(flags >>> 20) & 15]
  const strings = ['errorTitle', 'error', 'promptTitle', 'prompt'].map((key) => {
    const value = payload.nullableWide()
    return value === undefined ? '' : ` ${key}="${escapeAttr(value)}"`
  }).join('')
  return ` type="${type ?? 'unknown'}" allowBlank="${blank}" showDropDown="${dropdown}" showInputMessage="${showInput}" showErrorMessage="${showError}"`
    + (style ? ` errorStyle="${style}"` : '') + (operator ? ` operator="${operator}"` : '') + (ime ? ` imeMode="${ime}"` : '') + strings
}
