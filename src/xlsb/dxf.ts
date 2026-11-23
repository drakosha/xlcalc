import { escapeAttr } from '../xlsx/xml/text.js'
import { colorAttrs } from './fields.js'
import type { Payload } from './records.js'
import { HORIZONTAL, LINES, PATTERNS, UNDERLINE, VERTICAL } from './styles.js'

const SIDES = ['top', 'bottom', 'left', 'right', 'diagonal', 'vertical', 'horizontal']
const FONT_FLAGS = ['i', 'strike', 'outline', 'shadow', 'condense', 'extend', 'charset', 'family']
const ALIGNMENT = ['horizontal', 'vertical', 'textRotation', 'indent', 'readingOrder', 'wrapText', 'justifyLastLine', 'shrinkToFit']
const ATTR = (key: string, value: string | number) => ` ${key}="${value}"`
const VAL = (key: string, value: string | number) => `<${key} val="${value}"/>`

/** The sized property list of BrtDXF; unknown future properties can be skipped without shifting later fields. */
export function differential(payload: Payload, future = false): string {
  if (future) payload.skip(4)
  payload.skip(4)
  const count = payload.u16()
  let numFmt: number | undefined
  let formatCode: string | undefined
  let font = ''
  let alignment = ''
  let border = ''
  let diagonal = ''
  let protection = ''
  let pattern = ''
  let gradient = ''
  let stops = ''
  let colors = ''
  const elementColor = (tag: string, field: Payload) => {
    const color = colorAttrs(field, true)
    return color ? `<${tag} ${color}/>` : ''
  }
  for (let i = 0; i < count; i++) {
    const type = payload.u16()
    const field = payload.section(payload.u16() - 4)
    switch (type) {
      case 0: pattern = ATTR('patternType', PATTERNS[field.u8()] ?? 'none'); break
      case 1: colors += elementColor('fgColor', field); break
      case 2: colors += elementColor('bgColor', field); break
      case 3: {
        const type = field.u32()
        const degree = field.f64()
        const left = field.f64()
        const right = field.f64()
        const top = field.f64()
        const bottom = field.f64()
        gradient = type ? ` type="path" left="${left}" right="${right}" top="${top}" bottom="${bottom}"` : ` degree="${degree}"`
        break
      }
      case 4: {
        field.u16()
        const position = field.f64()
        stops += `<stop position="${position}">${elementColor('color', field)}</stop>`
        break
      }
      case 5: font += elementColor('color', field); break
      case 6: case 7: case 8: case 9: case 10: case 11: case 12: {
        const color = elementColor('color', field)
        const style = LINES[field.u16()] ?? 'none'
        const side = SIDES[type - 6]
        border += `<${side} style="${style}">${color}</${side}>`
        break
      }
      case 13: case 14:
        diagonal += ATTR(type === 13 ? 'diagonalUp' : 'diagonalDown', field.u8())
        break
      case 15: alignment += ATTR(ALIGNMENT[0], HORIZONTAL[field.u8()] ?? 'general'); break
      case 16: alignment += ATTR(ALIGNMENT[1], VERTICAL[field.u8()] ?? 'bottom'); break
      case 17: case 19: case 20: case 21: case 22:
        alignment += ATTR(ALIGNMENT[type - 15], field.u8())
        break
      case 18: alignment += ATTR(ALIGNMENT[3], field.u16()); break
      case 23: field.u8(); break
      case 24: font += VAL('name', escapeAttr(field.wide(field.u16()))); break
      case 25: font += VAL('b', field.u16() >= 700 ? 1 : 0); break
      case 26: font += VAL('u', UNDERLINE[field.u16()] ?? 'none'); break
      case 27: font += VAL('vertAlign', ['baseline', 'superscript', 'subscript'][field.u16()] ?? 'baseline'); break
      case 28: case 29: case 30: case 31: case 32: case 33: case 34: case 35:
        font += VAL(FONT_FLAGS[type - 28], field.u8())
        break
      case 36: font += VAL('sz', field.u32() / 20); break
      case 37: font += VAL('scheme', ['none', 'major', 'minor'][field.u8()] ?? 'none'); break
      case 38: formatCode = field.wide(field.u16()); break
      case 41: numFmt = field.u16(); break
      case 42: field.i16(); break
      case 43: case 44: protection += ATTR(type === 43 ? 'locked' : 'hidden', field.u8()); break
      default: field.skip(field.remaining)
    }
    field.done()
  }
  payload.done()
  const fill = (pattern || colors ? `<patternFill${pattern}>${colors}</patternFill>` : '')
    + (gradient ? `<gradientFill${gradient}>${stops}</gradientFill>` : '')
  const format = (numFmt === undefined ? '' : ATTR('numFmtId', numFmt))
    + (formatCode === undefined ? '' : ATTR('formatCode', escapeAttr(formatCode)))
  return (font ? `<font>${font}</font>` : '') + (fill ? `<fill>${fill}</fill>` : '')
    + (border || diagonal ? `<border${diagonal}>${border}</border>` : '')
    + (alignment ? `<alignment${alignment}/>` : '') + (format ? `<numFmt${format}/>` : '')
    + (protection ? `<protection${protection}/>` : '')
}
