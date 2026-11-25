import { indexToCol } from '../model.js'
import { generalText } from '../numfmt/general.js'
import { errorText } from '../xlsb/fields.js'
import { textLiteral, type XlsFormula, type Formula, type LegacyLink } from '../xlsb/formula.js'
import type { Payload } from '../xlsb/records.js'
import type { LegacyEncoding } from './encoding.js'
import { unicode } from './records.js'

const biff8: XlsFormula = {
  string: (payload) => unicode(payload, payload.u8()),
  array: constantArray,
  label: labelRange,
}

/** Excel upgrades BIFF8 natural-language labels to ranges. Legacy intersection is applied when reading the formula. */
function labelRange(type: number, tokens: Payload, extra: Payload, readArea: () => string): string | undefined {
  if (tokens.remaining < 4) return undefined
  if (type === 1 || type === 0x10) {
    tokens.skip(4)
    if (type === 0x10 && labelArea(tokens, readArea) === undefined) return undefined
    return '#NAME?'
  }
  const byRow = type === 2 || type === 6
  const stackedColumn = type === 0x0d || type === 0x0f
  const byColumn = type === 3 || type === 7 || stackedColumn
  const byArea = type === 0x0a || type === 0x0b
  if (!byRow && !byColumn && !byArea) return undefined

  let location = tokens
  let labelRelative: boolean | undefined
  if (stackedColumn || type === 0x0b) {
    if (extra.remaining < 4) return undefined
    tokens.skip(4)
    const flags = extra.u32()
    const count = flags & 0x3fffffff
    if (!count || count * 4 > extra.remaining) return undefined
    // The innermost label ends the stack. The extra record holds its own relativity bit.
    extra.skip((count - 1) * 4)
    location = extra
    labelRelative = Boolean(flags >>> 31)
  }
  const row = location.u16()
  const packedColumn = location.u16()
  const column = packedColumn & 0x3fff
  if (column > 255) return undefined
  if (byArea) return labelArea(tokens, readArea)
  const relative = labelRelative ?? Boolean(packedColumn & 0x8000)
  const absolute = relative ? '' : '$'
  if (byRow) return `${absolute}${row + 1}:${absolute}${row + 1}`
  if (row === 65535) return '#REF!'
  const col = absolute + indexToCol(column)
  return `${col}$${row + 2}:${col}$65536`
}

function labelArea(tokens: Payload, readArea: () => string): string | undefined {
  // PtgElfRadical qualifies the PtgArea (or PtgAreaErr) that follows it.
  if (tokens.remaining < 9) return undefined
  const range = tokens.u8()
  if (range < 0x20) return undefined
  switch (range & 31) {
    case 5: return readArea()
    case 11:
      tokens.skip(8)
      return '#REF!'
    default: return undefined
  }
}

function constantArray(payload: Payload, encoding?: LegacyEncoding): string | undefined {
  const width = payload.u8()
  const height = payload.u16()
  const columns = encoding ? width || 256 : width + 1
  const rows = encoding ? height : height + 1
  // A SerAr with an empty string is shortest: byte text in BIFF2–5, Unicode in BIFF8.
  const shortest = encoding ? 2 : 4
  if (columns * rows * shortest > payload.remaining) payload.fail('invalid formula constant dimensions')
  const lines: string[] = []
  for (let row = 0; row < rows; row++) {
    const line: string[] = []
    for (let column = 0; column < columns; column++) {
      switch (payload.u8()) {
        case 0: payload.skip(8); line.push(''); break
        case 1: line.push(generalText(payload.f64(), 21)); break
        case 2: line.push(textLiteral(encoding ? encoding.text(payload, payload.u8()) : unicode(payload))); break
        case 4: line.push(payload.u8() ? 'TRUE' : 'FALSE'); payload.skip(7); break
        case 16: line.push(errorText(payload.u8())); payload.skip(7); break
        default: return undefined
      }
    }
    lines.push(line.join(','))
  }
  return '{' + lines.join(';') + '}'
}

export function parsedFormula(payload: Payload, length = payload.u16(), encoding?: LegacyEncoding, links?: readonly LegacyLink[]): Formula {
  const xls: XlsFormula = encoding ? {
    old: encoding.version,
    links,
    string: (tokens) => encoding.text(tokens, tokens.u8()),
    array: (values) => constantArray(values, encoding),
  } : biff8
  return { tokens: payload.section(length), extra: payload.section(payload.remaining), xls }
}
