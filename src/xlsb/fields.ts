import { MAX_COLS, MAX_ROWS, formatAddress } from '../model.js'
import type { Payload } from './records.js'

export function address(payload: Payload, row: number, column: number): string {
  if (row < 0 || row >= MAX_ROWS || column < 0 || column >= MAX_COLS) payload.fail(`cell out of bounds (row ${row}, column ${column})`)
  return formatAddress(row, column)
}

/** RfX: the two row numbers precede the two column numbers. */
export function range(payload: Payload): string {
  return rangeAt(payload, payload.u32(), payload.u32(), payload.u32(), payload.u32())
}

export function rangeAt(payload: Payload, top: number, bottom: number, left: number, right: number): string {
  if (top > bottom || left > right) payload.fail('range ends before it starts')
  const first = address(payload, top, left)
  const last = address(payload, bottom, right)
  return first === last ? first : `${first}:${last}`
}

const rk = new DataView(new ArrayBuffer(8))

export function rkNumber(encoded: number): number {
  let number: number
  if (encoded & 2) number = encoded >> 2
  else {
    rk.setUint32(4, encoded & ~3, true)
    number = rk.getFloat64(0, true)
  }
  return encoded & 1 ? number / 100 : number
}

const errors: Readonly<Partial<Record<number, string>>> = {
  0: '#NULL!',
  7: '#DIV/0!',
  15: '#VALUE!',
  23: '#REF!',
  29: '#NAME?',
  36: '#NUM!',
  42: '#N/A',
  43: '#GETTING_DATA',
}

export function errorText(code: number): string {
  return errors[code] ?? '#UNKNOWN!'
}

export function colorAttrs(payload: Payload, differential = false): string {
  const kind = payload.u8() >>> 1
  const index = payload.u8()
  const tint = payload.i16()
  const red = payload.u8()
  const green = payload.u8()
  const blue = payload.u8()
  const alpha = payload.u8()
  let color: string
  switch (kind) {
    case 0: color = 'auto="1"'; break
    case 1: color = `indexed="${index}"`; break
    case 2:
      color = `rgb="${[alpha, red, green, blue].map((n) => n.toString(16).padStart(2, '0')).join('').toUpperCase()}"`
      break
    case 3: color = `theme="${index}"`; break
    case 4:
      if (differential) return ''
      return payload.fail(`unknown color type ${kind}`)
    default: return payload.fail(`unknown color type ${kind}`)
  }
  return tint ? `${color} tint="${tint / 32767}"` : color
}
