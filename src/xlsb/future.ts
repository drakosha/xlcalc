import { range } from './fields.js'
import type { Formula } from './formula.js'
import type { Payload } from './records.js'

export function ranges(payload: Payload): string[] {
  const count = payload.u32()
  if (count * 16 > payload.remaining) payload.fail(`invalid range count ${count}`)
  const values: string[] = []
  for (let i = 0; i < count; i++) values.push(range(payload))
  return values
}

/** The corner of all the ranges, even when the first range starts farther down or to the right. */
export function rangeOrigin(refs: readonly string[]): { row: number; col: number } {
  let row = Infinity
  let col = Infinity
  for (const ref of refs) {
    const first = parseAddress(ref.split(':', 1)[0])
    row = Math.min(row, first.row)
    col = Math.min(col, first.col)
  }
  return { row, col }
}

interface Future {
  ranges: string[]
  formulas: Formula[]
  rel?: string
}

/** FRTHeader has counted arrays of ranges and formulas in flag order ([MS-XLSB] 2.5.61). */
export function future(payload: Payload): Future {
  const flags = payload.u32()
  const result: Future = { ranges: [], formulas: [] }
  if (flags & 1) {
    const count = payload.u32()
    if (count * 20 > payload.remaining) payload.fail(`invalid future reference count ${count}`)
    for (let i = 0; i < count; i++) {
      payload.u32()
      result.ranges.push(range(payload))
    }
  }
  if (flags & 2) {
    const count = payload.u32()
    if (count * 8 > payload.remaining) payload.fail(`invalid future range count ${count}`)
    for (let i = 0; i < count; i++) {
      payload.u32()
      for (const ref of ranges(payload)) result.ranges.push(ref)
    }
  }
  if (flags & 4) {
    const count = payload.u32()
    if (count * 12 > payload.remaining) payload.fail(`invalid future formula count ${count}`)
    for (let i = 0; i < count; i++) {
      payload.u32()
      const tokenBytes = payload.u32()
      const extraBytes = payload.u32()
      result.formulas.push({ tokens: payload.section(tokenBytes), extra: payload.section(extraBytes) })
    }
  }
  if (flags & 8) result.rel = payload.wide(payload.u16())
  return result
}

export function guid(payload: Payload): string {
  const hex = (value: number, length: number) => value.toString(16).toUpperCase().padStart(length, '0')
  const first = hex(payload.u32(), 8)
  const second = hex(payload.u16(), 4)
  const third = hex(payload.u16(), 4)
  const fourth = hex(payload.u8(), 2) + hex(payload.u8(), 2)
  let last = ''
  for (let i = 0; i < 6; i++) last += hex(payload.u8(), 2)
  return `{${first}-${second}-${third}-${fourth}-${last}}`
}
import { parseAddress } from '../model.js'
