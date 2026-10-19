/** The checks of input's cell addresses and ranges in A1 style, within the sheet; the module needs nothing of the engine. */
import { InputError, what } from '../input-error.js'
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { parseRange, type Rect } from '../structure/area.js'
import { checkString } from './basic.js'

const ADDRESS = /^([A-Z]{1,3})([1-9][0-9]*)$/
const MAX_ADDRESS = 'XFD1048576'

/** A cell address in A1 style within the Excel grid. */
export function checkAddress(addr: unknown, path: string): string {
  if (typeof addr !== 'string') throw new InputError(path, `expected a cell address such as "B12", got ${what(addr)}`)
  const m = ADDRESS.exec(addr)
  if (!m) {
    const hint = ADDRESS.test(addr.toUpperCase()) ? ` (column letters must be upper-case: "${addr.toUpperCase()}")` : ''
    throw new InputError(path, `invalid cell address ${JSON.stringify(addr)}: expected A1 style such as "B12"${hint}`)
  }
  if (m[1].length < 3 && m[2].length < 7) return addr
  let col = 0
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64)
  if (col > MAX_COLS || Number(m[2]) > MAX_ROWS) throw new InputError(path, `cell address ${JSON.stringify(addr)} is outside the sheet (the last cell is ${MAX_ADDRESS})`)
  return addr
}

/**
 * A range in A1 style within the sheet (`B2`, `B2:D9`, `B:D`, `2:9`), as its rectangle. `typed`: a value other than a
 * string is refused as one (`expected a string`); without it, as a range, as a sheet's looks and a rule's ranges refuse it.
 */
export function checkRange(v: unknown, path: string, typed = true): Rect {
  if (typed) checkString(v, path)
  const x = typeof v === 'string' ? parseRange(v) : null
  if (!x) throw new InputError(path, `expected a range such as "B2", "B2:D9", "B:D" or "2:9" within the sheet, got ${JSON.stringify(v)}`)
  return x
}
