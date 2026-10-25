/** A1 addresses and dates as Excel 1900 serial numbers (internal, not part of the contract). */
import type { CellInput, Scalar } from './contract.js'

// ---------- addresses ----------

/** The number of rows on a sheet, as in Excel: 1,048,576. */
export const MAX_ROWS = 1_048_576
/** The number of columns on a sheet, as in Excel: 16,384, A to XFD. */
export const MAX_COLS = 16_384
/** Excel's limit on the length of a text value: a cell's, a formula's text, a note's. */
export const MAX_TEXT = 32_767
/** The smallest positive normal double: Excel has no smaller numbers (subnormal results are 0). */
export const MIN_NORMAL = 2.2250738585072014e-308

export function colToIndex(letters: string): number {
  let n = 0
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

export function indexToCol(i: number): string {
  let s = ''
  let n = i + 1
  while (n > 0) {
    const r = (n - 1) % 26
    s = String.fromCharCode(65 + r) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

/** A cell address in A1 style: one to three upper-case letters and a row, both within the sheet when checked. */
export const ADDR = /^([A-Z]{1,3})([1-9][0-9]*)$/

/** `B12` → { row: 11, col: 1 } (zero-based). */
export function parseAddress(addr: string): { row: number; col: number } {
  const m = ADDR.exec(addr)
  if (!m) throw new Error(`invalid address: ${addr}`)
  return { row: Number(m[2]) - 1, col: colToIndex(m[1]) }
}

/** Column letters by index, made once each: loops over a sheet's cells write the same columns over and over. */
const LETTERS: string[] = []

export function formatAddress(row: number, col: number): string {
  return (LETTERS[col] ??= indexToCol(col)) + (row + 1)
}

/** A valid address's zero-based row and column, read into `out` without a pattern or an allocation: for loops over a book's cells. */
export function readAddress(addr: string, out: { r: number; c: number }): void {
  let i = 0
  let col = 0
  for (let ch = addr.charCodeAt(0); ch >= 65; ch = addr.charCodeAt(++i)) col = col * 26 + ch - 64
  let row = 0
  for (; i < addr.length; i++) row = row * 10 + addr.charCodeAt(i) - 48
  out.r = row - 1
  out.c = col - 1
}

/** Cell order by address: row, then column. */
export function compareAddress(a: string, b: string): number {
  const x = parseAddress(a)
  const y = parseAddress(b)
  return x.row - y.row || x.col - y.col
}

// ---------- numbers ----------

/** Decimal number text: `12`, `-1.5`, `.5`, `1E-7`. */
export const NUMBER_TEXT = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/

/**
 * Decimal number text (matching `NUMBER_TEXT`) as Excel reads it, in a formula or in text: digits past the
 * 15th significant one are dropped, not rounded (123456789012345678 is 123456789012345000). `shift` moves the decimal
 * point as percent does, in the text: `1.1` with -2 is 0.011, as Excel reads `1.1%`, where 1.1/100 is 0.011000000000000001.
 */
export function readNumber(text: string, shift = 0): number {
  const m = /^([+-]?)(\d*)\.?(\d*)(?:[eE]([+-]?\d+))?$/.exec(text)!
  const digits = m[2] + m[3]
  const first = digits.search(/[1-9]/)
  if ((first < 0 || digits.length - first <= 15) && !shift) return Number(text)
  const kept = first < 0 ? digits : digits.slice(0, first + 15).padEnd(digits.length, '0')
  return Number(`${m[1]}${kept.slice(0, m[2].length)}.${kept.slice(m[2].length)}e${Number(m[4] ?? 0) + shift}`)
}

// ---------- dates ----------

export const DAY_MS = 86_400_000
/** Excel 1900 serial number: 1 = 1900-01-01, with the phantom 1900-02-29 (`leapYear1900`). */
export const EPOCH_MS = /* @__PURE__ */ Date.UTC(1899, 11, 30)

/** Days from 1899-12-30 as Excel's serial number: before 1900-03-01 Excel is one day behind (the phantom 1900-02-29). */
export const fromUtcDays = (days: number) => (days < 61 ? days - 1 : days)

/**
 * The serial number of an ISO date and time without a time zone (a `Z` or an offset after it, as xlsx may write, is
 * passed over); a date or time that is not one, which `Date.UTC` would roll into another (`2025-02-30`, `T25:00`),
 * throws, but Excel's day 0 (`1900-01-00`) and its phantom 1900-02-29.
 */
export function dateToSerial(iso: string): number {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?(?:Z|[+-]\d{2}:?\d{2})?$/.exec(iso)
  if (!m) throw new Error(`invalid date: ${iso}`)
  const [y, mo, d, h, min, s] = [Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] ?? 0), Number(m[5] ?? 0), Number(m[6] ?? 0)]
  const at = new Date(Date.UTC(y, mo, d, h, min))
  const excel = y === 1900 && ((mo === 0 && d === 0) || (mo === 1 && d === 29))
  if (!(s < 60 && at.getUTCHours() === h && at.getUTCMinutes() === min && (excel || (at.getUTCFullYear() === y && at.getUTCMonth() === mo && at.getUTCDate() === d)))) throw new Error(`invalid date: ${iso}`)
  const days = (at.getTime() + s * 1000 - EPOCH_MS) / DAY_MS
  // the phantom 1900-02-29 is serial 60 (Date.UTC rolls it to 1900-03-01)
  return y === 1900 && mo === 1 && d === 29 ? days - 1 : fromUtcDays(days)
}

/** The moment (UTC ms) of a serial number, the inverse of `msToSerial`. */
export function serialToUtcMs(serial: number): number {
  const s = serial < 61 ? serial + 1 : serial
  return EPOCH_MS + Math.round(s * DAY_MS)
}

/** Scalar value of an input cell: dates become serial numbers. */
export function inputScalar(v: CellInput['value']): Scalar {
  if (v === undefined) return null
  if (v !== null && typeof v === 'object') return dateToSerial(v.date)
  return v
}

/** A typed date in the book counts from its epoch rather than from the 1900 system's day 0. */
export function bookScalar(v: CellInput['value'], epoch: number): Scalar {
  const x = inputScalar(v)
  return epoch && typeof v === 'object' && v !== null ? (x as number) - epoch : x
}

/** A moment (UTC ms) as a serial number with a day fraction; the engine clock runs in UTC. */
export function msToSerial(ms: number): number {
  return fromUtcDays((ms - EPOCH_MS) / DAY_MS)
}

/** An address's place in the order of `sortAddresses`: row, then column (`row * MAX_COLS + col`). */
export function addressOrder(a: string): number {
  // an address of a row read exactly by characters (15 digits at most) is read without a match and an object
  if (!SHORT_ADDR.test(a)) {
    const { row, col } = parseAddress(a)
    return row * MAX_COLS + col
  }
  readAddress(a, at)
  return at.r * MAX_COLS + at.c
}
const SHORT_ADDR = /^[A-Z]{1,3}[1-9]\d{0,14}$/
const at = { r: 0, c: 0 }

/** A cell's place as one number, which orders places as addresses: by sheet, row, then column; no text is made for it. */
export const placeOf = (s: number, r: number, c: number) => (s * MAX_ROWS + r) * MAX_COLS + c

/** The sheet, row and column of a place (`placeOf`). */
export function placeParts(n: number): [number, number, number] {
  const c = n % MAX_COLS
  const rest = (n - c) / MAX_COLS
  const r = rest % MAX_ROWS
  return [(rest - r) / MAX_ROWS, r, c]
}

/** Addresses in order (row, then column): each address is parsed once, not per comparison. */
export function sortAddresses(addrs: Iterable<string>): string[] {
  return [...addrs]
    .map((a) => [addressOrder(a), a] as const)
    .sort((x, y) => x[0] - y[0])
    .map((x) => x[1])
}
