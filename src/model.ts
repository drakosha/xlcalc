/** Адреса A1 и даты как серийные номера Excel 1900 (внутреннее, не контракт). */
import type { CellInput, Scalar } from './contract'

// ---------- адреса ----------

export const MAX_ROWS = 1_048_576
export const MAX_COLS = 16_384

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

const ADDR = /^([A-Z]{1,3})([1-9][0-9]*)$/

/** `B12` → { row: 11, col: 1 } (с нуля). */
export function parseAddress(addr: string): { row: number; col: number } {
  const m = ADDR.exec(addr)
  if (!m) throw new Error(`bad address: ${addr}`)
  return { row: Number(m[2]) - 1, col: colToIndex(m[1]) }
}

export function formatAddress(row: number, col: number): string {
  return `${indexToCol(col)}${row + 1}`
}

/** Порядок ячеек по адресу: строка, потом колонка. */
export function compareAddress(a: string, b: string): number {
  const x = parseAddress(a)
  const y = parseAddress(b)
  return x.row - y.row || x.col - y.col
}

// ---------- даты ----------

const DAY_MS = 86_400_000
/** Серийный номер Excel 1900: 1 = 1900-01-01, с ложным 29.02.1900 (`leapYear1900`). */
const EPOCH_MS = Date.UTC(1899, 11, 30)

export function dateToSerial(iso: string): number {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/.exec(iso)
  if (!m) throw new Error(`bad date: ${iso}`)
  const ms = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] ?? 0), Number(m[5] ?? 0), 0) + Number(m[6] ?? 0) * 1000
  const serial = (ms - EPOCH_MS) / DAY_MS
  // до 1900-03-01 Excel на день меньше (ложный 29.02.1900)
  return serial < 61 ? serial - 1 : serial
}

export function serialToUtcMs(serial: number): number {
  const s = serial < 61 ? serial + 1 : serial
  return EPOCH_MS + Math.round(s * DAY_MS)
}

/** Скалярное значение входной ячейки: даты превращаются в серийный номер. */
export function inputScalar(v: CellInput['value']): Scalar {
  if (v === undefined) return null
  if (v !== null && typeof v === 'object') return dateToSerial(v.date)
  return v
}

/** Момент (мс UTC) в серийный номер с долей суток; часы движка считаются в UTC. */
export function msToSerial(ms: number): number {
  const serial = (ms - EPOCH_MS) / DAY_MS
  return serial < 61 ? serial - 1 : serial
}
