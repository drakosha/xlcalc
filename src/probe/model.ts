/**
 * Модель входа книги для замера движков (шаг 1 трека «книга», `docs/design-workbook.md`). Это не
 * контракт: `contract.ts` появится в шаге 2. Общая для обоих вариантов (H и C).
 */

export type Scalar = number | string | boolean | null
export type JsonValue = Scalar | JsonValue[] | { [k: string]: JsonValue }

export interface CellInput {
  value?: Scalar | { date: string }
  /** Текст формулы без `=`. */
  formula?: string
  format?: string
  meta?: JsonValue
}

export interface SheetInput {
  name: string
  cells: Record<string, CellInput>
}

export interface NameInput {
  name: string
  /** Ссылка с листом, например `'Data'!$D$2:$D$3401` или `'Data'!$D:$D`. */
  ref: string
  /** Лист, на котором имя видно; без него имя глобальное. */
  scope?: string
}

export interface BookInput {
  sheets: SheetInput[]
  names?: NameInput[]
}

export type ErrorCode = 'DIV0' | 'NA' | 'NAME' | 'NUM' | 'REF' | 'VALUE' | 'CYCLE' | 'SPILL' | 'ERROR'

/** Значение ячейки после расчёта: скаляр, пусто (`null`) или ошибка. */
export type CellValue = number | string | boolean | null | { error: ErrorCode }

export interface Clock {
  /** Момент ISO (`2026-10-04T12:00:00Z`). */
  now: string
  seed: number
}

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
