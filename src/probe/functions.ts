/**
 * Белый список функций замера и общие для обоих вариантов волатильные функции от момента и зерна.
 * Окончательный список закрепляется тестом в шаге 2.
 */
import { dateToSerial, type Clock } from './model'

/** Закрытый список семантики шага 1. */
export const SEMANTICS = ['SUM', 'SUMIFS', 'COUNTIFS', 'IF', 'IFERROR', 'XLOOKUP', 'ROUND', 'EOMONTH', 'DATE', 'TEXT'] as const

export const VOLATILE = ['TODAY', 'NOW', 'RAND', 'RANDBETWEEN'] as const

/** Белый список проверки формул в замере: семантика, волатильные и функции из книг пробы. */
export const WHITELIST: ReadonlySet<string> = new Set<string>([...SEMANTICS, ...VOLATILE, 'SUMIF', 'TRUE', 'FALSE'])

/** Серийный момент часов (дробная часть это время суток, UTC). */
export const nowSerial = (clock: Clock) => dateToSerial(clock.now)

/** Псевдослучайное [0, 1) от зерна и адреса ячейки: один и тот же для одного `clock`. */
export function randAt(seed: number, sheet: string, row: number, col: number): number {
  // FNV-1a по строке ключа, потом splitmix32
  let h = 0x811c9dc5 ^ (seed >>> 0)
  const key = `${sheet}\u0000${row}\u0000${col}`
  for (let i = 0; i < key.length; i++) {
    h ^= key.charCodeAt(i)
    h = Math.imul(h, 0x01000193)
  }
  h = (h + 0x9e3779b9) | 0
  h = Math.imul(h ^ (h >>> 16), 0x85ebca6b)
  h = Math.imul(h ^ (h >>> 13), 0xc2b2ae35)
  h ^= h >>> 16
  return (h >>> 0) / 4_294_967_296
}

export function randBetween(seed: number, sheet: string, row: number, col: number, lo: number, hi: number): number {
  const a = Math.ceil(lo)
  const b = Math.floor(hi)
  return a + Math.floor(randAt(seed, sheet, row, col) * (b - a + 1))
}
