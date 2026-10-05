/**
 * Сравнение ответа с эталоном: числа с допуском |a−b| ≤ 1e-9·max(1,|a|) (мягкий 1e-6), строки,
 * логические и ошибки (по коду) точно. Несовпадение получает вид.
 */
import type { Val } from './formulajs.js'

export const TOL = 1e-9
export const SOFT = 1e-6

export type Kind =
  | 'число'
  | 'ошибка вместо значения'
  | 'значение вместо ошибки'
  | 'другой код ошибки'
  | 'другой тип'
  | 'строка'
  | 'логическое'
  | 'JS Date'
  | 'NaN или бесконечность'
  | 'исключение'
  | 'зависает'
  | 'прочее'

export interface Verdict {
  ok: boolean
  /** Совпало только с мягким допуском 1e-6. */
  soft?: boolean
  kind?: Kind
}

const close = (a: number, b: number, tol: number) => Math.abs(a - b) <= tol * Math.max(1, Math.abs(a))

/** `got` против эталона `ref`; `threw`: ответ #VALUE! получен из исключения formulajs. */
export function compare(ref: Val, got: Val, threw?: boolean): Verdict {
  if (got.t === 'other' && (got.v === 'hang' || got.v === 'crash')) return { ok: false, kind: 'зависает' }
  if (got.t === 'throw') return { ok: false, kind: 'исключение' }
  if (got.t === 'nan') return { ok: false, kind: 'NaN или бесконечность' }
  if (got.t === 'date') {
    // голая formulajs: дата объектом JS; значение сравнивается отдельно в отчёте
    return { ok: false, kind: 'JS Date' }
  }
  if (got.t === 'other') return { ok: false, kind: 'прочее' }
  if (ref.t === 'e') {
    if (got.t !== 'e') return { ok: false, kind: 'значение вместо ошибки' }
    return got.v === ref.v ? { ok: true } : { ok: false, kind: threw ? 'исключение' : 'другой код ошибки' }
  }
  if (got.t === 'e') return { ok: false, kind: threw ? 'исключение' : 'ошибка вместо значения' }
  if (ref.t !== got.t) return { ok: false, kind: 'другой тип' }
  if (ref.t === 'n') {
    const a = ref.v
    const b = got.v as number
    if (close(a, b, TOL)) return { ok: true }
    if (close(a, b, SOFT)) return { ok: false, soft: true, kind: 'число' }
    return { ok: false, kind: 'число' }
  }
  if (ref.t === 's') return got.v === ref.v ? { ok: true } : { ok: false, kind: 'строка' }
  if (ref.t === 'b') return got.v === ref.v ? { ok: true } : { ok: false, kind: 'логическое' }
  return { ok: false, kind: 'прочее' }
}

/** Значение для таблиц отчёта. */
export function show(v: Val | undefined): string {
  if (!v) return 'нет'
  switch (v.t) {
    case 'n':
      return String(Number(v.v.toPrecision(12)))
    case 's':
      return JSON.stringify(v.v)
    case 'b':
      return v.v ? 'TRUE' : 'FALSE'
    case 'e':
      return `#${v.v}`
    case 'date':
      return `Date(${Number(v.v.toPrecision(12))})`
    default:
      return `${v.t}:${v.v}`
  }
}
