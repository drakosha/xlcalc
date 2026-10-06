/**
 * Сравнение ответа с эталоном: числа с допуском |a−b| ≤ 1e-9·max(1,|a|) (мягкий 1e-6), комплексные числа
 * текстом (эталон с мнимой единицей) по частям с тем же допуском, прочие строки, логические и ошибки (по коду)
 * точно. Несовпадение получает вид.
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
  | 'строка: регистр'
  | 'строка: формат числа'
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
  if (ref.t === 's') {
    const g = got.v as string
    if (g === ref.v || sameComplex(ref.v, g, TOL)) return { ok: true }
    if (g !== ref.v && g.toUpperCase() === ref.v) return { ok: false, kind: 'строка: регистр' }
    return { ok: false, kind: sameNumbers(ref.v, g) ? 'строка: формат числа' : 'строка' }
  }
  if (ref.t === 'b') return got.v === ref.v ? { ok: true } : { ok: false, kind: 'логическое' }
  return { ok: false, kind: 'прочее' }
}

const NUM = '\\d+(?:\\.\\d+)?(?:[eE][+-]?\\d+)?'
/** Комплексное число в записи Excel с мнимой единицей: `3+4i`, `-0.5-2.5j`, `4i`, `-i`, `1E+21+i`. */
const COMPLEX = new RegExp(`^(?:(-?${NUM})([+-])|(-?))(${NUM})?([ij])$`)

function complexParts(s: string): [number, number, string] | null {
  const m = COMPLEX.exec(s)
  if (!m) return null
  const sign = (m[2] ?? m[3]) === '-' ? -1 : 1
  return [m[1] === undefined ? 0 : Number(m[1]), sign * (m[4] === undefined ? 1 : Number(m[4])), m[5]]
}

/**
 * Комплексный ответ (эталон с мнимой единицей) совпадает по частям с допуском и по единице: последняя из 15
 * значащих цифр зависит от библиотеки синусов (Excel на Маке берёт libm Apple, V8 свою). Строка без мнимой
 * единицы сравнивается точно: числа текстом (TEXT, &, BASE с нулями слева) должны совпадать знак в знак.
 */
function sameComplex(ref: string, got: string, tol: number): boolean {
  const a = complexParts(ref)
  const b = complexParts(got)
  return !!a && !!b && a[2] === b[2] && close(a[0], b[0], tol) && close(a[1], b[1], tol)
}

/** Регистр: formulajs отдаёт строчными то, что Excel пишет заглавными (шестнадцатеричные цифры). Числа: строки с одинаковым текстом между числами и равными (с допуском) числами: "1.5+2i" и "1.50000001+2i". */
function sameNumbers(a: string, b: string): boolean {
  const re = /-?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?/g
  const na = a.match(re) ?? []
  const nb = b.match(re) ?? []
  if (!na.length || na.length !== nb.length || a.replace(re, '#') !== b.replace(re, '#')) return false
  return na.every((x, i) => close(Number(x), Number(nb[i]), TOL))
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
