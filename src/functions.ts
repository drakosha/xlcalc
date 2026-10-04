/**
 * Белый список функций: число аргументов и какие аргументы диапазоны. Список закреплён тестом;
 * функция попадает в него только с тестом против документации Excel и оракула.
 */

export interface FunctionSpec {
  min: number
  max: number
  /** Аргумент с этим номером диапазон (в остальных местах диапазон это операция над массивом). */
  range: (i: number) => boolean
  /** Чётность числа аргументов (SUMIFS нечётное, COUNTIFS чётное). */
  parity?: 'odd' | 'even'
  volatile?: boolean
}

const none = () => false

export const FUNCTIONS: Readonly<Record<string, FunctionSpec>> = {
  SUM: { min: 1, max: 255, range: () => true },
  SUMIF: { min: 2, max: 3, range: (i) => i === 0 || i === 2 },
  SUMIFS: { min: 3, max: 255, range: (i) => i === 0 || i % 2 === 1, parity: 'odd' },
  COUNTIFS: { min: 2, max: 254, range: (i) => i % 2 === 0, parity: 'even' },
  IF: { min: 2, max: 3, range: none },
  IFERROR: { min: 2, max: 2, range: none },
  XLOOKUP: { min: 3, max: 6, range: (i) => i === 1 || i === 2 },
  ROUND: { min: 2, max: 2, range: none },
  EOMONTH: { min: 2, max: 2, range: none },
  DATE: { min: 3, max: 3, range: none },
  TEXT: { min: 2, max: 2, range: none },
  TRUE: { min: 0, max: 0, range: none },
  FALSE: { min: 0, max: 0, range: none },
  TODAY: { min: 0, max: 0, range: none, volatile: true },
  NOW: { min: 0, max: 0, range: none, volatile: true },
  RAND: { min: 0, max: 0, range: none, volatile: true },
  RANDBETWEEN: { min: 2, max: 2, range: none, volatile: true },
}

export const FUNCTION_NAMES: readonly string[] = Object.keys(FUNCTIONS).sort()
