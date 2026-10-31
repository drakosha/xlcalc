/**
 * The functions' words of a PivotTable in the workbook's language, which name a value field (`Sum of Amount`, `Сумма по
 * полю Amount`) and a subtotal, and a blank item: in the root, as an edit of a PivotTable names its value fields and a
 * slicer labels its items without `xlcalc/pivot`.
 */
import { caseKey } from './collate.js'
import type { PivotFunction } from './contract.js'

export const EN_FUNCTIONS: Record<PivotFunction, string> = {
  sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count',
  stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp',
}

/** The functions of a value field, as `PivotFunction` names them. */
export const PIVOT_FUNCTIONS = Object.keys(EN_FUNCTIONS) as PivotFunction[]

// the words two functions share once each: a bundle writes each Cyrillic letter in six characters
const COUNT = 'Количество'
const DEVIATION = 'Стандартное отклонение'
const VARIANCE = 'Дисперсия'

export const RU_FUNCTIONS: Record<PivotFunction, string> = {
  sum: 'Сумма', count: COUNT, average: 'Среднее', max: 'Максимум', min: 'Минимум', product: 'Произведение',
  countNums: `${COUNT} чисел`, stdDev: DEVIATION, stdDevp: `${DEVIATION} (несмещённое)`, var: VARIANCE, varp: `${VARIANCE} (несмещённая)`,
}

/** A value field's name by its function and its source field's name, in the workbook's language (en-US's for a locale without its own). */
export const defaultValueName = (fn: PivotFunction, field: string, locale?: string) => (locale === 'ru-RU' ? `${RU_FUNCTIONS[fn]} по полю ${field}` : `${EN_FUNCTIONS[fn]} of ${field}`)

/**
 * The name Excel gives a value field named `name` when its function changes from `from` to `to` (Value Field Settings,
 * `DataField.Function`), the source field being `field`: a name beginning with the old function's words in the
 * workbook's language, in any case, keeps the rest after the new function's (`sum of qty2` is `Count of qty2`); any other
 * name, its own too or another language's default, is the new function's default name (`Count of Qty`).
 */
export function renamedValueField(name: string | undefined, from: PivotFunction, to: PivotFunction, field: string, locale?: string): string {
  const old = defaultValueName(from, '', locale)
  return name && caseKey(name.slice(0, old.length)) === caseKey(old) ? defaultValueName(to, '', locale) + name.slice(old.length) : defaultValueName(to, field, locale)
}

/** A blank item's label in the workbook's language: `(blank)`, `(пусто)` in ru-RU. */
export const blankWord = (locale?: string) => (locale === 'ru-RU' ? '(пусто)' : '(blank)')
