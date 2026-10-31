/**
 * The functions' words of a PivotTable in the workbook's language, which name a value field (`Sum of Amount`, `Сумма по
 * полю Amount`) and a subtotal: in the root, as an edit of a PivotTable names its value fields without `xlcalc/pivot`.
 */
import type { PivotFunction } from './contract.js'

export const EN_FUNCTIONS: Record<PivotFunction, string> = {
  sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count',
  stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp',
}

// the words two functions share once each: a bundle writes each Cyrillic letter in six characters
const COUNT = 'Количество'
const DEVIATION = 'Стандартное отклонение'
const VARIANCE = 'Дисперсия'

export const RU_FUNCTIONS: Record<PivotFunction, string> = {
  sum: 'Сумма', count: COUNT, average: 'Среднее', max: 'Максимум', min: 'Минимум', product: 'Произведение',
  countNums: `${COUNT} чисел`, stdDev: DEVIATION, stdDevp: `${DEVIATION} (несмещённое)`, var: VARIANCE, varp: `${VARIANCE} (несмещённая)`,
}

/** A value field's name by its function and its source field's name, in the workbook's language (en-US's for a locale without its own). */
export const valueFieldName = (fn: PivotFunction, field: string, locale?: string) => (locale === 'ru-RU' ? `${RU_FUNCTIONS[fn]} по полю ${field}` : `${EN_FUNCTIONS[fn]} of ${field}`)
