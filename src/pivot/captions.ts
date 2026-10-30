/**
 * The words a PivotTable shows in its cells in the workbook's language (`Grand Total`, `Row Labels`, `(blank)`, `Sum of
 * Amount`): en-US, and ru-RU as Excel for Windows ru-RU writes them. A PivotTable's own captions replace them.
 */
import type { PivotFunction } from '../contract.js'

export interface PivotWords {
  grandTotal: string
  rowHeader: string
  columnHeader: string
  values: string
  /** A subtotal's label after its item's (`East Total`). */
  total: string
  blank: string
  all: string
  multiple: string
  /** A value field's name by its function and the field's (`Sum of Amount`). */
  valueName: (fn: PivotFunction, field: string) => string
  /** A subtotal's function after its item's label (`East Count`), and before a value field's name for a grand total (`Total Sum of Amount`). */
  functions: Record<PivotFunction, string>
}

const EN_FUNCTIONS: Record<PivotFunction, string> = {
  sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count',
  stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp',
}

const RU_FUNCTIONS: Record<PivotFunction, string> = {
  sum: 'Сумма', count: 'Количество', average: 'Среднее', max: 'Максимум', min: 'Минимум', product: 'Произведение',
  countNums: 'Количество чисел', stdDev: 'Стандартное отклонение', stdDevp: 'Стандартное отклонение (несмещённое)',
  var: 'Дисперсия', varp: 'Дисперсия (несмещённая)',
}

const WORDS: Record<string, PivotWords> = {
  'en-US': {
    grandTotal: 'Grand Total', rowHeader: 'Row Labels', columnHeader: 'Column Labels', values: 'Values', total: 'Total',
    blank: '(blank)', all: '(All)', multiple: '(Multiple Items)', functions: EN_FUNCTIONS,
    valueName: (fn, field) => `${EN_FUNCTIONS[fn]} of ${field}`,
  },
  'ru-RU': {
    grandTotal: 'Общий итог', rowHeader: 'Названия строк', columnHeader: 'Названия столбцов', values: 'Значения', total: 'Итог',
    blank: '(пусто)', all: '(Все)', multiple: '(несколько элементов)', functions: RU_FUNCTIONS,
    valueName: (fn, field) => `${RU_FUNCTIONS[fn]} по полю ${field}`,
  },
}

/** The words of a workbook's locale, en-US's for a locale without its own. */
export const wordsOf = (locale: string | undefined): PivotWords => WORDS[locale ?? 'en-US'] ?? WORDS['en-US']
