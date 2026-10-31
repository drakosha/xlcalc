/**
 * The set `checked-over` of fields placed from Excel's PivotTable Fields list (`checked.ts`) so that the PivotTable
 * grows over cells below its empty place (rows 21 on): a value, answered Yes and No, a formula, a value with a number
 * format, a number format alone, merged cells, a table; what Excel asks, and what is left under it.
 */
import type { CheckedCase, CheckedSet, Under } from './checked.js'

export const buildCheckedCases: CheckedSet['buildCheckedCases'] = (): CheckedCase[] => {
  const c = (id: string, answer: 'yes' | 'no', under: { cells?: Under[]; merge?: string; table?: string }): CheckedCase => ({ id, data: 'sales', fields: [{ name: 'Region' }, { name: 'Product' }], under, answer, watch: 'A21:C40' })
  return [
    c('value-yes', 'yes', { cells: [{ cell: 'A25', value: 'x' }] }),
    c('value-no', 'no', { cells: [{ cell: 'A25', value: 'x' }] }),
    c('formula-yes', 'yes', { cells: [{ cell: 'A26', formula: '=1+1' }] }),
    c('value-format-yes', 'yes', { cells: [{ cell: 'A27', value: 5, format: '0.000' }] }),
    c('format-yes', 'yes', { cells: [{ cell: 'A27', format: '0.000' }] }),
    c('merged-yes', 'yes', { cells: [{ cell: 'A26', value: 'm' }], merge: 'A26:B26' }),
    c('table-yes', 'yes', { cells: [{ cell: 'A22', value: 'H1' }, { cell: 'B22', value: 'H2' }, { cell: 'A23', value: 1 }, { cell: 'B23', value: 2 }], table: 'A22:B23' }),
  ]
}
