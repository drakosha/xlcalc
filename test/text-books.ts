/** What the sheets of the tests of typed and pasted text share. */
import type { SheetInput } from '../src/contract.js'

/** A table T over E1:E2 (its column Qty, one row), column D in percent, whole numbers over 0 in G1:G9 with a warning. */
export const typingParts = (): Pick<SheetInput, 'columns' | 'tables' | 'validations'> => ({
  columns: { D: { numberFormat: '0%' } },
  tables: [{ name: 'T', range: 'E1:E2', columns: [{ name: 'Qty' }] }],
  validations: [{ range: 'G1:G9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'warning', error: 'Positive only' }],
})
