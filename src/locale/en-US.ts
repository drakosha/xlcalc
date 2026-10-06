/** en-US: the invariant syntax itself, as Excel with the English interface shows formulas. */
import type { Locale } from './types.js'

export const enUS: Locale = {
  tag: 'en-US',
  decimalSeparator: '.',
  listSeparator: ',',
  arrayColumnSeparator: ',',
  arrayRowSeparator: ';',
  booleans: { TRUE: 'TRUE', FALSE: 'FALSE' },
  errors: {},
  functions: {},
}
