/**
 * en-US: formula text is the invariant syntax itself, as Excel with the English interface shows formulas. Its rules of
 * computation are the engine's own code (`src/locale/rules.ts`), so the file has no regional data.
 */
import type { LocaleFile } from '../locale/file.js'

export const EN_US: LocaleFile = {
  text: {
    tag: 'en-US',
    decimalSeparator: '.',
    listSeparator: ',',
    arrayColumnSeparator: ',',
    arrayRowSeparator: ';',
    booleans: { TRUE: 'TRUE', FALSE: 'FALSE' },
    errors: {},
    functions: {},
  },
}
