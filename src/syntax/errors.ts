/** Error codes as Excel writes them in formulas and cells in en-US (`#DIV/0!`); data alone, for the locale files too. */
import type { ErrorCode } from '../contract.js'

export const ERROR_TEXT: Record<ErrorCode, string> = {
  DIV0: '#DIV/0!',
  NA: '#N/A',
  NAME: '#NAME?',
  NUM: '#NUM!',
  REF: '#REF!',
  VALUE: '#VALUE!',
  NULL: '#NULL!',
  CYCLE: '#CYCLE!',
  SPILL: '#SPILL!',
  CALC: '#CALC!',
  ERROR: '#ERROR!',
}

/** An error's text in a locale's words (`errors`, by stored text: `#DIV/0!` is `#ДЕЛ/0!` in ru-RU), the stored text where it has none. */
export const errorIn = (code: ErrorCode, errors?: Readonly<Record<string, string>>): string => errors?.[ERROR_TEXT[code]] ?? ERROR_TEXT[code]
