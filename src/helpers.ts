/**
 * Helpers without a book that answer by the engine's own rules: a sheet name as Excel takes it (the book check's), an
 * error's text as a cell shows it (the locale's words), and values compared and written as the values hash compares and
 * writes them (`canon.ts`).
 */
import type { CellValue, ErrorCode, NameOptions, SheetNameCheck, SheetNameCode } from './contract.js'
import { canonValue as canonText, outNumber } from './canon.js'
import { checkCellValue, checkNameOptions, sheetNameProblem } from './input.js'
import { InputError } from './input-error.js'
import { checkString } from './input/basic.js'
import { optionalLocale } from './locale/load.js'
import { say, type MessageKey } from './locale/messages.js'
import { EN_US } from './locales/en-US.js'
import type { Locale } from './locale/types.js'
import { ERROR_TEXT, errorIn } from './syntax/errors.js'

/** The code of each reason the book check gives for a sheet name. */
const SHEET_CODES: Partial<Record<MessageKey, SheetNameCode>> = { sheetNameEmpty: 'empty', sheetNameLong: 'too_long', sheetNameChar: 'character', sheetNameApostrophe: 'apostrophe', sheetNameHistory: 'reserved' }

/**
 * Whether Excel accepts a sheet name, by the rules a workbook's sheets are checked against: 1 to 31 characters, none of
 * `: \ / ? * [ ]`, no apostrophe at either end, not `History` in any case. A rejection's message is in `opts.locale`'s
 * language, English without one; `at` is the position of the character Excel rejects.
 */
export function validateSheetName(name: string, opts?: NameOptions): SheetNameCheck {
  const o = checkNameOptions(opts)
  const problem = sheetNameProblem(checkString(name, 'name'))
  if (!problem) return { ok: true }
  const at = problem.key === 'sheetNameChar' ? name.indexOf(problem.char) : problem.key === 'sheetNameApostrophe' ? (name.startsWith("'") ? 0 : name.length - 1) : undefined
  return { ok: false, code: SHEET_CODES[problem.key]!, message: say(o.locale, problem), ...(at === undefined ? {} : { at }) }
}

/** An error as a cell displays it in `locale` (a loaded locale's tag or a locale object, en-US without one): `#DIV/0!`, `#ДЕЛ/0!` in ru-RU. */
export function errorText(code: ErrorCode, locale?: string | Locale): string {
  if (!Object.hasOwn(ERROR_TEXT, code)) throw new InputError('code', `expected an error code (${Object.keys(ERROR_TEXT).join(', ')}), got ${JSON.stringify(code)}`)
  return errorIn(code, (optionalLocale(locale, 'locale') ?? EN_US).errors)
}

/** A value checked at the boundary; an absent one (`undefined`) is an empty cell. */
const valueOf = (v: CellValue | undefined, path: string): CellValue => (v === undefined ? null : checkCellValue(v, path))

/**
 * A value's canonical text, as its line in the values hash holds it: `n 1234.5` (15 significant digits, `-0` as 0),
 * `s "text"`, `b TRUE`, `e DIV0`, `_` for an empty cell.
 */
export const canonValue = (value: CellValue): string => canonText(valueOf(value, 'value'))

/**
 * Whether two values are the same for the values hash: numbers equal to 15 significant digits (`0.1 + 0.2` is `0.3`,
 * `-0` is 0), errors by their code; an absent value (`undefined`) is an empty cell, like null.
 */
export function sameValue(a: CellValue | undefined, b: CellValue | undefined): boolean {
  const [x, y] = [valueOf(a, 'a'), valueOf(b, 'b')]
  return typeof x === 'number' && typeof y === 'number' ? outNumber(x) === outNumber(y) : canonText(x) === canonText(y)
}
