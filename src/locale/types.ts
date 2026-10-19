/**
 * A locale: its formula text, what Excel translates when it shows a formula in another language, and for a locale a
 * book computes in its rules of computation and its messages. Stored formulas stay invariant, as in xlsx (English
 * function names, `,` between arguments, `.` as the decimal separator); a locale changes how formulas are typed and
 * shown, and a book's values as Excel's regional settings do. en-US is built in, the package's other locales are
 * modules (`xlcalc/locales/ru-RU`) that `registerLocale` loads; formula text in another language is another object
 * of this shape, which `toLocal` and `toInvariant` take without loading it.
 */
import type { Messages } from './messages.js'
import type { RegionalData } from './rules/data.js'
import type { FunctionCategory, Refusal } from '../contract.js'
import type { TableItem } from '../syntax/table.js'

export type { TableItem }

/** A locale (`ruRU` from `xlcalc/locales/ru-RU`): how formulas are written in it and, for a workbook calculated in it, its calculation rules and messages. */
export interface Locale {
  /** BCP 47 tag: `en-US`, `ru-RU`. A locale without messages of its own uses those of the loaded locale with its tag (English for any other tag). */
  readonly tag: string
  /** Decimal separator of number literals: `.` in en-US, `,` in ru-RU. */
  readonly decimalSeparator: string
  /** Separator of function arguments, also the union operator: `,` in en-US, `;` in ru-RU. */
  readonly listSeparator: string
  /** Separator of columns in array constants: `,` in en-US (`{1,2;3,4}`). */
  readonly arrayColumnSeparator: string
  /** Separator of rows in array constants: `;` in en-US. */
  readonly arrayRowSeparator: string
  /** The words for TRUE and FALSE. */
  readonly booleans: { readonly TRUE: string; readonly FALSE: string }
  /** Error literals by their invariant text (`'#N/A': '#Н/Д'`); an error not listed keeps its English text. */
  readonly errors: Readonly<Record<string, string>>
  /** Function names by invariant name (`SUM: 'СУММ'`); a function not listed keeps its English name. */
  readonly functions: Readonly<Record<string, string>>
  /**
   * The item specifiers of structured references by their invariant text (`'#Data': '#Данные'`); without them, or for a
   * specifier not listed, the English one is used. A reference's items are separated by the list separator
   * (`[#Данные];[Qty]` in ru-RU).
   */
  readonly tableItems?: Readonly<Partial<Record<TableItem, string>>>
  /**
   * The words used in argument names in function tooltips (`functionSignature`), by their English word (`number:
   * 'число'`, so `number1` becomes `число1`); a word not listed stays in English. The package's locales keep them, and
   * `functionText`, in separate modules (`xlcalc/function-text/ru-RU`), which a host merges into the locale.
   */
  readonly argumentWords?: Readonly<Record<string, string>>
  /**
   * The text of Excel's Insert Function dialog: the category names (`functionSignature`) and what each function does
   * (`functionDescription`), by its English name; anything not listed stays in English.
   */
  readonly functionText?: { readonly categories: Readonly<Partial<Record<FunctionCategory, string>>>; readonly descriptions: Readonly<Record<string, string>> }
  /**
   * The calculation rules as data: number and date text, names, `TEXT` format codes, text sort order, the code page.
   * Every locale a workbook is calculated in has them, except en-US, whose rules are built into the engine's code; a
   * locale without them parses and displays values as en-US does.
   */
  readonly regional?: RegionalData
  /**
   * Messages for users in the locale's language (`messages.ts`): why a formula does not parse, why a structure operation
   * is rejected, the words of `explain`, what was lost from an xlsx file. A message the locale lacks is in English. They
   * are functions, the only part of a locale that a structured clone cannot carry: a pool's thread imports the package's
   * module for the tag itself.
   */
  readonly messages?: Partial<Messages>
}

/**
 * The result of translating formula text: the text, or a check code with a message in the locale's language and the
 * position in the given text: `syntax`, or `external_ref` for a reference to a workbook not among the given links.
 */
export type Translation = { ok: true; formula: string } | (Refusal<'syntax' | 'external_ref'> & { at: number })
