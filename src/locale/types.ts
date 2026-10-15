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

/** A locale (`ruRU` of `xlcalc/locales/ru-RU`): its formula text and, for a book that computes in it, its rules of computation and messages. */
export interface Locale {
  /** BCP 47 tag: `en-US`, `ru-RU`. A locale without messages of its own takes those of the loaded locale of its tag (English for another tag). */
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
   * The specifiers of structured references by their invariant text (`'#Data': '#Данные'`); without them, or for one
   * not listed, the English one. Between a reference's items the list separator stands (`[#Данные];[Qty]` in ru-RU).
   */
  readonly tableItems?: Readonly<Partial<Record<TableItem, string>>>
  /**
   * The words of the arguments' names in the tips of functions (`functionSignature`) by their English word (`number:
   * 'число'`, so `number1` is `число1`); a word not listed stays English. The package's locales leave them, and
   * `functionText`, to modules of their own (`xlcalc/function-text/ru-RU`), which a host joins to the locale.
   */
  readonly argumentWords?: Readonly<Record<string, string>>
  /**
   * The words of Excel's Insert Function window: the categories' names (`functionSignature`) and what each function does
   * (`functionDescription`), by its English name; what is not listed stays English.
   */
  readonly functionText?: { readonly categories: Readonly<Partial<Record<FunctionCategory, string>>>; readonly descriptions: Readonly<Record<string, string>> }
  /**
   * The rules of computation as data: number and date text, names, `TEXT`'s codes, the text order, the code page. A
   * locale a book computes in has them, but en-US, whose rules are the engine's own code; a locale without them reads
   * and shows values as en-US does.
   */
  readonly regional?: RegionalData
  /**
   * Messages for people in the locale's language (`messages.ts`): why a formula does not read, why a structure operation
   * is refused, the words of `explain`, what an xlsx file lost. A message the locale lacks is English. They are functions, the one part of a
   * locale a structured clone does not carry: a pool's thread imports the package's module of the tag itself.
   */
  readonly messages?: Partial<Messages>
}

/**
 * The result of translating formula text: the text, or a check code with a message in the locale's language and the
 * position in the given text: `syntax`, or `external_ref` for a reference to a workbook the given links do not hold.
 */
export type Translation = { ok: true; formula: string } | (Refusal<'syntax' | 'external_ref'> & { at: number })
