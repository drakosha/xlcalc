/**
 * A locale of formula text: what Excel translates when it shows a formula in another language. Stored
 * formulas stay invariant, as in xlsx (English function names, `,` between arguments, `.` as the decimal
 * separator); a locale changes how formulas are typed and shown. A locale is plain data: a new locale is a
 * new object of this shape, no engine code.
 */
export interface Locale {
  /** BCP 47 tag: `en-US`, `ru-RU`. */
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
}

/** The result of translating formula text: the text, or a check code with the position in the given text. */
export type Translation = { ok: true; formula: string } | { ok: false; code: 'syntax'; message: string; at: number }
