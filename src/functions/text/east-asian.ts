/**
 * Functions of East Asian text, as Excel computes them in the workbook's locales, which use no double-byte characters
 * (Excel for Mac and Windows in en-US alike, the set `east-asian-held`): ASC and DBCS give their text as it is; JIS,
 * the name Japanese Excel gives DBCS, is not a function there (#NAME?); PHONETIC, the readings (furigana) a file
 * stores with Japanese text, is #N/A for any reference, as in Excel for Windows (Excel for Mac answers #NUM!).
 */
import type { FnDef } from '../../args.js'
import { E } from '../../values.js'

export const functions: Record<string, FnDef> = {
  ASC: { args: 's', fn: (s: string) => s },
  DBCS: { args: 's', fn: (s: string) => s },
  JIS: { args: 'a', fn: () => E.NAME },
  PHONETIC: { args: 'R', areas: true, fn: () => E.NA },
}
