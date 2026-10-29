/**
 * The English text of the messages of CSV issues: the `./csv` subpath's own, as the xlsx issues' are `./xlsx`'s. A
 * locale writes them in its `messages` (ru-RU's), and a locale without its own text of one gives this one.
 */
import { ownText, type CsvMessageParams, type Message } from '../locale/messages.js'
import type { Locale } from '../locale/types.js'

/** A message of a CSV issue. */
export type CsvMessage = Message<keyof CsvMessageParams>

/** The messages of CSV issues in English. */
export const CSV_EN_MESSAGES: { readonly [K in keyof CsvMessageParams]: (p: CsvMessageParams[K]) => string } = {
  csvFormulaText: () => 'fields Excel reads as formulas are kept as text (readCsv takes them as formulas with formulas: true)',
  csvFormulaUnparsed: () => 'fields whose formula does not parse are kept as text',
  csvTextCut: () => "a field longer than Excel's limit of 32,767 characters is truncated",
  csvOutside: () => 'rows past row 1,048,576 and fields past column XFD are left out, as Excel leaves them',
  csvSheets: (p) => `the file holds the sheet ${JSON.stringify(p.sheet)} alone: the other sheets are left out`,
  csvCharsLost: (p) => `characters the code page ${p.page} lacks are written as ?`,
}

/** An issue's message in a locale (en-US without one): the locale's own text, else the English one. */
export function csvText(locale: string | Locale | undefined, m: CsvMessage): string {
  return ownText(locale, m) ?? (CSV_EN_MESSAGES[m.key] as (p: CsvMessage) => string)(m)
}
