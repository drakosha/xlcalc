/**
 * Month names after a day, the system long date and time (`[$-F800]`, `[$-F400]`), improper fractions of large
 * numbers, numbers under a text format in narrow columns.
 */
import { collector } from './cases.js'

export function buildDisplayCases() {
  const { out, add } = collector()
  for (const f of ['d mmmm yyyy', 'dd mmmm', 'd mmm yyyy', 'mmmm', 'mmmm d', 'd" "mmmm', 'd-mmmm', '[$-F800]dddd, mmmm dd, yyyy', '[$-F400]h:mm:ss AM/PM', '[$-409]d mmmm yyyy'])
    for (const v of [45688.75, 45658.25]) add('names', v, f, { width: 40 })
  for (const v of [99999999, 100000000, 999999999, 1000000000, 1000000001, 2000000000, 123456789.5, 1e9 + 0.5]) for (const f of ['?/?', '??/??', '# ?/?']) add('fraction', v, f, { width: 40 })
  for (const width of [2, 3, 4, 5, 6, 8, 12]) for (const v of [1 / 3, 1234567.891, -0.5]) add('text-number', v, '@', { width })
  for (const [v, f] of [[0, '[>5]0;[<-5]0'], [3, '[>5]0;[<-5]0;'], [-1, '[>5]0'], [0, '[Red][<0]0;[Blue][>0]0'], [3, '[>5]"big";[<-5]"small"']] as const) add('unmatched', v, f, { width: 10 })
  return out
}
