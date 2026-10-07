/**
 * Built-in number formats by number (`numFmtId` 0 to 49 without a code, as an xlsx file stores them), each on a
 * positive and a negative number, zero, a fraction, a date with a time and text, in a wide column; with the code
 * Excel reports for each (`NumberFormat`, `NumberFormatLocal`). The same book is shown by Excel en-US and ru-RU.
 */
import { collector } from './cases.js'

const VALUES = [1234.5678, -1234.5678, 0, 0.5, 45688.75, 'abc']

export function buildDisplayCases() {
  const { out, add } = collector()
  for (let id = 0; id <= 49; id++) for (const v of VALUES) add(`id${id}`, v, undefined, { width: 40, numFmtId: id })
  return out
}
