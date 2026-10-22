/** What the writer's parts share: the XML declaration and the SpreadsheetML namespace they start with, colours, and text cut to Excel's limits. */
import type { Issues, Where, XlsxMessage } from '../issues.js'

export const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
export const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

/** A colour (`#RRGGBB`) as a file's opaque ARGB. */
export const rgb = (c: string) => `FF${c.slice(1).toUpperCase()}`

export { MAX_TEXT } from '../../model.js'

/** Text cut to Excel's limit `max`, reported as `textLength` when it was longer; a character of two code units is not split. */
export function clip(text: string, max: number, issues: Issues, message: XlsxMessage, where: Where): string {
  if (text.length <= max) return text
  issues.add('textLength', message, where)
  const high = (text.charCodeAt(max - 1) & 0xfc00) === 0xd800 && (text.charCodeAt(max) & 0xfc00) === 0xdc00
  return text.slice(0, high ? max - 1 : max)
}
