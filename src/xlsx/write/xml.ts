/** What the writer's parts share: the XML declaration and the SpreadsheetML namespace they start with, colours, and text cut to Excel's limits. */
import type { Issues, Where } from '../issues.js'

export const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
export const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

/** A colour (`#RRGGBB`) as a file's opaque ARGB. */
export const rgb = (c: string) => `FF${c.slice(1).toUpperCase()}`

export { MAX_TEXT } from '../../model.js'

/** Text cut to Excel's limit `max`, reported as `textLength` when it was longer. */
export function clip(text: string, max: number, issues: Issues, message: string, where: Where): string {
  if (text.length <= max) return text
  issues.add('textLength', message, where)
  return text.slice(0, max)
}
