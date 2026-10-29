/**
 * The delimiter of a CSV file and its fields as text. A first line `sep=;` names the delimiter, as Excel takes it, and
 * is no row; otherwise a host's delimiter, tabs for a `.tsv` file, Excel's list separator of the locale, or a guess over
 * the first lines: of `;`, `,` and a tab, the one that splits them most evenly (`guessDelimiter`). Fields are read as
 * RFC 4180 and Excel read them (`delimitedRows`); a line break after the last row is no row. Written, a field is quoted
 * when it holds the delimiter, a quote, a tab or a line break, a quote in it doubled.
 */
import { delimitedRows } from '../delimited.js'

/** The delimiters a guess chooses from, in the order a tie is settled by. */
const CANDIDATES = [';', ',', '\t'] as const

/** The lines a guess reads. */
const SAMPLE = 10

/** The delimiter a first line `sep=X` names and the text after that line, or none. */
export function sepLine(text: string): { delimiter: string; rest: string } | undefined {
  const m = /^sep=([^\r\n])(\r\n|\r|\n|$)/i.exec(text)
  return m ? { delimiter: m[1], rest: text.slice(m[0].length) } : undefined
}

/**
 * The delimiter of the first lines of `text` that splits them most evenly: of those that split a line, the one that
 * gives the most of them its most frequent count of fields; on a tie, `;`, then `,`, then a tab (a number with a
 * decimal comma splits at `,` as evenly as its fields at `;`). None splits a line: undefined.
 */
export function guessDelimiter(text: string): string | undefined {
  let best: string | undefined
  let most = 0
  for (const d of CANDIDATES) {
    const lines = delimitedRows(text, d, SAMPLE).filter((r) => r.length > 1 || r[0] !== '')
    const counts = new Map<number, number>()
    for (const r of lines) if (r.length > 1) counts.set(r.length, (counts.get(r.length) ?? 0) + 1)
    const even = Math.max(0, ...counts.values())
    if (even > most) [best, most] = [d, even]
  }
  return best
}

/** The rows of a file's text, without the empty one a final line break gives. */
export function csvRows(text: string, delimiter: string): string[][] {
  const rows = delimitedRows(text, delimiter)
  const last = rows[rows.length - 1]
  if (last.length === 1 && last[0] === '' && (rows.length > 1 || text === '')) rows.pop()
  return rows
}

/** A field as written: in quotes, a quote in it doubled, when it holds the delimiter, a quote, a tab or a line break. */
export function quoted(field: string, delimiter: string): string {
  return field.includes(delimiter) || /["\t\r\n]/.test(field) ? `"${field.replaceAll('"', '""')}"` : field
}
