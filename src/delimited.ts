/**
 * Delimited text as rows of fields, as Excel reads text on the clipboard (`entry/clipboard.ts`, tabs) and a CSV file
 * (`csv/`): a row at each line break (`\r\n`, `\r` or `\n`), a field at each delimiter. A field begun by a quote is read
 * to its closing quote, delimiters and line breaks inside it kept and a doubled quote one quote, and what follows the
 * closing quote up to the field's end joined to it (`"a"b` is `ab`); one never closed runs to the end of the text, and a
 * quote elsewhere is a character. A line break after the last row gives one more row, empty.
 */

const QUOTE = 34
const CR = 13
const LF = 10

/** The end of a field's unquoted part from `i`: the next delimiter (`d`, a char code) or line break, or the text's end. */
function fieldEnd(text: string, i: number, d: number): number {
  for (; i < text.length; i++) {
    const c = text.charCodeAt(i)
    if (c === d || c === LF || c === CR) return i
  }
  return i
}

/** The rows of `text` split at `delimiter` (one character), the first `max` of them. */
export function delimitedRows(text: string, delimiter: string, max = Infinity): string[][] {
  const d = delimiter.charCodeAt(0)
  const rows: string[][] = []
  let row: string[] = []
  for (let i = 0; ; ) {
    let field = ''
    if (text.charCodeAt(i) === QUOTE) {
      // the quoted part: to the closing quote, a doubled quote one quote
      for (let from = i + 1; ; ) {
        const q = text.indexOf('"', from)
        if (q < 0) {
          field += text.slice(from)
          i = text.length
          break
        }
        if (text.charCodeAt(q + 1) === QUOTE) {
          field += text.slice(from, q + 1)
          from = q + 2
          continue
        }
        field += text.slice(from, q)
        i = q + 1
        break
      }
    }
    const end = fieldEnd(text, i, d)
    row.push(field + text.slice(i, end))
    i = end
    if (i >= text.length) break
    if (text.charCodeAt(i) === d) {
      i++
      continue
    }
    i += text.charCodeAt(i) === CR && text.charCodeAt(i + 1) === LF ? 2 : 1
    rows.push(row)
    row = []
    if (rows.length >= max) return rows
  }
  rows.push(row)
  return rows
}
