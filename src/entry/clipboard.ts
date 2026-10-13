/**
 * Text on the clipboard as Excel pastes it (the sets `ui/paste-held`, `rows/keys` and `rows/quotes`): a row at each line
 * break (`\r\n`, `\r` or `\n`), a field at each tab, the rows filled out to a rectangle as wide as the longest, so a
 * short row's missing fields are empty and clear their cells. A line break after the last row is one more row, empty,
 * which clears the cells under the paste. A field begun by a quote is read to its closing quote, tabs and line breaks
 * inside it kept and a doubled quote one quote, and what follows the closing quote up to the field's end joined to it
 * (`"a"b` is `ab`); one never closed runs to the end of the text, and a quote elsewhere is a character. `allot` is given
 * the rectangle's cells before it is made (a session's `limits.writes`).
 */
export function clipboardRows(text: string, allot?: (cells: number) => void): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let start = true
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (start && ch === '"') {
      // a quoted part: to the closing quote, a doubled quote one quote
      for (i++; i < text.length; i++) {
        if (text[i] !== '"') field += text[i]
        else if (text[i + 1] === '"') field += text[i++]
        else break
      }
      start = false
    } else if (ch === '\t') {
      row.push(field)
      field = ''
      start = true
    } else if (ch === '\r' || ch === '\n') {
      if (ch === '\r' && text[i + 1] === '\n') i++
      row.push(field)
      rows.push(row)
      row = []
      field = ''
      start = true
    } else {
      field += ch
      start = false
    }
  }
  row.push(field)
  rows.push(row)
  const width = rows.reduce((n, r) => Math.max(n, r.length), 0)
  allot?.(rows.length * width)
  for (const r of rows) while (r.length < width) r.push('')
  return rows
}
