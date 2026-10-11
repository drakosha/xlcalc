/**
 * Text on the clipboard as Excel pastes it (the set `ui/paste-held`, Excel for Mac and Windows alike): a row at each
 * line break (`\r\n`, `\r` or `\n`), a field at each tab, the rows filled out to a rectangle as wide as the longest, so a
 * short row's missing fields are empty and clear their cells. A line break after the last row is one more row, empty,
 * which clears the cells under the paste. Quotes are not read: a field is the text between its tabs.
 */
export function clipboardRows(text: string): string[][] {
  const rows = text.split(/\r\n|\r|\n/).map((line) => line.split('\t'))
  const width = rows.reduce((n, row) => Math.max(n, row.length), 0)
  for (const row of rows) while (row.length < width) row.push('')
  return rows
}
