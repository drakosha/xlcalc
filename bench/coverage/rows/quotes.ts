/**
 * Quotes in text pasted from the clipboard, after the set `keys` showed that Excel for Windows reads a field begun by a
 * quote to its closing quote, line breaks inside: a tab inside quotes, text after the closing quote, a quote inside a
 * field, quoted fields side by side and in rows, a space before the quote, an empty quoted field, a doubled quote
 * alone. Pasted with Cmd+V on the Mac and through `Worksheet.Paste` on Windows.
 */
import { collector, type RowCase } from './cases.js'

const PASTES: [text: string, more: Partial<RowCase>][] = [
  ['"a\tb"', { fillTo: 1 }],
  ['"a"b', {}],
  ['a"b', {}],
  ['a"b\nc"d', { below: true }],
  ['"a"\t"b"', { fillTo: 1 }],
  ['"a\nb"\n"c"', { below: true }],
  [' "a\nb"', { below: true }],
  ['""', {}],
  ['""""', {}],
  ['"a\r\nb"', { below: true }],
  ['"x\ny"\tz\n1\t2', { fillTo: 1, below: true }],
  ['"=1+\n2"', { below: true }],
]

export function buildRowCases(): RowCase[] {
  const { out, add } = collector()
  for (const [text, more] of PASTES) add('paste', { typed: [[0, text]], via: 'paste', ...more })
  return out
}
