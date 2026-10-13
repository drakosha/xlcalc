/** Plain-text export quotes fields as Excel does, preserving characters inside a field. */
import { describe, expect, it } from 'vitest'
import * as text from '../src/text.js'
import { InputError } from '../src/index.js'

it('exports CRLF-separated rows with a terminal CRLF on Windows and none on Mac', () => {
  const rows = [['a', 'b'], ['c', 'd']]
  expect(text.formatClipboard(rows)).toBe('a\tb\r\nc\td\r\n')
  expect(text.formatClipboard(rows, { trailingNewline: false })).toBe('a\tb\r\nc\td')
})

it('quotes controls, doubling quotes in such a field, but leaves a quote alone as text', () => {
  expect(text.formatClipboard([['a\tb', 'a\nb', 'a"b', 'a\r"b']])).toBe('"a\tb"\t"a\nb"\ta"b\t"a\r""b"\r\n')
})

it('handles an empty grid and refuses invalid rows, fields and options', () => {
  expect(text.formatClipboard([])).toBe('')
  for (const rows of [null, [1], [[2]], [['a'], ['b', 'c']], [[]]]) expect(() => text.formatClipboard(rows as never)).toThrow(InputError)
  expect(() => text.formatClipboard([['a']], { unknown: true } as never)).toThrow(InputError)
  expect(() => text.formatClipboard([['a']], { trailingNewline: 1 } as never)).toThrow(InputError)
})
