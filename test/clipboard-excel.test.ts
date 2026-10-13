/** Exact copy text recorded through COM and AppleScript, without a screen. */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { clipboardCases } from '../bench/coverage/clipboard/cases.js'
import { formatClipboard } from '../src/text.js'

const stored = (path: string) => JSON.parse(readFileSync(new URL(path, import.meta.url), 'utf8')) as { answers: Record<string, { text: string }> }
describe.each(['en-US', 'ru-RU'])('clipboard, %s', (tag) => {
  for (const platform of ['Mac', 'Windows']) {
    const path = platform === 'Mac' ? `./excel/clipboard/copy-${tag}.json` : `./excel/win/clipboard/${tag === 'ru-RU' ? 'ru-RU/' : ''}copy-held.json`
    const answers = stored(path).answers
    for (const c of clipboardCases(tag)) it(`${platform}: ${c.id}`, () => {
      // The host supplies visible rows and columns and unclipped formatted text, not narrow-cell hashes.
      let rows = c.rows
      if (c.id === 'hidden-row') rows = [rows[0], rows[2]]
      if (c.id === 'hidden-column') rows = rows.map((r) => [r[0]])
      if (c.id === 'empty') rows = [['']]
      if (c.id === 'narrow') rows = [['2026-03-15'], [platform === 'Windows' && tag === 'ru-RU' ? '1,23457E+11' : '1.23457E+11'], [platform === 'Windows' && tag === 'ru-RU' ? '1,23456789' : '1.23456789']]
      expect(formatClipboard(rows, { trailingNewline: platform === 'Windows' })).toBe(answers[c.id].text)
    })
  }
})

it('pins the lack of a reversible plain-text format for arbitrary quoted text and CRLF', () => {
  expect(formatClipboard([['"first'], ['last"']])).toBe(formatClipboard([['first\r\nlast']]))
  const { answers } = stored('./excel/win/clipboard/ru-RU/quotes-held.json') as unknown as { answers: Record<string, [string, unknown]> }
  expect(answers['quote-only'][1]).toEqual([['"abc"', null], ['"first', null], ['last"', null], ['sentinel', null]])
  expect(answers['quoted-controls'][1]).toEqual([['a\tb', 'a\nb'], ['a\n"b"', 'end'], [null, null], ['sentinel', null]])
})
