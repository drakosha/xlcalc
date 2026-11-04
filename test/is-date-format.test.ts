/**
 * Whether a number format shows a date or a time (`isDateFormat`), and the codes of digit placeholders with a date's bare
 * letter, which Excel for Windows refuses a file for (`bench/coverage/display/digits-date-letters.ts`).
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { isDateFormat } from '../src/typed-entry.js'

describe('isDateFormat', () => {
  it.each(['dd.mm.yyyy', '[h]:mm', 'yyyy-mm-dd hh:mm', 'mm:ss', 'mm:ss.00', '[h]:mm:ss.000', '[s].00', 'h AM/PM', '[$-409]d mmm', '0;dd.mm', 14, 22, 46])('%s shows a date or a time', (f) => expect(isDateFormat(f)).toBe(true))
  it.each(['General', '0.00', '#,##0 "m"', '0 samples', '#,##0 m', '"date" 0', '\\d0', '@', '0%', '[Red]0', '0.0 h', '.00 s', 'd 0', 'h:mm 0', 0, 4])('%s does not', (f) => expect(isDateFormat(f)).toBe(false))

  it('refuses what is no format', () => expect(() => isDateFormat(1000 as never)).toThrow(/format/))

  it('Excel for Windows refused the files of the codes taken for no date, and read the quoted letter and the dates', () => {
    const set = JSON.parse(readFileSync(new URL('./excel/win/display/digits-date-letters.json', import.meta.url), 'utf8')) as { cases: { id: string; cell: { format: string } }[]; rejected: string[]; answers: Record<string, { text: string }> }
    const refused = new Set(set.cases.filter((c) => set.rejected.includes(c.id)).map((c) => c.cell.format))
    expect([...refused]).toEqual(['#,##0 m', '0 m', '0.0 m', '0 d', '0 y', '0 h', '0 s', '0.0 h', '0 samples', 'mm 0', 'd 0', '0 yy', '[Red]0 m', '0 m;0'])
    for (const f of refused) expect(isDateFormat(f), f).toBe(false)
    expect(set.answers['c10#1'].text).toBe('1235 m')
    expect(set.answers['c17#2'].text).toBe('31.01.2025')
    expect(isDateFormat('dd.mm.yyyy')).toBe(true)
  })
})
