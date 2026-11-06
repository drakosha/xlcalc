/**
 * Letters in number format codes against Excel: TEXT by the table of Excel for Mac (`links/letters-text`) and in Excel for Windows, the codes
 * that table has not (`links/letters-held`, `links/letters-more`, a sample on the Mac in `links/letters-mac`), ru-RU's (`sets/text-letters-ru`),
 * and which codes a file opens with (`display/letters-open`, `letters-open-2`, `letters-open-3`). Each set's prediction was committed before Excel ran it, and
 * what it missed is named here; the engine answers as Excel but for `ENGINE_KNOWN`.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredLinks } from '../bench/coverage/links.js'
import type { DisplayCase } from '../bench/coverage/display/cases.js'
import { REFUSED_PREDICTED as OPEN_REFUSED } from '../bench/coverage/display/letters-open.js'
import { REFUSED_PREDICTED as OPEN_2_REFUSED } from '../bench/coverage/display/letters-open-2.js'
import { REFUSED_PREDICTED as OPEN_3_REFUSED } from '../bench/coverage/display/letters-open-3.js'
import { buildLinksCases as heldCases } from '../bench/coverage/links/letters-held.js'
import { buildLinksCases as macCases } from '../bench/coverage/links/letters-mac.js'
import { buildLinksCases as moreCases } from '../bench/coverage/links/letters-more.js'
import { buildLinksCases as textCases, MAC_TABLE } from '../bench/coverage/links/letters-text.js'
import { RU_PREDICTED } from '../bench/coverage/sets/text-letters-ru.js'
import { compute, displayValue, stepDecimals } from '../src/index.js'
import { splitSections } from '../src/numfmt/parse.js'
import { formatRefused } from '../src/text.js'
import { writeXlsx } from '../src/xlsx.js'

const json = <T>(path: string) => JSON.parse(readFileSync(new URL(`./excel/${path}`, import.meta.url), 'utf8')) as T
/** What TEXT gives, as a stored answer gives it; an answer of no text is stored as nothing. */
const answer = (v: unknown) => (typeof v === 'object' && v !== null ? { t: 'e', v: (v as { error: string }).error } : { t: 's', v: v ?? '' })
const engine = (formula: string) => answer(compute({ sheets: [{ name: 'S', cells: { A1: { formula: `=${formula}` } } }] }).sheets[0].cells.A1?.value)

/** Predictions Excel did not confirm, by what a case's `about` says. */
const MISSED: Record<string, string[]> = {
  'win/links/letters-held': [
    // runs of `b` and `B`: `B` refuses only at the end of the code after no other `B`, and `a` runs are of one case
    'bB', 'BbB', 'Byy', 'Blue', 'Bob', 'aAa', 'AaA', 'aAaA',
    // `m` after hours across the date's other codes
    'h bb mm',
    // a word in a fourth section is no date's code
    '0;0;0;General', '0;0;0;AM/PM',
    // a condition in a third section refuses
    '0;[>1]0;[>2]0', '0;0;[>1]0',
    // no digit before the point
    '.E+0',
    // a digit beside a date refuses; one calendar to a section
    'b3', 'yyb1', 'B1B2',
    // Hijri names, and the prediction forgot that `A/P` comes first and keeps its case
    'B2aaa', 'a/p', 'aa/p', 'aaa/p', 'aaaa/p', 'aaam/pm',
  ],
  'win/links/letters-more': ['B;0', 'aaaAAA', 'h dd mm', 'h yy mm', 'h e mm', 'h ddd mm', 'mm aaa ss', 'mm bb ss', '0;[>1]0;0'],
}

/**
 * Where the engine answers otherwise: a mantissa with no digit before its point (`.0E+0` is `.5E+5`, `.E+0` the odd `.E+1`); a value a first section
 * without a condition does not take when the second has one; the two names of `aaaAAA`, which Excel makes one long name.
 */
const ENGINE_KNOWN = new Set(['TEXT(46096.5,".E+0")', 'TEXT(46096.5,".0E+0")', 'TEXT(-1,"0;[>1]0")', 'TEXT(-1,"[Red]0;[>1]0")', 'TEXT(46096.5,"aaaAAA")'])

describe.each([
  ['win/links/letters-text', textCases],
  ['win/links/letters-held', heldCases],
  ['win/links/letters-more', moreCases],
  ['links/letters-mac', macCases],
] as const)('%s', (path, cases) => {
  const set = json<StoredLinks>(`${path}.json`)
  const shown = (id: string) => answer((set.answers[id].reads[1].values['S!A1'] as { v?: unknown; t?: string } | undefined)?.t === 'e' ? { error: 'VALUE' } : (set.answers[id].reads[1].values['S!A1'] as { v?: unknown } | undefined)?.v)

  it('the predictions Excel did not confirm are the ones named', () => {
    const missed = cases().cases.filter((c) => JSON.stringify(shown(c.id)) !== JSON.stringify(answer(c.predicted!['S!A1']))).map((c) => c.about.slice(c.about.indexOf(' with ') + 6))
    expect(missed.sort()).toEqual([...(MISSED[path] ?? [])].sort())
  })

  it('the engine answers as Excel', () => {
    const off = cases().cases.filter((c) => !ENGINE_KNOWN.has(c.linked.S.A1 as string) && JSON.stringify(engine(c.linked.S.A1 as string)) !== JSON.stringify(shown(c.id)))
    expect(off.map((c) => `${c.linked.S.A1}: ${JSON.stringify(shown(c.id))}, the engine ${JSON.stringify(engine(c.linked.S.A1 as string))}`)).toEqual([])
  })

  it('what the engine knows of differs still', () => {
    for (const c of cases().cases.filter((x) => ENGINE_KNOWN.has(x.linked.S.A1 as string))) expect(engine(c.linked.S.A1 as string)).not.toEqual(shown(c.id))
  })
})

describe("Excel for Mac's table", () => {
  it('a cell shows what TEXT gives, and a file is refused where TEXT is #VALUE!', () => {
    for (const [code, text, fill] of MAC_TABLE) {
      expect(formatRefused(code), code).toBe(text === null)
      if (text !== null && !fill) expect(displayValue(46096.5, { numberFormat: code, width: 40 }).text, code).toBe(text)
    }
  })
})

describe.each([
  ['letters-open', OPEN_REFUSED, ['0;[>1]0;[>2]0', '0;0;[>1]0']],
  ['letters-open-2', OPEN_2_REFUSED, []],
  ['letters-open-3', OPEN_3_REFUSED, []],
] as const)('files with the codes of %s', (name, predicted, missed) => {
  const set = json<{ cases: (DisplayCase & { predicted?: { text: string } })[]; rejected: string[]; answers: Record<string, { text: string }> }>(`win/display/${name}.json`)
  const refused = set.cases.filter((c) => set.rejected.includes(c.id))

  it('Excel for Windows refused what was predicted, and the codes named', () => {
    expect(refused.map((c) => c.cell.format).sort()).toEqual([...predicted, ...missed].sort())
  })

  it('formatRefused says the same of each code', () => {
    for (const c of set.cases) expect(formatRefused(c.cell.format as string), c.cell.format as string).toBe(refused.includes(c))
  })

  it('what it opened shows as predicted, and as the engine shows it but a fill, which Windows lays out in pixels', () => {
    for (const c of set.cases.filter((x) => set.answers[x.id])) {
      const code = c.cell.format as string
      if (c.predicted) expect(set.answers[c.id].text, code).toBe(c.predicted.text)
      if (!code.includes('*') && !ENGINE_KNOWN.has(`TEXT(46096.5,"${code}")`)) expect(displayValue(46096.5, { numberFormat: code, width: 40 }).text, code).toBe(set.answers[c.id].text)
    }
  })
})

describe('text-letters-ru', () => {
  const set = json<{ cases: { id: string; formula: string }[]; values: Record<string, { t: string; v: unknown }> }>('win/cases/ru-RU/text-letters-ru.json')

  it('Excel for Windows in ru-RU answered each as predicted: Latin letters read as in en-US but its date letters, which stay literals', () => {
    RU_PREDICTED.forEach(([code, text], i) => {
      const got = set.values[`letters#${i + 1}`]
      expect(got?.t === 'e' ? null : (got?.v ?? ''), code).toBe(text)
    })
  })
})

describe('a padding and a fill take the next character', () => {
  const CORPUS: [string, number][] = [
    [String.raw`[Red]\+#,##0_ ;[Green]\-#,##0_;;`, 3],
    [String.raw`[RED]\+#,##0_ ;[GREEN]\-#,##0_;;;`, 4],
    ['0_;;0', 2],
    ['0*;;0', 2],
    ['0\\;;0', 2],
    ['"_;";0', 2],
  ]

  it.each(CORPUS)('%s has its sections, opens and is written', async (code, sections) => {
    expect(splitSections(code)).toHaveLength(sections)
    expect(formatRefused(code)).toBe(false)
    expect((await writeXlsx({ book: { sheets: [{ name: 'S', cells: { A1: { value: 1, numberFormat: code } } }] } })).data.length).toBeGreaterThan(0)
  })

  it('Increase Decimal keeps the character a padding took', () => {
    expect(stepDecimals('0_;;0', 1, 1)).toBe('0.0_;;0.0')
    expect(stepDecimals('0*;;0', 1, 1)).toBe('0.0*;;0.0')
  })

  it('five sections and three conditions stay refused', () => {
    for (const code of ['0;0;0;@;', '0;0;0;0;0', '[>=1][Green]0.0%;[>=0.8][Blue]0.0%;[>=0.5]0.0%']) expect(formatRefused(code), code).toBe(true)
  })
})

describe('a code no format reads', () => {
  it('shows as General in a cell, whatever the file held', () => {
    for (const code of ['n', 'B', '0g', 'aaa0']) expect(displayValue(46096.5, { numberFormat: code }).text, code).toBe('46096.5')
  })

  it('literals and whole words hold an `n`', () => {
    for (const code of ['"n"', '\\n', '_n0', '*n0', 'General', 'GENERAL', '[$n-409]0']) expect(formatRefused(code), code).toBe(false)
  })
})
