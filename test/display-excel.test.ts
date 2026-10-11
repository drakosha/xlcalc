/**
 * What cells show against Excel (`test/excel/display/`, Excel 16.113 for Mac in en-US and in ru-RU): every case's
 * shown text (Range.Text) and its section's colour, computed by `compute` in a book laid out as the harness laid
 * it out; held-out cases also keep the answer predicted before Excel ran. Where the regional data of Excel for Mac
 * and for Windows part, the engine shows what Windows shows: those cases are covered by the rules of `MAC_SHOWN`.
 */
import { describe, expect, it } from 'vitest'
import type { DisplayAnswer, DisplayCase } from '../bench/coverage/display/cases.js'
import { compute, displayValue, type CellDisplay } from '../src/index.js'
import { CLOCK, DISPLAY_SETS, displayBook, formatted, ROUNDED_MOMENT, sameShown } from './display-cases.js'

/** A case as Excel for Mac showed it and as the engine shows it. */
interface Shown {
  c: DisplayCase
  locale: string
  mac: DisplayAnswer
  engine: CellDisplay
}

const SHORT_MONTHS: [RegExp, string][] = [[/фев(?!р)/g, 'февр'], [/мар(?!т)/g, 'март'], [/июн(?!ь)/g, 'июнь'], [/июл(?!ь)/g, 'июль'], [/сен(?!т)/g, 'сент'], [/ноя(?!б)/g, 'нояб']]
const lower = (t: string) => t.toLowerCase()
/** Excel's default width as xlsx stores it (8 digits and the padding): `displayValue` without a width has no limit. */
const DEFAULT_WIDTH = 8.83203125

/**
 * Why a case shows otherwise in Excel for Mac: where its regional data part from Windows's, the engine follows Windows
 * (decision 22). Each reason with the test of the cases it covers; a rule that covers no case fails.
 */
const MAC_SHOWN: readonly [string, (s: Shown) => boolean][] = [
  ['Excel for Mac shows the built-in short date by number as `m/d/yy` in en-US (`1/31/25`); the engine as Windows, `m/d/yyyy` (`1/31/2025`)', (s) => s.locale === 'en-US' && s.c.numFmtId !== undefined && s.mac.code !== undefined && typeof s.c.cell.value === 'number' && displayValue(s.c.cell.value, { numberFormat: s.mac.code, locale: s.locale, width: s.c.width ?? DEFAULT_WIDTH }).text === s.mac.text],
  ['Excel for Mac writes ru-RU month names in lower case (`январь`); the engine with a capital, as Windows (`Январь`)', (s) => s.mac.text !== s.engine.text && lower(s.mac.text) === lower(s.engine.text)],
  ['a ru-RU month name with a capital takes more room: where Excel for Mac\'s lower-case name just fits, the engine\'s does not', (s) => /^#+$/.test(s.engine.text) && typeof s.c.cell.value === 'number' && lower(displayValue(s.c.cell.value, { numberFormat: formatted(s.c).numberFormat, locale: s.locale, width: 100 }).text) === lower(s.mac.text)],
  ['Excel for Mac writes ru-RU short month names `февр`, `март`, `июнь`, `июль`, `сент`, `нояб`; the engine three letters, as Windows', (s) => SHORT_MONTHS.reduce((t, [re, x]) => t.replace(re, x), lower(s.engine.text)) === lower(s.mac.text)],
  ['Excel for Mac writes ru-RU times of the built-in formats and of `[$-F400]` with a leading zero (`06:00`); the engine without, as Windows (`6:00`)', (s) => s.mac.text.replace(/(^|\s)0(\d:)/g, '$1$2') === s.engine.text],
  ['Excel for Mac\'s longer ru-RU month names in the built-in dates by number (15 to 17) do not fit where the engine\'s three letters do', (s) => s.locale === 'ru-RU' && [15, 16, 17].includes(s.c.numFmtId ?? -1) && /^#+$/.test(s.mac.text)],
]

/** The rules that covered a case, filled as the sets run. */
const used = new Set<string>()
/** The rule of MAC_SHOWN that covers a case, noted as used. */
function macRule(s: Shown): boolean {
  const rule = MAC_SHOWN.find(([, covers]) => covers(s))
  if (rule) used.add(rule[0])
  return !!rule
}

/** Whether the known difference `ROUNDED_MOMENT` covers a case, noted as used. */
function rounded(c: DisplayCase, locale: string, excel: DisplayAnswer): boolean {
  const covers = ROUNDED_MOMENT[1](c, locale, excel.text)
  if (covers) used.add(ROUNDED_MOMENT[0])
  return covers
}

/** Held-out cases whose prediction Excel did not confirm, by `<locale>:<set>:<id>` (the engine was fixed after the run). */
const MISSED: readonly string[] = [
  'en-US:held:condition#6',
  'ru-RU:held-ru:condition#6',
  ...[9, 12, 17, 18, 19, 22, 25, 26, 28, 29, 31, 32, 33, 34, 35].map((n) => `en-US:seconds-held:day#${n}`),
]

for (const set of DISPLAY_SETS) {
  describe(`cells shown as in Excel ${set.locale}: set ${set.name}`, async () => {
    const built = ((await import(`../bench/coverage/display/${set.name}.ts`)) as { buildDisplayCases: () => typeof set.cases }).buildDisplayCases()

    it('the stored cases are the current ones, all answered', () => {
      expect(set.cases).toEqual(built)
      expect(built.filter((c) => !set.answers[c.id])).toEqual([])
    })

    it('every case shows what Excel showed', () => {
      const { book, at } = displayBook(set.cases, set.locale)
      const cells = compute(book, { clock: CLOCK }).sheets[0].cells
      const off = set.cases.flatMap((c) => {
        const out = cells[at.get(c.id)!]
        const engine = { text: out.text!, ...(out.color ? { color: out.color } : {}) }
        if (sameShown(set.answers[c.id], engine, out.value) || macRule({ c, locale: set.locale, mac: set.answers[c.id], engine }) || rounded(c, set.locale, set.answers[c.id])) return []
        return [`${c.id} ${JSON.stringify(c.cell)} width ${c.width ?? 'default'}: Excel ${JSON.stringify(set.answers[c.id])}, engine ${JSON.stringify(engine)}`]
      })
      expect(off).toEqual([])
    })

    it('displayValue shows each value as its cell', () => {
      const off = set.cases.flatMap((c) => {
        if (c.cell.formula !== undefined || (c.span ?? 1) > 1) return []
        const value = c.cell.value === undefined || (typeof c.cell.value === 'object' && c.cell.value !== null) ? null : c.cell.value
        const shown = displayValue(value, { numberFormat: formatted(c).numberFormat, locale: set.locale, width: c.width ?? DEFAULT_WIDTH })
        if (sameShown(set.answers[c.id], shown, value) || macRule({ c, locale: set.locale, mac: set.answers[c.id], engine: shown })) return []
        return [`${c.id}: Excel ${JSON.stringify(set.answers[c.id])}, displayValue ${JSON.stringify(shown)}`]
      })
      expect(off).toEqual([])
    })

    const held = set.cases.filter((c) => c.predicted)
    if (held.length)
      it('held-out cases: the prediction committed before Excel ran is what Excel showed, but for the misses listed', () => {
        // the codes Excel reports for a built-in format are no part of the prediction
        const shown = ({ text, color }: DisplayAnswer) => JSON.stringify({ text, color })
        const missed = held.filter((c) => shown(c.predicted!) !== shown(set.answers[c.id])).map((c) => `${set.locale}:${set.name}:${c.id}`)
        expect(missed).toEqual(MISSED.filter((k) => k.startsWith(`${set.locale}:${set.name}:`)))
      })
  })
}

describe('known differences from Excel for Mac: shown text', () => {
  it('every reason of MAC_SHOWN covers a case', () => {
    expect(MAC_SHOWN.map(([why]) => why).filter((why) => !used.has(why))).toEqual([])
  })

  it('the known difference ROUNDED_MOMENT covers a case', () => {
    expect(used.has(ROUNDED_MOMENT[0])).toBe(true)
  })
})
