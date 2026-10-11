/**
 * What cells show against Excel for Windows (`test/excel/win/display/`, `bench/windows`): every case's shown text and
 * colour, computed by `compute` in the book the harness laid out, as `display-excel.test.ts` does for the Mac. Excel
 * for Windows lays text out in pixels (a digit of Calibri 11 is 7 pixels at 96 dpi) where the engine follows the Mac's
 * points, so cases near a column's edge part; such differences, and the long date the engine keeps as the Mac's, are
 * covered by the rules of `WIN_SHOWN`, each with its reason.
 */
import { describe, expect, it } from 'vitest'
import type { StoredDisplay } from '../bench/coverage/display.js'
import type { DisplayAnswer, DisplayCase } from '../bench/coverage/display/cases.js'
import { localeDir, TAGS } from '../bench/coverage/locales.js'
import { compute, displayValue, type CellValue } from '../src/index.js'
import { CLOCK, displayBook, formatted, sameShown } from './display-cases.js'
import { winFiles } from './excel-win-cases.js'
import './locales.js'

/** The reasons that covered a case, filled as the sets run. */
const used = new Set<string>()

const SETS = TAGS.flatMap((locale) => winFiles<StoredDisplay>(`display/${localeDir(locale)}`).map((s) => ({ ...s, locale })))

/** A case's cell as Excel for Windows showed it and as the engine shows it. */
interface Shown {
  c: DisplayCase
  locale: string
  /** The cell's value as the engine computed it. */
  value: CellValue
  win: DisplayAnswer
  engine: { text: string; color?: string }
}

/** Widths from 0 to 256 characters (the widest column) in steps of a twelfth. */
const WIDTHS = Array.from({ length: 256 * 12 + 1 }, (_, i) => i / 12)
/** Windows's text is one the engine shows at another width (a merged area's too), with the cell's format or, for a built-in format by number, the code Windows reports. */
const atSomeWidth = (s: Shown) => {
  const formats = [formatted(s.c).numberFormat, ...(s.c.numFmtId !== undefined && s.win.code !== undefined ? [s.win.code] : [])]
  return formats.some((format) => WIDTHS.some((w) => displayValue(s.value, { numberFormat: format, locale: s.locale, width: w }).text === s.win.text))
}

/**
 * Why a case shows otherwise in Excel for Windows: each reason with the test of the cases it covers. The engine shows
 * cells as Excel for Windows does where the two Excels' regional data part, and lays them out as Excel for Mac does; a
 * difference no rule covers fails, and so does a rule that covers no case.
 */
const WIN_SHOWN: readonly [string, (s: Shown) => boolean][] = [
  ['Excel for Windows lays text out in pixels (a digit of Calibri 11 is 7 pixels at 96 dpi), the engine in the Mac\'s points: near the edge of a column the cell shows `#`, fewer digits of General, or a fill, as the engine shows it at another width', atSomeWidth],
  ['General with an exponent of three digits (below 1E-99) in a narrow column: Excel for Windows shows `#`, Excel for Mac and the engine `0` or the number', (s) => typeof s.value === 'number' && s.value !== 0 && Math.abs(s.value) < 1e-99 && /^#+$/.test(s.win.text)],
  ['`[$-F800]` is the system\'s long date: Windows ru-RU writes `31 января 2025 г.` without the weekday', (s) => /\[\$-F800\]/i.test(String(s.c.cell.format))],
]

for (const set of SETS) {
  describe(`cells shown as in Excel for Windows ${set.locale}: set ${set.name}`, async () => {
    const built = ((await import(`../bench/coverage/display/${set.name}.ts`)) as { buildDisplayCases: () => DisplayCase[] }).buildDisplayCases()

    it('the stored cases are the current ones, all answered', () => {
      expect(set.cases).toEqual(built)
      expect(built.filter((c) => !set.answers[c.id])).toEqual([])
    })

    it('every case shows what Excel for Windows showed, or differs for a reason of WIN_SHOWN', () => {
      const { book, at } = displayBook(set.cases, set.locale)
      const cells = compute(book, { clock: CLOCK }).sheets[0].cells
      const off = set.cases.flatMap((c) => {
        const out = cells[at.get(c.id)!]
        const engine = { text: out.text!, ...(out.color ? { color: out.color } : {}) }
        if (sameShown(set.answers[c.id], engine, out.value)) return []
        const rule = WIN_SHOWN.find(([, covers]) => covers({ c, locale: set.locale, value: out.value, win: set.answers[c.id], engine }))
        if (rule) {
          used.add(rule[0])
          return []
        }
        return [`${c.id} ${JSON.stringify(c.cell)} width ${c.width ?? 'default'}: Windows ${JSON.stringify(set.answers[c.id])}, engine ${JSON.stringify(engine)}`]
      })
      expect(off).toEqual([])
    })
  })
}

describe('known differences from Excel for Windows: shown text', () => {
  it.skipIf(!SETS.length)('every reason of WIN_SHOWN covers a case', () => {
    expect(WIN_SHOWN.map(([why]) => why).filter((why) => !used.has(why))).toEqual([])
  })
})
