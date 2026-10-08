/**
 * What cells show against Excel for Windows (`test/excel/win/display/`, `bench/windows`): every case's shown text and
 * colour, computed by `compute` in the book the harness laid out, as `display-excel.test.ts` does for the Mac. Excel
 * for Windows lays text out in pixels (a digit of Calibri 11 is 7 pixels at 96 dpi) where the engine follows the Mac's
 * points, so cases near a column's edge part; such differences, and the places where Windows's own regional data differ
 * from the Mac's, are covered by the rules of `WIN_SHOWN`, each with its reason.
 */
import { describe, expect, it } from 'vitest'
import type { StoredDisplay } from '../bench/coverage/display.js'
import type { DisplayAnswer, DisplayCase } from '../bench/coverage/display/cases.js'
import { compute, displayValue, type CellValue } from '../src/index.js'
import { CLOCK, displayBook, formatted, sameShown } from './display-cases.js'
import { winFiles } from './excel-win-cases.js'

/** The reasons that covered a case, filled as the sets run. */
const used = new Set<string>()

const SETS = [...winFiles<StoredDisplay>('display/').map((s) => ({ ...s, locale: 'en-US' })), ...winFiles<StoredDisplay>('display/ru-RU/').map((s) => ({ ...s, locale: 'ru-RU' }))]

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
  const formats = [formatted(s.c).format, ...(s.c.numFmtId !== undefined && s.win.code !== undefined ? [s.win.code] : [])]
  return formats.some((format) => WIDTHS.some((w) => displayValue(s.value, { format, locale: s.locale, width: w }).text === s.win.text))
}
const SHORT_MONTHS: [RegExp, string][] = [[/февр/g, 'фев'], [/март/g, 'мар'], [/июнь/g, 'июн'], [/июль/g, 'июл'], [/сент/g, 'сен'], [/нояб/g, 'ноя']]

/**
 * Why a case shows otherwise in Excel for Windows: each reason with the test of the cases it covers. The engine shows
 * cells as Excel for Mac does; a difference no rule covers fails, and so does a rule that covers no case.
 */
const WIN_SHOWN: readonly [string, (s: Shown) => boolean][] = [
  ['a built-in format by number is the code of Windows\'s regional settings (en-US 14 is `m/d/yyyy`, `1/31/2025`; the Mac\'s `m/d/yy`, `1/31/25`): the engine shows Windows\'s text with the code Windows reports', (s) => {
    if (s.c.numFmtId === undefined || s.win.code === undefined || typeof s.c.cell.value !== 'number') return false
    return displayValue(s.c.cell.value, { format: s.win.code, locale: s.locale, width: s.c.width }).text === s.win.text
  }],
  ['Excel for Windows lays text out in pixels (a digit of Calibri 11 is 7 pixels at 96 dpi), the engine in the Mac\'s points: near the edge of a column the cell shows `#`, fewer digits of General, or a fill, as the engine shows it at another width', atSomeWidth],
  ['General with an exponent of three digits (below 1E-99) in a narrow column: Excel for Windows shows `#`, Excel for Mac and the engine `0` or the number', (s) => typeof s.value === 'number' && s.value !== 0 && Math.abs(s.value) < 1e-99 && /^#+$/.test(s.win.text)],
  ['ru-RU built-in date formats by number (15 to 17): Windows reports its own regional code (`ДД.МММ.ГГ`, three-letter months), which fits columns where the Mac\'s `dd.mmm.yy` with `нояб` does not', (s) => s.locale === 'ru-RU' && [15, 16, 17].includes(s.c.numFmtId ?? -1)],
  ['ru-RU month names: Excel for Windows writes them with a capital (`Январь`, `Я`), Excel for Mac and the engine in lower case', (s) => s.win.text !== s.engine.text && s.win.text.toLowerCase() === s.engine.text.toLowerCase()],
  ['ru-RU short month names: Excel for Windows takes three letters (`фев`, `мар`, `июн`, `июл`, `сен`, `ноя`), Excel for Mac and the engine `февр`, `март`, `июнь`, `июль`, `сент`, `нояб`', (s) => SHORT_MONTHS.reduce((t, [re, x]) => t.replace(re, x), s.engine.text).toLowerCase() === s.win.text.toLowerCase()],
  ['`[$-F800]` and `[$-F400]` are the system\'s long date and time: Windows ru-RU writes `31 января 2025 г.` without the weekday and the hour without a leading zero', (s) => /\[\$-F[48]00\]/i.test(String(s.c.cell.format))],
  ['ru-RU built-in time formats by number: Windows takes its own short time, the hour without a leading zero (`6:00`, the Mac `06:00`)', (s) => s.c.numFmtId !== undefined && s.engine.text.replace(/(^|\s)0(\d:)/g, '$1$2') === s.win.text],
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
