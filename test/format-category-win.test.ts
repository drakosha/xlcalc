/**
 * `numberFormatCategory` against the ribbon's Number Format box of Excel for Windows (`test/excel/win/ribbon/`, the
 * Windows step `ribbon` over the books of the ui sets `ribbon` and `ribbon-ru-held`): each cell's box against the
 * function's name for the cell's format, the known differences by cell with their reasons. Excel for Mac read the set
 * `ribbon` too (`test/excel/ui/`): how many cells the two Excels name alike. The predictions committed before the run
 * (`bench/windows/ribbon/<set>.predicted.json`) are kept with how many cells Excel confirmed.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredUi } from '../bench/coverage/ui.js'
import { ribbonOf, type UiSet } from '../bench/coverage/ui/cases.js'
import { localeDir, TAGS } from '../bench/coverage/locales.js'
import { ribbonCases, type StoredRibbon } from '../bench/windows/ribbon.js'
import './locales.js'

const LISTS =
  "Excel for Windows names a date or time with ;@ by the lists of its Format Cells window, which the package does not hold: in ru-RU dd/mm/yy;@, d/m;@, [$-419]d mmm;@ and en-US's m/d/yyyy;@ are Дата, but dd/mm;@, d mmm;@ and [$-419]d mmmm yyyy;@ none, and dd/mm/yy h:mm;@ is Время; numberFormatCategory names every date or time with ;@ by its letters, as the codes with ;@ come from such lists"

/** The cells where the function knowingly names another category than Excel for Windows, by `<tag>:<set>:<case>:<cell>`, with the reason. */
const KNOWN: Record<string, string> = Object.fromEntries(
  [...['B2', 'B6', 'B7', 'B13', 'B15', 'B16', 'B17', 'B18', 'B19', 'B20', 'B21', 'B22', 'B23', 'B24'].map((x) => `held-1:${x}`), 'held-2:B4', 'held-2:B6'].map((k) => [`ru-RU:ribbon-ru-held:${k}`, LISTS]),
)

/** How many cells of each set's prediction Excel for Windows confirmed, by `<tag>:<set>`. */
const CONFIRMED: Record<string, number> = { 'ru-RU:ribbon': 318, 'ru-RU:ribbon-ru-held': 29 }

/** How many cells of the set `ribbon` Excel for Mac and for Windows name alike, by tag. */
const MAC_SAME: Record<string, number> = { 'ru-RU': 266 }

const RUNS = TAGS.flatMap((tag) => {
  const dir = new URL(`./excel/win/ribbon/${localeDir(tag)}`, import.meta.url)
  if (!existsSync(dir)) return []
  return readdirSync(dir)
    .filter((f) => f.endsWith('.json'))
    .map((f) => ({ tag, set: f.slice(0, -5), stored: JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredRibbon }))
})

/** Each cell of a run: its case, its address, what Excel for Windows showed. */
const cellsOf = (stored: StoredRibbon) => stored.cases.flatMap((c) => Object.entries(stored.answers[c.id]).map(([x, a]) => ({ c, x, box: a.box })))

for (const { tag, set, stored } of RUNS)
  describe(`ribbon ${tag} ${set} (Excel for Windows)`, () => {
    it("the stored cases are the set's current ribbon cases", async () => {
      expect(stored.cases).toEqual(ribbonCases((await import(`../bench/coverage/ui/${set}.ts`)) as UiSet))
    })

    it("each cell's box is numberFormatCategory's name for its format, but the known differences, each still one", () => {
      const off: string[] = []
      const fixed: string[] = []
      for (const { c, x, box } of cellsOf(stored)) {
        const key = `${tag}:${set}:${c.id}:${x}`
        const got = ribbonOf(c, tag)[x]
        if (got !== box && !KNOWN[key]) off.push(`${c.id}:${x} ${JSON.stringify(c.cells[x]?.numberFormat ?? 0)}: Excel ${box}, the function ${got}`)
        if (got === box && KNOWN[key]) fixed.push(`${key}: ${KNOWN[key]}`)
      }
      expect({ off, fixed }).toEqual({ off: [], fixed: [] })
    })

    it('the predictions written before Excel ran it', () => {
      const predicted = JSON.parse(readFileSync(new URL(`../bench/windows/ribbon/${set}.predicted.json`, import.meta.url), 'utf8')).locales[tag] as Record<string, Record<string, string>>
      expect(cellsOf(stored).filter(({ c, x, box }) => predicted[c.id][x] === box).length).toBe(CONFIRMED[`${tag}:${set}`])
    })

    if (set === 'ribbon')
      it('Excel for Mac over the same books: the cells it names as Excel for Windows does', () => {
        const mac = JSON.parse(readFileSync(new URL(`./excel/ui/${localeDir(tag)}ribbon.json`, import.meta.url), 'utf8')) as StoredUi
        expect(cellsOf(stored).filter(({ c, x, box }) => mac.answers[c.id].cells[x].ribbon!.box === box).length).toBe(MAC_SAME[tag])
      })
  })

it('every known difference is a cell of a stored run', () => {
  const cells = new Set(RUNS.flatMap(({ tag, set, stored }) => cellsOf(stored).map(({ c, x }) => `${tag}:${set}:${c.id}:${x}`)))
  expect(Object.keys(KNOWN).filter((k) => !cells.has(k))).toEqual([])
})
