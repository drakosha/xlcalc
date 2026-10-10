/**
 * `entryText` against Excel (`test/excel/[win/]entry/[ru-RU/]edit*.json`, the sets `bench/coverage/entry/edit*.ts`): each
 * cell a text was typed into, as Excel saved it, opens with the text Excel's formula bar showed for it on the Mac (read
 * through the window's accessibility). Both Excels' object model gives the number itself for a cell under a date, time or
 * percent format (`Range.FormulaLocal` is `45688`, `0.12`); under any other format it gives what the bar shows, the
 * apostrophe being the prefix character, and the cells under other formats are checked against it on Windows and the
 * Mac. The predictions committed before Excel for Mac ran keep their misses listed.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredEntry } from '../bench/coverage/entry.js'
import { formulaBar, type EntryAnswer } from '../bench/coverage/entry/cases.js'
import { localeDir, TAGS } from '../bench/coverage/locales.js'
import { entryText, parseEntry, sameValue, type CellInput } from '../src/index.js'

/** Why a case's text differs from Excel's formula bar, by `<tag>:<set>:<id>`. */
const KNOWN: Readonly<Record<string, string>> = {}

/** The reason of a text typed with an apostrophe it did not need: Excel keeps it as the cell's prefix, a book holds no such mark. */
const KEPT = "Excel's formula bar keeps the apostrophe a person typed before text that needed none; a book holds no such mark, and entryText puts one only before text that would read as something else"

/** Cases whose predicted formula bar Excel for Mac did not show, by `<tag>:<set>:<id>`. */
const MISSED: readonly string[] = []

interface Run {
  excel: 'Mac' | 'Windows'
  tag: string
  set: string
  stored: StoredEntry
}

const RUNS: Run[] = (['Mac', 'Windows'] as const).flatMap((excel) =>
  TAGS.flatMap((tag) => {
    const dir = new URL(`./excel/${excel === 'Windows' ? 'win/' : ''}entry/${localeDir(tag)}`, import.meta.url)
    if (!existsSync(dir)) return []
    return readdirSync(dir)
      .filter((f) => /^edit.*\.json$/.test(f))
      .map((f) => ({ excel, tag, set: f.slice(0, -5), stored: JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredEntry }))
  }),
)

/** The text a cell opens with by `entryText`, for a cell that holds a value (a formula has its own text). */
function opened(cell: CellInput | null, tag: string, date: string): string | null {
  if (!cell || cell.value === undefined || cell.formula !== undefined) return null
  return entryText(cell.value as string | number | boolean, { locale: tag, now: date, ...(cell.format !== undefined ? { numberFormat: cell.format } : {}) })
}

/** A typed text Excel kept with an apostrophe it did not need: its value reads as itself without one. */
const kept = (a: EntryAnswer, tag: string, date: string) => a.prefix === "'" && typeof a.cell?.value === 'string' && opened(a.cell, tag, date) === a.cell.value

/** Whether the cell's format shows a date, a time or a percent, where the object model gives the number itself. */
const datedOrPercent = (cell: CellInput, tag: string, date: string) => typeof cell.value === 'number' && opened(cell, tag, date) !== opened({ value: cell.value }, tag, date)

for (const run of RUNS) {
  const { excel, tag, set, stored } = run
  const date = String(stored.meta.date)
  describe(`entryText against Excel for ${excel} ${tag}: ${set}`, () => {
    const cases = stored.cases.filter((c) => stored.answers[c.id] && opened(stored.answers[c.id].cell, tag, date) !== null)

    const bars = cases.filter((c) => stored.answers[c.id].bar !== undefined)
    if (bars.length)
      it('the formula bar Excel showed', () => {
        const off = bars.flatMap((c) => {
          const a = stored.answers[c.id]
          const want = formulaBar(a.bar!)
          const got = opened(a.cell, tag, date)
          const key = `${tag}:${set}:${c.id}`
          if (got === want) return KNOWN[key] ? [`${key} is listed as known but agrees`] : []
          return KNOWN[key] || kept(a, tag, date) ? [] : [`${key} ${JSON.stringify(c.text)}: Excel ${JSON.stringify(want)}, entryText ${JSON.stringify(got)}`]
        })
        expect(off).toEqual([])
      })

    it('the object model under formats that show no date, time or percent', () => {
      const off = cases.flatMap((c) => {
        const a = stored.answers[c.id]
        if (a.formulaLocal === undefined || datedOrPercent(a.cell!, tag, date) || kept(a, tag, date)) return []
        const want = (a.prefix ?? '') + a.formulaLocal
        const got = opened(a.cell, tag, date)
        return got === want ? [] : [`${c.id} ${JSON.stringify(c.text)}: Excel ${JSON.stringify(want)}, entryText ${JSON.stringify(got)}`]
      })
      expect(off).toEqual([])
    })

    it('the text reads back to the value Excel saved, but day 0 (`1/0/1900`, which Excel reads as text too)', () => {
      for (const c of cases) {
        const { cell } = stored.answers[c.id]
        if (cell!.value === 0 && datedOrPercent(cell!, tag, date)) continue
        const r = parseEntry(opened(cell, tag, date)!, { locale: tag, now: date, ...(typeof cell!.value === 'string' && cell!.format !== undefined ? { cell: { format: cell!.format } } : {}) })
        const back = r.ok ? r.input?.value : undefined
        // numbers to the 15 digits the value keeps (`0.333333333333333`)
        expect(typeof back === 'number' && typeof cell!.value === 'number' ? sameValue(back, cell!.value) : back === cell!.value, `${c.id}: ${JSON.stringify(back)}`).toBe(true)
      }
    })

    const predicted = new URL(`../bench/coverage/entry/${set}.predicted.json`, import.meta.url)
    if (excel === 'Mac' && bars.length && existsSync(predicted))
      it('the prediction committed before Excel ran is what its formula bar showed, but for the misses listed', () => {
        const edit = (JSON.parse(readFileSync(predicted, 'utf8')) as { edit: Record<string, Record<string, string | null>> }).edit[tag]
        const missed = bars.filter((c) => edit[c.id] !== formulaBar(stored.answers[c.id].bar!)).map((c) => `${tag}:${set}:${c.id}`)
        expect(missed).toEqual(MISSED.filter((k) => k.startsWith(`${tag}:${set}:`)))
      })
  })
}

describe('known differences of entryText', () => {
  it('the reason of a kept apostrophe covers a case', () => {
    const barred = RUNS.flatMap(({ tag, stored }) => stored.cases.flatMap((c) => (stored.answers[c.id]?.bar === undefined ? [] : [[stored.answers[c.id], tag, String(stored.meta.date)] as const])))
    expect(!barred.length || barred.some(([a, tag, date]) => kept(a, tag, date)), KEPT).toBe(true)
  })
})
