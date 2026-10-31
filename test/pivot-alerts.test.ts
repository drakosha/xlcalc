/**
 * Changes of a pivot table's cells against what Excel for Windows did with them (the set `alerts`, its answers in
 * `test/excel/win/pivot/alerts.json`), each made in a session as `bench/coverage/pivot/alerts-predict.ts` makes it:
 * where Excel refused with one of its three alerts, the refusal or the error has that `reason`; where it refused with
 * an alert of its own (a table over the pivot table) the refusal has none; where it took the change, the cells of the
 * pivot table after it and the cells the case watches are Excel's, and a pivot table Excel deleted is gone.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { predictedAlert, type AlertPrediction } from '../bench/coverage/pivot/alerts-predict.js'
import type { PivotAnswer, PivotCase, PivotSet } from '../bench/coverage/pivot/cases.js'

const SETS = ['alerts', 'alerts-clear', 'alerts-held']

/**
 * Where the engine knowingly parts from Excel: text set through the object model over a range whose first cell is
 * outside a pivot table and the rest over it (`Formula2` over `A3:B5`, `A3` the blank row under its filters) leaves
 * every cell as it was, without an error; `fillCellText` with its active cell there refuses it (`part`), as the
 * interface's Ctrl+Enter there is not measured.
 */
const KNOWN: Record<string, string> = { 'typed-first-outside': 'part' }

/** Excel's alert by its text, as `reason` names it; another alert of its own as `own`. */
function alertOf(error: string): string {
  if (error.includes("We can't change this part of the PivotTable")) return 'part'
  if (error.includes('Cannot enter a formula for an item or field name')) return 'formula'
  if (error.includes('it will affect a PivotTable')) return 'affect'
  return 'own'
}

/** What the engine did, in the terms of Excel's answer: the alert, or the cells of each pivot table and those watched. */
function ours(p: AlertPrediction): string | Record<string, unknown> {
  if ('pivots' in p) return p.cells ? { pivots: p.pivots, cells: p.cells } : p.pivots
  const r = 'refused' in p ? p.refused : p.threw
  return r.reason ?? 'own'
}

describe.each(SETS)('the alerts of the set %s', (set) => {
  const answers = JSON.parse(readFileSync(new URL(`excel/win/pivot/${set}.json`, import.meta.url), 'utf8')) as { answers: Record<string, PivotAnswer> }
  it('are Excel’s, and a change Excel takes gives its cells', async () => {
    const { buildPivotCases } = (await import(`../bench/coverage/pivot/${set}.ts`)) as PivotSet
    const misses: string[] = []
    for (const c of buildPivotCases() as PivotCase[]) {
      const after = answers.answers[c.id].phases.at(-1)!
      const error = after.errors?.find((e) => e.startsWith('op '))
      const pivots = Object.fromEntries(Object.entries(after.pivots).flatMap(([k, v]) => ('error' in v ? [] : [[k, v.values]])))
      const want = KNOWN[c.id] ?? (error ? alertOf(error) : c.watch ? { pivots, cells: after.cells } : pivots)
      const got = ours(await predictedAlert(c))
      if (JSON.stringify(got) !== JSON.stringify(want)) misses.push(`${c.id}: ours ${JSON.stringify(got)}, Excel ${JSON.stringify(want)}`)
    }
    expect(misses).toEqual([])
  })
})
