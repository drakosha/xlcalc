/**
 * The stored answers of the fit sets in each locale of `LOCALES` (`test/excel/fit/`, `test/excel/win/fit/`, from `run.ts
 * fit <set> [locale]` on the Mac and `bench/windows/run.ts fit <set> [locale]`), and the engine's answer to one case:
 * its column fitted by `fitColumns`, or its texts typed in turn by `parseEntry` with the column widened by `entryWidth`
 * (texts set through an object model widen no column in either Excel).
 */
import type { StoredFit } from '../bench/coverage/fit.js'
import type { FitAnswer, FitCase, FitSet } from '../bench/coverage/fit/cases.js'
import { engineFit, type FitPrediction } from '../bench/coverage/fit/engine.js'
import { compute, entryWidth, fitColumns, parseEntry } from '../src/index.js'
import { localeSets } from './stored-sets.js'
import './locales.js'

export interface FitSetFile extends StoredFit {
  name: string
  tag: string
}

/** The sets stored under `test/excel/<root>fit/`, every locale's. */
export function fitSets(root: '' | 'win/'): FitSetFile[] {
  return localeSets<StoredFit>(`${root}fit`).map((r) => ({ name: r.set, tag: r.tag, ...r.stored }))
}

/** A set's current cases in a locale, by its generator. */
export async function currentCases(name: string, tag: string): Promise<FitCase[]> {
  return ((await import(`../bench/coverage/fit/${name}.ts`)) as FitSet).buildFitCases(tag)
}

/** The width Excel left the column, as the model reads the saved file: a width set by hand or fitted (255 at most), else null. */
export const excelWidth = (a: FitAnswer): number | null => (a.saved && (a.saved.customWidth || a.saved.bestFit) ? Math.min(255, a.saved.width) : null)

/** The engine's answer to a stored case, on the day of the run. */
export function engineAnswer(set: FitSetFile, c: FitCase): FitPrediction {
  const clock = { now: `${set.meta.date}T12:00:00Z`, seed: 1 }
  const engine = { compute: (book: Parameters<typeof compute>[0]) => compute(book, { clock }), parseEntry, fitColumns, entryWidth }
  return engineFit(engine, c, set.tag, String(set.meta.date), !c.typed || c.via === 'keys')
}
