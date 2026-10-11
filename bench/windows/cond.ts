/**
 * Conditional formats in Excel for Windows: each case of `bench/coverage/cond/<set>.ts` in the file the Mac step sends
 * (`writeXlsx`), the rules it makes through COM (`FormatConditions.Add` and its kin), its operation, the book computed,
 * the read cells' display format (`DisplayFormat`: Range.Text, fill and font colours, bold, italic), the rules each
 * sheet lists, and the book saved. Answers go to `test/excel/win/cond/[<tag>/]<set>.json` in the Mac's shape, the saved
 * files beside them; `test/cond-win.test.ts` reads them.
 */
import { condSet } from '../coverage/cond.js'
import { fromBgr, type CondAnswer, type CondCase, type Look } from '../coverage/cond/cases.js'
import { localeDir } from '../coverage/locales.js'
import { fileAnswers, secondsSince } from '../coverage/stored.js'
import { fileJobs, makeErrors, opSteps } from './books.js'
import { winFile } from './functions.js'
import { metaOf, type Step, type Tag, type WinEnv } from './remote.js'

type RawLook = [address: string, text: string, fill: number, colorIndex: number, font: number, bold: boolean, italic: boolean]
type RawRules = [sheet: string, rules: [number, string, string, number, boolean][]][]

/** xlNone: no fill. */
const NO_FILL = -4142

export async function stepCond(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { built, files, sent } = await condSet(set)
  const { books, cases } = built
  const store = fileAnswers<CondCase, CondAnswer>(winFile(`cond/${localeDir(tag)}${set}.json`), cases, sent)
  const t0 = Date.now()
  const day = new Date().toLocaleDateString('sv-SE')
  const save = () => store.save({ ...metaOf(e, tag), day, seconds: secondsSince(t0) }, { books, cases, ...(built.locale ? { locale: built.locale } : {}) })
  await fileJobs(
    `cond ${set}`,
    `cond-${set}`,
    store,
    files,
    (c) => [...(c.make ?? []).map((m): Step => ({ do: 'cond', kind: 'make', ...m })), ...opSteps(c), { do: 'full' }, { do: 'cond', kind: 'looks', read: c.read }, { do: 'cond', kind: 'rules' }, { do: 'saveAs' }],
    (c, a) => {
      const { n, errors } = makeErrors(c, a.answers)
      const looks: CondAnswer['looks'] = {}
      for (const [sheet, cells] of a.answers[n + 1] as [string, RawLook[]][])
        looks[sheet] = Object.fromEntries(cells.map(([addr, text, fill, ci, font, bold, italic]): [string, Look] => [addr, { text, fill: ci === NO_FILL ? null : fromBgr(fill), font: fromBgr(font), bold, italic }]))
      const rules = Object.fromEntries((a.answers[n + 2] as RawRules).map(([sheet, list]) => [sheet, list]))
      return { ...(errors.length ? { error: errors.join('; ') } : {}), looks, rules }
    },
    save,
  )
}
