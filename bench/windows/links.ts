/**
 * Links to other workbooks in Excel for Windows: each case of `bench/coverage/links/<set>.ts` in the file the Mac step
 * sends, opened without updating its links (`Workbooks.Open(path, 0)`, as every book of this harness), the linked books
 * listed (`LinkSources`), the rectangles read through COM (`Value2`, and each cell's `Text`, `Formula`, `FormulaLocal`)
 * after opening and after a full calculation, and the book saved. Answers go to `test/excel/win/links/<set>.json` in
 * the Mac's shape, the saved files beside them.
 */
import { cellsOf, eachRectCell } from '../coverage/case-book.js'
import { linksSet } from '../coverage/links.js'
import type { LinksAnswer, LinksCase, LinksRead } from '../coverage/links/cases.js'
import { fileAnswers, secondsSince } from '../coverage/stored.js'
import { fileJobs } from './books.js'
import { valOf, winFile, type WinCell } from './functions.js'
import { metaOf, type Step, type Tag, type WinEnv } from './remote.js'

/** One read: each rectangle's values, then its cells' shown text and formulas. */
const readSteps = (c: LinksCase): Step[] => c.read.flatMap((g) => [{ do: 'values', sheet: g.sheet, range: g.range }, { do: 'cells', sheet: g.sheet, cells: cellsOf(g.range).addrs }])

function readOf(c: LinksCase, got: unknown[]): LinksRead {
  const out: LinksRead = { values: {}, formulas: {}, local: {}, shown: {} }
  eachRectCell(c.read, (key, i, r, j) => {
    const vals = got[i * 2] as WinCell[][]
    const cell = (got[i * 2 + 1] as string[][])[r * vals[0].length + j]
    const v = valOf(vals[r][j])
    if (v) out.values[key] = v
    const [text, , , , , formula, local] = cell
    if (formula.startsWith('=')) {
      out.formulas[key] = formula
      out.local[key] = local
    }
    out.shown[key] = text
  })
  return out
}

export async function stepLinks(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { built, files, sent } = await linksSet(set)
  const store = fileAnswers<LinksCase, LinksAnswer>(winFile(`links/${set}.json`), built.cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ ...metaOf(e, tag), seconds: secondsSince(t0) }, built)
  await fileJobs(
    `links ${set}`,
    `links-${set}`,
    store,
    files,
    (c) => [{ do: 'links' }, ...readSteps(c), { do: 'full' }, ...readSteps(c), { do: 'saveAs' }],
    (c, a) => {
      const [folder, sources] = a.answers[0] as [string, string[]]
      const n = c.read.length * 2
      return { sources, folder, reads: [readOf(c, a.answers.slice(1, 1 + n)), readOf(c, a.answers.slice(2 + n, 2 + 2 * n))] }
    },
    save,
  )
}
