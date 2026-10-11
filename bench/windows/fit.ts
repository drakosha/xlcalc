/**
 * Column widths in Excel for Windows: the Mac step's books (`fitBook`, `bench/coverage/fit.ts`), each case's column fitted
 * (`EntireColumn.AutoFit`) or its texts typed through the object model (`Value2` for `value`, `Formula2Local` for
 * `local`; the keyboard's cases are the Mac's alone), then each column's `ColumnWidth`, `Width` and `UseStandardWidth`
 * and what its cells show are read, and the book is saved, its `<col>` elements read back. The answers go to
 * `test/excel/win/fit/[<tag>/]<set>.json`.
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname } from 'node:path'
import { writeXlsx } from '../../src/xlsx.js'
import { indexToCol } from '../../src/model.js'
import { fitFile, keptAnswers, partAnswers } from '../coverage/fit.js'
import { caseColumn, caseRows, fitBook, type FitCase, type FitSet } from '../coverage/fit/cases.js'
import { inCaseOrder } from '../coverage/stored.js'
import { inJobs, metaOf, type Step, type Tag, type WinEnv } from './remote.js'

const PART = 60

export async function stepFit(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const cases = ((await import(`../coverage/fit/${set}.js`)) as FitSet).buildFitCases(tag).filter((c) => c.via !== 'keys')
  const file = fitFile(set, tag, true)
  const { answers } = keptAnswers(file, cases)
  const t0 = Date.now()
  const save = () => {
    mkdirSync(dirname(file), { recursive: true })
    const meta = { ...metaOf(e, tag), date: new Date().toISOString().slice(0, 10), seconds: Math.round((Date.now() - t0) / 100) / 10 }
    writeFileSync(file, JSON.stringify({ meta, cases, answers: inCaseOrder(cases, answers) }, null, 1) + '\n')
  }
  const todo = cases.filter((c) => !answers[c.id])
  const parts = Array.from({ length: Math.ceil(todo.length / PART) }, (_, k) => todo.slice(k * PART, (k + 1) * PART))
  const book = async (part: FitCase[], f: string) => {
    const { data, issues } = await writeXlsx({ book: fitBook(part) }, { values: 'omit' })
    if (issues.length) throw new Error(`fit ${set}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
    writeFileSync(f, data)
    const cols = part.map((_, k) => indexToCol(caseColumn(k)))
    const acts: Step[] = part.flatMap((c, k) => [
      ...(!c.typed || c.fitFirst ? [{ do: 'autofit', sheet: 'S', range: `${cols[k]}1` }] : []),
      ...(c.typed ?? []).map(([r, text]) => ({ do: 'set', sheet: 'S', cell: `${cols[k]}${r}`, kind: c.via === 'value' ? 's' : 'l2', value: text })),
    ])
    const cells = part.flatMap((c, k) => caseRows(c).map((r) => `${cols[k]}${r}`))
    return { file: f, steps: [...acts, { do: 'calc' }, { do: 'widths', sheet: 'S', columns: cols }, { do: 'shown', sheet: 'S', cells }, { do: 'saveAs' }] }
  }
  const results: [FitCase[], unknown[], Uint8Array][] = []
  await inJobs(`fit-${set}-${tag}`, parts, book, (part, a, cutShort) => {
    if (a && 'answers' in a && a.saved) {
      results.push([part, a.answers, a.saved])
      return []
    }
    if (!a && cutShort) return [part]
    throw new Error(`fit ${set} ${tag}: a part did not run: ${a && 'rejected' in a ? a.rejected : 'the job stopped on it'}`)
  }, save)
  for (const [part, got, saved] of results) {
    // the answers of the steps: autofit and set (null, or the error a set got), calc, widths, shown, saveAs
    let i = 0
    const refused = part.map((c) => {
      if (!c.typed || c.fitFirst) i++
      return (c.typed ?? []).map(() => {
        const x = got[i++]
        return x && typeof x === 'object' && 'error' in x ? String((x as { error: string }).error) : ''
      })
    })
    const widths = got[i + 1] as [number, number, boolean][]
    const shown = (got[i + 2] as [string, unknown][]).map(([t]) => t)
    let at = 0
    const read = part.map((c, k): [number, number, boolean, string[]] => {
      const rows = caseRows(c)
      const texts = shown.slice(at, (at += rows.length))
      return [widths[k][0], widths[k][1], widths[k][2], texts]
    })
    Object.assign(answers, await partAnswers(part, refused, read, saved))
  }
  save()
  console.log(`fit ${set} ${tag}: ${todo.length} cases run, ${Object.keys(answers).length} of ${cases.length} answered`)
}
