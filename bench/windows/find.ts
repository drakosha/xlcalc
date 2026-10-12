/**
 * Excel's Find in Excel for Windows (`bench/coverage/find/`): each book of a set written by `writeXlsx`, opened, computed,
 * its autofilter applied (`Range.AutoFilter`) and its texts typed (`Formula2Local`); then each case's query through
 * `Range.Find` from after the range's last cell and `FindNext` until Find comes back to the first cell, at most
 * `MAX_FOUND` cells, each with its `Formula`, `FormulaLocal` and `Text`. The cases of one book run in one file. The
 * answers go to `test/excel/win/find/[<tag>/]<set>.json`; a case Find refuses keeps its error.
 */
import { writeFileSync } from 'node:fs'
import { writeXlsx } from '../../src/xlsx.js'
import { MAX_FOUND, findFile, sentBook, type FindAnswer, type FindCase, type FindSet, type StoredFind } from '../coverage/find/cases.js'
import { answeredLine, inCaseOrder, keptSame, loadStored, secondsSince, setModule, writeAnswers } from '../coverage/stored.js'
import { inJobs, metaOf, type Step, type Tag, type WinEnv } from './remote.js'

export async function stepFind(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { BOOKS, buildFindCases } = await setModule<FindSet>('find', set)
  const cases = buildFindCases()
  const file = findFile(set, tag)
  // a case is kept while it and its book are the same
  const answers = keptSame<FindCase, FindAnswer, Pick<StoredFind, 'books' | 'cases'>>({ books: BOOKS, cases }, loadStored<StoredFind>(file), (c, s) => JSON.stringify([c, s.books[c.book]]))
  const t0 = Date.now()
  const save = () => writeAnswers(file, { meta: { ...metaOf(e, tag), date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0) }, books: BOOKS, cases, answers: inCaseOrder(cases, answers) })
  const todo = cases.filter((c) => !answers[c.id])
  const parts = Object.keys(BOOKS).map((b) => todo.filter((c) => c.book === b)).filter((p) => p.length)
  const book = async (part: FindCase[], f: string) => {
    const b = BOOKS[part[0].book]
    const { data, issues } = await writeXlsx({ book: sentBook(b) }, { values: 'omit' })
    if (issues.length) throw new Error(`find ${set} ${part[0].book}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
    writeFileSync(f, data)
    const filter: Step[] = b.filter ? [{ do: 'filter', sheet: 'S', range: b.filter.range, field: b.filter.field, criterion: b.filter.keep }] : []
    const typed = Object.entries(b.typed ?? {}).map(([cell, value]): Step => ({ do: 'set', sheet: 'S', cell, kind: 'l2', value }))
    const finds = part.map((c): Step => ({ do: 'find', sheet: 'S', max: MAX_FOUND, ...c.query }))
    return { file: f, steps: [{ do: 'full' }, ...filter, ...typed, { do: 'calc' }, ...finds] }
  }
  await inJobs(`find-${set}-${tag}`, parts, book, (part, a, cutShort) => {
    if (a && 'answers' in a) {
      const before = a.answers.length - part.length
      const typed = a.answers.slice(0, before).find((x) => x && typeof x === 'object' && 'error' in x)
      if (typed) throw new Error(`find ${set} ${tag} ${part[0].book}: a step before the queries failed: ${JSON.stringify(typed)}`)
      part.forEach((c, i) => (answers[c.id] = answerOf(a.answers[before + i])))
      return []
    }
    if (!a && cutShort) return [part]
    throw new Error(`find ${set} ${tag} ${part[0].book}: the book did not run: ${a && 'rejected' in a ? a.rejected : 'the job stopped on it'}`)
  }, save)
  save()
  console.log(answeredLine(`find ${set} ${tag}`, todo.length, answers, cases))
}

/** A query's answer as `xl.ps1` gives it: `[[address, Formula, FormulaLocal, Text]...], more, FindNext's error`, or Find's error. */
function answerOf(x: unknown): FindAnswer {
  if (x && typeof x === 'object' && 'error' in x) return { error: String(x.error) }
  const [found, more, next] = x as [[string, string, string, string][], boolean, string]
  return { found: found.map(([cell, formula, formulaLocal, text]) => ({ cell, formula, formulaLocal, text })), ...(more ? { more: true as const } : {}), ...(next ? { next } : {}) }
}
