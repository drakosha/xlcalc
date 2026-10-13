/**
 * Excel's `Range.End` in Excel for Windows (`bench/coverage/end/`): each book of a set written by `writeXlsx`, opened,
 * computed, its autofilter applied (`Range.AutoFilter`) and its texts typed (`Formula2Local`); then each case's cell's
 * `End` in its direction (`xlUp`, `xlDown`, `xlToLeft`, `xlToRight`), the address of the cell it comes to. The cases of
 * one book run in one file. The answers go to `test/excel/win/end/[<tag>/]<set>.json`; a case Excel refuses keeps its error.
 */
import { writeFileSync } from 'node:fs'
import type { EndDirection } from '../../src/contract.js'
import { writeXlsx } from '../../src/xlsx.js'
import { endFile, sentBook, type EndAnswer, type EndCase, type EndSet, type StoredEnd } from '../coverage/end/cases.js'
import { answeredLine, inCaseOrder, keptSame, loadStored, secondsSince, setModule, writeAnswers } from '../coverage/stored.js'
import { inJobs, metaOf, type Step, type Tag, type WinEnv } from './remote.js'

/** Excel's `XlDirection`. */
const XL_DIRECTION: Record<EndDirection, number> = { up: -4162, down: -4121, left: -4159, right: -4161 }

export async function stepEnd(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { BOOKS, buildEndCases } = await setModule<EndSet>('end', set)
  const cases = buildEndCases()
  const file = endFile(set, tag)
  // a case is kept while it and its book are the same
  const answers = keptSame<EndCase, EndAnswer, Pick<StoredEnd, 'books' | 'cases'>>({ books: BOOKS, cases }, loadStored<StoredEnd>(file), (c, s) => JSON.stringify([c, s.books[c.book]]))
  const t0 = Date.now()
  const save = () => writeAnswers(file, { meta: { ...metaOf(e, tag), date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0) }, books: BOOKS, cases, answers: inCaseOrder(cases, answers) })
  const todo = cases.filter((c) => !answers[c.id])
  const parts = Object.keys(BOOKS).map((b) => todo.filter((c) => c.book === b)).filter((p) => p.length)
  const book = async (part: EndCase[], f: string) => {
    const b = BOOKS[part[0].book]
    const { data, issues } = await writeXlsx({ book: sentBook(b) }, { values: 'omit' })
    if (issues.length) throw new Error(`end ${set} ${part[0].book}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
    writeFileSync(f, data)
    const filter: Step[] = b.filter ? [{ do: 'filter', sheet: 'S', range: b.filter.range, field: b.filter.field, criterion: b.filter.keep }] : []
    const typed = Object.entries(b.typed ?? {}).map(([cell, value]): Step => ({ do: 'set', sheet: 'S', cell, kind: 'l2', value }))
    const ends = part.map((c): Step => ({ do: 'end', sheet: 'S', cell: c.cell, direction: XL_DIRECTION[c.direction] }))
    return { file: f, steps: [{ do: 'full' }, ...filter, ...typed, { do: 'calc' }, ...ends] }
  }
  await inJobs(`end-${set}-${tag}`, parts, book, (part, a, cutShort) => {
    if (a && 'answers' in a) {
      const before = a.answers.length - part.length
      const failed = a.answers.slice(0, before).find((x) => x && typeof x === 'object' && 'error' in x)
      if (failed) throw new Error(`end ${set} ${tag} ${part[0].book}: a step before the cases failed: ${JSON.stringify(failed)}`)
      part.forEach((c, i) => {
        const x = a.answers[before + i]
        answers[c.id] = x && typeof x === 'object' && 'error' in x ? { error: String(x.error) } : { cell: String(x) }
      })
      return []
    }
    if (!a && cutShort) return [part]
    throw new Error(`end ${set} ${tag} ${part[0].book}: the book did not run: ${a && 'rejected' in a ? a.rejected : 'the job stopped on it'}`)
  }, save)
  save()
  console.log(answeredLine(`end ${set} ${tag}`, todo.length, answers, cases))
}
