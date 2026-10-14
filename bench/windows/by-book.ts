/**
 * The steps of Excel for Windows whose cases run grouped by book (`end`, `find`): each book of a set written by
 * `writeXlsx`, opened, computed, its autofilter applied (`Range.AutoFilter`) and its texts typed (`Formula2Local`), then
 * each case's own step. The cases of one book run in one file; a case is kept while it and its book are the same.
 */
import { writeFileSync } from 'node:fs'
import type { BookInput } from '../../src/contract.js'
import { writeXlsx } from '../../src/xlsx.js'
import { answeredLine, inCaseOrder, keptSame, loadStored, secondsSince, writeAnswers } from '../coverage/stored.js'
import { inJobs, metaOf, type Step, type Tag, type WinEnv } from './remote.js'

/** A book of such a set: its filter and the texts typed into it beside what `sent` writes. */
interface ByBook {
  filter?: { range: string; field: number; keep: string }
  typed?: Record<string, string>
}

export interface ByBookStep<C, A, B> {
  /** The step's name and the set's, in labels and errors (`end`, `find`). */
  kind: string
  set: string
  tag: Tag
  e: WinEnv
  books: Record<string, B>
  cases: C[]
  /** The file of the stored answers. */
  file: string
  /** The book as Excel gets it, before its filter and its typed texts. */
  sent: (b: B) => BookInput
  /** A case's own step, and its answer from what `xl.ps1` gave. */
  step: (c: C) => Step
  answer: (x: unknown) => A
}

export async function stepByBook<C extends { id: string; book: string }, A, B extends ByBook>(o: ByBookStep<C, A, B>): Promise<void> {
  const { kind, set, tag, books, cases, file } = o
  const answers = keptSame<C, A, { books: Record<string, B>; cases: C[] }>({ books, cases }, loadStored<{ books: Record<string, B>; cases: C[]; answers: Record<string, A> }>(file), (c, s) => JSON.stringify([c, s.books[c.book]]))
  const t0 = Date.now()
  const save = () => writeAnswers(file, { meta: { ...metaOf(o.e, tag), date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0) }, books, cases, answers: inCaseOrder(cases, answers) })
  const todo = cases.filter((c) => !answers[c.id])
  const parts = Object.keys(books).map((b) => todo.filter((c) => c.book === b)).filter((p) => p.length)
  const book = async (part: C[], f: string) => {
    const b = books[part[0].book]
    const { data, issues } = await writeXlsx({ book: o.sent(b) }, { values: 'omit' })
    if (issues.length) throw new Error(`${kind} ${set} ${part[0].book}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
    writeFileSync(f, data)
    const filter: Step[] = b.filter ? [{ do: 'filter', sheet: 'S', range: b.filter.range, field: b.filter.field, criterion: b.filter.keep }] : []
    const typed = Object.entries(b.typed ?? {}).map(([cell, value]): Step => ({ do: 'set', sheet: 'S', cell, kind: 'l2', value }))
    return { file: f, steps: [{ do: 'full' }, ...filter, ...typed, { do: 'calc' }, ...part.map(o.step)] }
  }
  await inJobs(`${kind}-${set}-${tag}`, parts, book, (part, a, cutShort) => {
    if (a && 'answers' in a) {
      const before = a.answers.length - part.length
      const failed = a.answers.slice(0, before).find((x) => x && typeof x === 'object' && 'error' in x)
      if (failed) throw new Error(`${kind} ${set} ${tag} ${part[0].book}: a step before the cases failed: ${JSON.stringify(failed)}`)
      part.forEach((c, i) => (answers[c.id] = o.answer(a.answers[before + i])))
      return []
    }
    if (!a && cutShort) return [part]
    throw new Error(`${kind} ${set} ${tag} ${part[0].book}: the book did not run: ${a && 'rejected' in a ? a.rejected : 'the job stopped on it'}`)
  }, save)
  save()
  console.log(answeredLine(`${kind} ${set} ${tag}`, todo.length, answers, cases))
}
