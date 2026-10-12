/**
 * Excel's Replace in Excel for Windows (`bench/coverage/replace/`): each case's book written by `writeXlsx`, opened,
 * computed, its autofilter applied (`Range.AutoFilter`) and its texts typed (`Formula2Local`), as the find step does;
 * then the case's Replace through `Range.Replace` over its range or the sheet's `Cells`, the book computed, its cells
 * read (`Text`, the format codes, `Value2`, the formulas, an array's and a spill's) and the book saved, its file read
 * back for what each cell holds. Each case runs in a book of its own. The answers go to
 * `test/excel/win/replace/[<tag>/]<set>.json`; a case whose book Excel did not run stops the run.
 */
import { writeFileSync } from 'node:fs'
import { writeXlsx } from '../../src/xlsx.js'
import { sentBook } from '../coverage/find/cases.js'
import { readOf, replaceFile, type ReplaceAnswer, type ReplaceCase, type ReplaceSet, type StoredReplace } from '../coverage/replace/cases.js'
import { answeredLine, inCaseOrder, keptSame, loadStored, secondsSince, setModule, writeAnswers } from '../coverage/stored.js'
import { savedCells, type RawCell } from '../coverage/ui.js'
import { inJobs, metaOf, type Step, type Tag, type WinEnv } from './remote.js'

export async function stepReplace(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { BOOKS, buildReplaceCases } = await setModule<ReplaceSet>('replace', set)
  const cases = buildReplaceCases()
  const file = replaceFile(set, tag)
  // a case is kept while it and its book are the same
  const answers = keptSame<ReplaceCase, ReplaceAnswer, Pick<StoredReplace, 'books' | 'cases'>>({ books: BOOKS, cases }, loadStored<StoredReplace>(file), (c, s) => JSON.stringify([c, s.books[c.book]]))
  const t0 = Date.now()
  const save = () => writeAnswers(file, { meta: { ...metaOf(e, tag), date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0) }, books: BOOKS, cases, answers: inCaseOrder(cases, answers) })
  const todo = cases.filter((c) => !answers[c.id])
  const results: [ReplaceCase, unknown[], Uint8Array][] = []
  const book = async (c: ReplaceCase, f: string) => {
    const b = BOOKS[c.book]
    const { data, issues } = await writeXlsx({ book: sentBook(b) }, { values: 'omit' })
    if (issues.length) throw new Error(`replace ${set} ${c.id}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
    writeFileSync(f, data)
    const filter: Step[] = b.filter ? [{ do: 'filter', sheet: 'S', range: b.filter.range, field: b.filter.field, criterion: b.filter.keep }] : []
    const typed = Object.entries(b.typed ?? {}).map(([cell, value]): Step => ({ do: 'set', sheet: 'S', cell, kind: 'l2', value }))
    return { file: f, steps: [{ do: 'full' }, ...filter, ...typed, { do: 'calc' }, { do: 'replace', sheet: 'S', ...c.query }, { do: 'calc' }, { do: 'cells', sheet: 'S', cells: readOf(b) }, { do: 'saveAs' }] }
  }
  await inJobs(`replace-${set}-${tag}`, todo, book, (c, a, cutShort) => {
    if (a && 'answers' in a && a.saved) {
      results.push([c, a.answers, a.saved])
      return []
    }
    if (!a && cutShort) return [c]
    throw new Error(`replace ${set} ${tag} ${c.id}: the book did not run: ${a && 'rejected' in a ? a.rejected : 'the job stopped on it'}`)
  }, save)
  for (const [c, got, saved] of results) {
    // the steps' answers: full, the filter, the typed texts, calc, replace, calc, cells, saveAs
    const at = got.length - 4
    const before = got.slice(0, at).find((x) => x && typeof x === 'object' && 'error' in x)
    if (before) throw new Error(`replace ${set} ${tag} ${c.id}: a step before Replace failed: ${JSON.stringify(before)}`)
    answers[c.id] = { replaced: got[at] as ReplaceAnswer['replaced'], cells: await savedCells(readOf(BOOKS[c.book]), got[at + 2] as RawCell[], saved) }
  }
  save()
  console.log(answeredLine(`replace ${set} ${tag}`, todo.length, answers, cases))
}
