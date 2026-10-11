/**
 * Typed text in Excel for Windows: the Mac step's books (`entryBook`, `bench/coverage/entry.ts`), each case's text typed
 * into its cell through `Formula2Local`, as Excel 365 takes typed text in the user's language; then what each cell shows,
 * its format codes, its prefix character and its value (`Value2`) are read (for a set of `edit` its `Formula` and
 * `FormulaLocal` too), and the book is saved, its file read back for what each cell holds. The answers go to `test/excel/win/entry/[<tag>/]<set>.json`, with Excel's automatic percent
 * entry in `meta`.
 */
import { writeFileSync } from 'node:fs'
import { inParts } from '../coverage/book.js'
import { ENTRY_PART, entryFile, partAnswers, type StoredEntry } from '../coverage/entry.js'
import { entryBook, type EntryCase, type EntrySet } from '../coverage/entry/cases.js'
import { answeredLine, inCaseOrder, keptSame, loadStored, secondsSince, setModule, writeAnswers } from '../coverage/stored.js'
import { writeXlsx } from '../../src/xlsx.js'
import { inJobs, metaOf, type Tag, type WinEnv } from './remote.js'

export async function stepEntry(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const mod = await setModule<EntrySet>('entry', set)
  const cases = mod.buildEntryCases()
  const file = entryFile(set, tag, true)
  const prev = loadStored<StoredEntry>(file)
  const answers = keptSame({ cases }, prev)
  let autoPercent: unknown = prev?.meta.autoPercentEntry
  const t0 = Date.now()
  const save = () => {
    const meta = { ...metaOf(e, tag), date: new Date().toISOString().slice(0, 10), autoPercentEntry: String(autoPercent), seconds: secondsSince(t0) }
    writeAnswers(file, { meta, cases, answers: inCaseOrder(cases, answers) })
  }
  const todo = cases.filter((c) => !answers[c.id])
  const parts = inParts(todo, ENTRY_PART)
  const book = async (part: EntryCase[], f: string) => {
    const { data, issues } = await writeXlsx({ book: entryBook(part) }, { values: 'omit' })
    if (issues.length) throw new Error(`entry ${set}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
    writeFileSync(f, data)
    const cells = part.map((_, i) => `A${i + 1}`)
    const typed = part.map((c, i) => ({ do: 'set', sheet: 'S', cell: cells[i], kind: 'l2', value: c.text }))
    return { file: f, steps: [...typed, { do: 'calc' }, { do: 'entry', sheet: 'S', cells, ...(mod.edit ? { edit: true } : {}) }, { do: 'autoPercent' }, { do: 'saveAs' }, { do: 'values', sheet: 'S', range: `A1:A${part.length}` }] }
  }
  const results: [EntryCase[], string[], [string, string, string, string, unknown, string?, string?][], Uint8Array][] = []
  await inJobs(`entry-${set}-${tag}`, parts, book, (part, a, cutShort) => {
    if (a && 'answers' in a && a.saved) {
      const n = part.length
      const refused = a.answers.slice(0, n).map((x) => (x && typeof x === 'object' && 'error' in x ? String((x as { error: string }).error) : ''))
      autoPercent = a.answers[n + 2]
      const values = a.answers[n + 4] as unknown[][]
      results.push([part, refused, (a.answers[n + 1] as [string, string, string, string, string?, string?][]).map(([text, code, local, prefix, ...edit], i) => [text, code, local, prefix, values[i][0], ...edit]), a.saved])
      return []
    }
    if (!a && cutShort) return [part]
    throw new Error(`entry ${set} ${tag}: a part did not run: ${a && 'rejected' in a ? a.rejected : 'the job stopped on it'}`)
  }, save)
  for (const [part, refused, shown, saved] of results) Object.assign(answers, await partAnswers(part, refused, shown, saved))
  save()
  console.log(answeredLine(`entry ${set} ${tag}`, todo.length, answers, cases))
}
