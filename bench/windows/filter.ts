/**
 * The filter sets in Excel for Windows (`bench/coverage/filter/`), each case in its own book as the Mac step writes it.
 * Over SSH Excel has no screen, so the steps go through the object model: a typed or set text as `Formula2Local`, a
 * paste as text put on the clipboard and `Worksheet.Paste`, a copy as `Range.Copy` with a destination, Fill Down as
 * `Range.FillDown`, Delete as `ClearContents`; Undo has no object model of its own, so a case that undoes is the
 * Mac's alone. The criteria go on through `Range.AutoFilter`, and the AutoFilter's range is read after each step
 * (`AutoFilter.Range`). The answers go to `test/excel/win/filter/<set>.json`.
 */
import { writeFileSync } from 'node:fs'
import { writeXlsx } from '../../src/xlsx.js'
import { filterFile, savedFilter, type StoredFilter } from '../coverage/filter.js'
import { filterBook, type FilterCase, type FilterSet, type FilterStep } from '../coverage/filter/cases.js'
import { answeredLine, answerSaver, keptSame, loadStored, setModule } from '../coverage/stored.js'
import { inJobs, metaOf, savedInto, type Step, type Tag, type WinEnv } from './remote.js'

/** A step of a case as `xl.ps1` takes it. */
function winStep(s: FilterStep): Step {
  if ('type' in s || 'set' in s) return { do: 'set', sheet: 'S', cell: s.at, kind: 'l2', value: 'type' in s ? s.type : s.set }
  if ('paste' in s) return { do: 'paste', sheet: 'S', at: s.at, text: s.paste }
  if ('copy' in s) return { do: 'struct', op: 'copyRange', sheet: 'S', range: s.copy, target: 'S', to: s.to }
  if ('fill' in s) return { do: 'struct', op: 'fillDown', sheet: 'S', target: 'S', to: s.fill }
  if ('clear' in s) return { do: 'clear', sheet: 'S', range: s.clear, what: 'contents' }
  throw new Error('an undo has no step on Windows')
}

export async function stepFilter(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const cases = (await setModule<FilterSet>('filter', set)).buildFilterCases().filter((c) => c.steps.every((s) => !('undo' in s)))
  const file = filterFile(set, true)
  const answers = keptSame({ cases }, loadStored<StoredFilter>(file))
  const t0 = Date.now()
  const save = answerSaver(file, metaOf(e, tag), true, t0, cases, answers)
  const todo = cases.filter((c) => !answers[c.id])
  const results: [FilterCase, unknown[], Uint8Array][] = []
  const read: Step = { do: 'filterRange', sheet: 'S' }
  const book = async (c: FilterCase, f: string) => {
    const { data, issues } = await writeXlsx({ book: filterBook(c) }, { values: 'omit' })
    if (issues.length) throw new Error(`filter ${set} ${c.id}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
    writeFileSync(f, data)
    const criteria: Step[] = c.criteria ? [{ do: 'autoFilter', sheet: 'S', range: c.filter, field: c.criteria.field, value: c.criteria.value }] : []
    return { file: f, steps: [{ do: 'full' }, ...criteria, ...c.steps.flatMap((s) => [winStep(s), read]), { do: 'saveAs' }] }
  }
  await inJobs(`filter-${set}`, todo, book, savedInto(results, (c, why) => `filter ${set} ${c.id}: the book did not run: ${why}`), save)
  for (const [c, got, saved] of results) {
    // the answers: full, the criteria, then each step's (null or its error) and the range read after it, saveAs
    const at = c.criteria ? 2 : 1
    const errors = c.steps.map((_, k) => (got[at + 2 * k] as { error?: string } | null)?.error).filter(Boolean)
    if (errors.length) throw new Error(`filter ${set} ${c.id}: a step failed: ${errors.join('; ')}`)
    answers[c.id] = { after: c.steps.map((_, k) => got[at + 2 * k + 1] as string), saved: await savedFilter(saved) }
  }
  save()
  console.log(answeredLine(`filter ${set}`, todo.length, answers, cases))
}
