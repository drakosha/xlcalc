/**
 * The filter-acts sets in Excel for Windows (`bench/coverage/filter-acts/`), each case in its own book as the Mac step
 * writes it. A button or Data > Filter needs a session on the screen, so `filter-acts.ps1` runs every case there
 * (`screen.ts`), the object model's steps too: Alt+Down and O sort a column largest first, Ctrl+Shift+L is Data > Filter;
 * criteria on a column's button are the Mac's alone, their values being in the check boxes of its window. The answers go
 * to `test/excel/win/filter-acts/<set>.json`.
 */
import { actsFile, type StoredActs } from '../coverage/filter-acts.js'
import { actFile, actSpec, savedFilter, type ActAnswer, type ActCase, type ActSet, type ActState } from '../coverage/filter-acts/cases.js'
import { answeredLine, answerSaver, keptSame, loadStored, setModule } from '../coverage/stored.js'
import { metaOf, type Tag, type WinEnv } from './remote.js'
import { savedOnScreen } from './screen.js'

const onWindows = (c: ActCase) => c.steps.every((s) => !('criteria' in s && s.via === 'button'))

export async function stepFilterActs(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const cases = (await setModule<ActSet>('filter-acts', set)).buildActCases().filter(onWindows)
  const file = actsFile(set, true)
  const answers = keptSame({ cases }, loadStored<StoredActs>(file))
  const t0 = Date.now()
  const save = answerSaver(file, metaOf(e, tag), true, t0, cases, answers)
  const only = process.env.FACTS_ONLY?.split(',')
  const todo = cases.filter((c) => !answers[c.id] && (!only || only.includes(c.id)))
  const books = new Map<string, Uint8Array>()
  for (const c of todo) books.set(c.id, await actFile(c))
  const byId = new Map(todo.map((c) => [c.id, c]))
  const done = await savedOnScreen('facts', 'filter-acts.ps1', books, (id) => actSpec(byId.get(id)!), 120 + 20 * todo.length)
  for (const [id, { saved, answer }] of done) {
    const { opened, after, dialog } = answer as { opened: ActState; after: ActState[]; dialog: string }
    const a: ActAnswer = { opened, after, saved: await savedFilter(saved), ...(byId.get(id)!.keptOff ? { written: await savedFilter(books.get(id)!) } : {}), ...(dialog ? { dialog } : {}) }
    answers[id] = a
    console.log(`${id}: opened ${opened.range}; ${after.map((s) => `${s.range} hidden ${s.hidden.join(',')} total ${JSON.stringify(s.total)}`).join(' / ')}; saved ${a.saved?.ref}`)
  }
  save()
  console.log(answeredLine(`filter-acts ${set}`, done.size, answers, cases))
}
