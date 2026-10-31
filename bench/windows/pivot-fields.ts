/**
 * Fields placed from the PivotTable Fields list of Excel for Windows (`bench/coverage/pivot/checked.ts`): the list is
 * there only in a session on a screen, so `pivot-fields.ps1` runs as a task there (`screen.ts`), each case's book with
 * its fields checked by UI Automation, and saves it. The answer is the PivotTable the saved book holds (`checkedOf`),
 * or the step's error; the answers go to `test/excel/win/pivot-fields/<set>.json`, the books beside them
 * (`<set>/<id>.xlsx`). Without a session on the screen the step stops.
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { readXlsx, writeXlsx } from '../../src/xlsx.js'
import { checkedOf, type CheckedAnswer, type CheckedSet } from '../coverage/pivot/checked.js'
import { DATA, sourceBook, sourceText } from '../coverage/pivot/data.js'
import { setModule, writeAnswers } from '../coverage/stored.js'
import { winFile } from './functions.js'
import { metaOf, type Tag, type WinEnv } from './remote.js'
import { savedOnScreen } from './screen.js'

export async function stepPivotFields(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const cases = (await setModule<CheckedSet>('pivot', set)).buildCheckedCases()
  const books = new Map<string, Uint8Array>()
  for (const c of cases) books.set(c.id, (await writeXlsx({ book: sourceBook(DATA[c.data ?? 'kinds'](), 'range') }, { values: 'omit' })).data)
  const byId = new Map(cases.map((c) => [c.id, c]))
  const t0 = Date.now()
  const done = await savedOnScreen('pivotfields', 'pivot-fields.ps1', books, (id) => {
    const c = byId.get(id)!
    return { source: sourceText('range', DATA[c.data ?? 'kinds']()), fields: c.fields, ...c.under, ...(c.answer ? { answer: c.answer } : {}) }
  }, 120 + 40 * cases.length)
  const file = winFile(`pivot-fields/${set}.json`)
  const dir = file.slice(0, -'.json'.length)
  mkdirSync(dir, { recursive: true })
  const answers: Record<string, CheckedAnswer | { error: string }> = {}
  for (const c of cases) {
    const got = done.get(c.id)
    if (!got) {
      answers[c.id] = { error: 'not placed (its error above)' }
      continue
    }
    writeFileSync(join(dir, `${c.id}.xlsx`), got.saved)
    answers[c.id] = { ...checkedOf((await readXlsx(got.saved)).book, c.watch), ...(got.answer.dialog ? { dialog: got.answer.dialog } : {}) }
  }
  writeAnswers(file, { meta: { ...metaOf(e, tag), date: new Date().toISOString().slice(0, 10), seconds: Math.round((Date.now() - t0) / 1000) }, cases, answers })
  console.log(`pivot-fields ${set}: ${Object.values(answers).filter((a) => !('error' in a)).length} of ${cases.length} placed`)
}
