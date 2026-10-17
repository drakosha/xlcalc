/**
 * `readCells`, `getCell` and `getCells` answer as recorded (`read-answers.ts`): every field, order, page, cursor and
 * count, before and after edits, with display on and off, in a synchronous session, in the same thread and in a thread,
 * with closures and eval.
 */
import { it } from 'vitest'
import { answers, BOOKS, CLOCK, FIXTURE, type Recorded } from './read-answers.js'
import { expectRecorded, inEverySession, readRecorded } from './recorded.js'

const recorded = readRecorded<Recorded>(FIXTURE)

inEverySession('reading answers as recorded', (execution, open) => {
  it.each(Object.keys(BOOKS).flatMap((name) => [false, true].map((display) => [name, display] as const)))('book %s, display %s', async (name, display) => {
    const s = await open(BOOKS[name], { clock: CLOCK, execution, display })
    const got = await answers(s, name)
    await s.close()
    expectRecorded(got, recorded[name][String(display)])
  }, 60_000)
})
