/**
 * `setCells`, `undo` and `redo` answer as recorded (`edit-answers.ts`), and leave the book as recorded, with display on and
 * off, in a synchronous session, in the same thread and in a thread, with closures and eval.
 */
import { it } from 'vitest'
import { answers, BOOK, CLOCK, FIXTURE, type Recorded } from './edit-answers.js'
import { expectRecorded, inEverySession, readRecorded } from './recorded.js'

const recorded = readRecorded<Recorded>(FIXTURE)

inEverySession('editing answers as recorded', (execution, open) => {
  it.each([false, true])('display %s', async (display) => {
    const s = await open(BOOK, { clock: CLOCK, execution, display })
    const got = await answers(s)
    await s.close()
    expectRecorded(got, recorded[String(display)])
  }, 60_000)
})
