/**
 * `setCells`, `undo` and `redo` answer as recorded (`edit-answers.ts`; over many dependents `recompute-answers.ts`), and
 * leave the book as recorded, with display on and off, in a synchronous session, in the same thread and in a thread, with
 * closures and eval.
 */
import { it } from 'vitest'
import { BOOK, CLOCK, FIXTURE, series, STEPS, type Recorded } from './edit-answers.js'
import * as many from './recompute-answers.js'
import { expectRecorded, inEverySession, readRecorded } from './recorded.js'

for (const [title, book, steps, fixture] of [
  ['editing answers as recorded', BOOK, STEPS, FIXTURE],
  ['recalculation answers as recorded', many.BOOK, many.STEPS, many.FIXTURE],
] as const) {
  const recorded = readRecorded<Recorded>(fixture)
  inEverySession(title, (execution, open) => {
    it.each([false, true])('display %s', async (display) => {
      const s = await open(book, { clock: CLOCK, execution, display })
      const got = await series(s, steps)
      await s.close()
      expectRecorded(got, recorded[String(display)])
    }, 60_000)
  })
}
