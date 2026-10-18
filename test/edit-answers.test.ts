/**
 * `setCells`, `editStructure`, `undo` and `redo` answer as recorded (`edit-answers.ts`; over many dependents
 * `recompute-answers.ts`; structure operations `structure-answers.ts`), and leave the book as recorded, with display on
 * and off, in a synchronous session, in the same thread and in a thread, with closures and eval.
 */
import { it } from 'vitest'
import { BOOK, CLOCK, FIXTURE, series, STEPS, type Recorded, type SeriesOptions, type Step } from './edit-answers.js'
import * as many from './recompute-answers.js'
import * as moved from './structure-answers.js'
import { expectRecorded, inEverySession, readRecorded } from './recorded.js'
import type { BookInput } from '../src/index.js'

const SERIES: [string, BookInput, Step[], string, SeriesOptions?][] = [
  ['editing answers as recorded', BOOK, STEPS, FIXTURE],
  ['recalculation answers as recorded', many.BOOK, many.STEPS, many.FIXTURE],
  ['structure answers as recorded', moved.BOOK, moved.STEPS, moved.FIXTURE, moved.OPTIONS],
]

for (const [title, book, steps, fixture, options] of SERIES) {
  const recorded = readRecorded<Recorded>(fixture)
  inEverySession(title, (execution, open) => {
    it.each([false, true])('display %s', async (display) => {
      const s = await open(book, { clock: CLOCK, execution, display })
      const got = await series(s, steps, options)
      await s.close()
      expectRecorded(got, recorded[String(display)])
    }, 60_000)
  })
}
