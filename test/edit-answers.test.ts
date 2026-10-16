/**
 * `setCells`, `undo` and `redo` answer as recorded (`edit-answers.ts`), and leave the book as recorded, with display on and
 * off, in a synchronous session, in the same thread and in a thread, with closures and eval.
 */
import { readFileSync } from 'node:fs'
import { afterAll, describe, expect, it } from 'vitest'
import type { Execution } from '../src/index.js'
import { answers, BOOK, CLOCK, FIXTURE, type Recorded } from './edit-answers.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

afterAll(closeCalcs)

const recorded = JSON.parse(readFileSync(FIXTURE, 'utf8')) as Recorded

describe.each(['closures', 'eval'] as Execution[])('editing answers as recorded, execution %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    const open = openerFor(kind)
    it.each([false, true])('display %s', async (display) => {
      const s = await open(BOOK, { clock: CLOCK, execution, display })
      const got = await answers(s)
      await s.close()
      const want = recorded[String(display)]
      expect(got.map(([label]) => label)).toEqual(want.map(([label]) => label))
      for (let i = 0; i < want.length; i++) expect(got[i], want[i][0]).toEqual(want[i])
    }, 60_000)
  })
})
