/**
 * `readCells`, `getCell` and `getCells` answer as recorded (`read-answers.ts`): every field, order, page, cursor and
 * count, before and after edits, with display on and off, in a synchronous session, in the same thread and in a thread,
 * with closures and eval.
 */
import { readFileSync } from 'node:fs'
import { afterAll, describe, expect, it } from 'vitest'
import type { Execution } from '../src/index.js'
import { answers, BOOKS, CLOCK, FIXTURE, type Recorded } from './read-answers.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

afterAll(closeCalcs)

const recorded = JSON.parse(readFileSync(FIXTURE, 'utf8')) as Recorded

describe.each(['closures', 'eval'] as Execution[])('reading answers as recorded, execution %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    const open = openerFor(kind)
    it.each(Object.keys(BOOKS).flatMap((name) => [false, true].map((display) => [name, display] as const)))('book %s, display %s', async (name, display) => {
      const s = await open(BOOKS[name], { clock: CLOCK, execution, display })
      const got = await answers(s, name)
      await s.close()
      const want = recorded[name][String(display)]
      expect(got.map(([label]) => label)).toEqual(want.map(([label]) => label))
      for (let i = 0; i < want.length; i++) expect(got[i], want[i][0]).toEqual(want[i])
    }, 60_000)
  })
})
