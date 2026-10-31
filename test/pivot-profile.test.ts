import { readFileSync } from 'node:fs'
import { afterEach, expect, it } from 'vitest'
import { createInlineCalc, createSession, registerPivot, type Execution } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'
import { readXlsx } from '../src/xlsx.js'
import { closedAfter } from './fixtures/pool.js'

const { made } = closedAfter()
afterEach(() => registerPivot(pivotEngine))

it('checks the sorting-thread profile', () => {
  for (const sortThreads of [0, -1, 1.5, Infinity, NaN, null, '2', 2 ** 53])
    expect(() => registerPivot(pivotEngine, { sortThreads } as never)).toThrow('options.sortThreads')
  expect(() => registerPivot(pivotEngine, { threads: 2 } as never)).toThrow('unknown field "threads"')
})

for (const execution of ['closures', 'eval'] as Execution[])
  for (const kind of ['sync', 'inline', 'pool'])
    it(`changes the sort profile in an existing ${kind} session (${execution})`, async () => {
      registerPivot(pivotEngine, { sortThreads: 1 })
      const calc = kind === 'sync' ? null : kind === 'inline' ? createInlineCalc() : await made({ threads: 1 })
      const book = (await readXlsx(readFileSync(new URL('./excel/win/pivot/order-held/items-513.xlsx', import.meta.url)))).book
      const session = calc ? await calc.createSession(book, { execution }) : createSession(book, { execution })
      for (const [sortThreads, first] of [[1, 'Item 0256'], [2, 'Item 0128'], [1, 'Item 0256']] as const) {
        registerPivot(pivotEngine, { sortThreads })
        expect((await session.editStructure({ op: 'refreshPivots' })).ok).toBe(true)
        expect(await session.getValues('P', 'A4')).toEqual([[first]])
      }
      await session.close()
    })
