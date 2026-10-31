import { readFileSync } from 'node:fs'
import { expect, it } from 'vitest'
import { createInlineCalc, createSession, registerPivot, type Execution } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'
import { readXlsx } from '../src/xlsx.js'
import { closedAfter } from './fixtures/pool.js'

const { made } = closedAfter()

for (const execution of ['closures', 'eval'] as Execution[])
  for (const kind of ['sync', 'inline', 'pool'])
    it(`sorts more than 512 items in Excel's two parts in a ${kind} session (${execution})`, async () => {
      registerPivot(pivotEngine)
      const calc = kind === 'sync' ? null : kind === 'inline' ? createInlineCalc() : await made({ threads: 1 })
      const book = (await readXlsx(readFileSync(new URL('./excel/win/pivot/order-held/items-513.xlsx', import.meta.url)))).book
      const session = calc ? await calc.createSession(book, { execution }) : createSession(book, { execution })
      expect((await session.editStructure({ op: 'refreshPivots' })).ok).toBe(true)
      expect(await session.getValues('P', 'A4')).toEqual([['Item 0128']])
      await session.close()
    })
