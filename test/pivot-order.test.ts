import { readFileSync } from 'node:fs'
import { expect, it } from 'vitest'
import { clockOptions } from '../src/result.js'
import type { PivotPhase } from '../bench/coverage/pivot/cases.js'
import { createState } from '../src/engine.js'
import { refreshPivots } from '../src/pivot/engine.js'
import { stateValues } from '../src/structure/values.js'
import { readXlsx } from '../src/xlsx.js'
import { rounded, shown } from '../bench/coverage/pivot/phases.js'

const root = new URL('./excel/win/pivot/', import.meta.url)
for (const set of ['corpus-shapes', 'order-probe', 'order-held']) {
  const answers = JSON.parse(readFileSync(new URL(`${set}.json`, root), 'utf8')).answers
  for (const [id, answer] of Object.entries(answers) as [string, { phases: PivotPhase[] }][])
    if (id.startsWith('ties-') || set === 'order-held')
      it(`sorts ${set}/${id} without reading Excel's resulting layout`, async () => {
        for (let k = 0; k < answer.phases.length; k++) {
          const suffix = k < answer.phases.length - 1 ? `-${k + 1}` : ''
          const b = (await readXlsx(readFileSync(new URL(`${set}/${id}${suffix}.xlsx`, root)))).book
          for (const sheet of b.sheets) for (const t of sheet.pivotTables ?? []) {
            delete t.rowLines
            delete t.columnLines
          }
          const values = stateValues(createState(b, clockOptions({ now: '2026-10-30', seed: 0 })))
          const r = refreshPivots(b, undefined, values, 0, false)
          expect(r.ok).toBe(true)
          if (!r.ok) return
          const want = answer.phases[k].pivots.PT1
          if ('error' in want) throw new Error(want.error)
          expect(rounded(shown(r.book, 'P', want.area)), `phase ${k}`).toEqual(rounded(want.values))
        }
      })
}
