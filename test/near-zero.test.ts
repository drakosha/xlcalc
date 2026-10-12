/**
 * A sum or difference near 0 against Excel without the acceptance tests' tolerance (the sets `near-zero-held` and
 * `sum-zero-held`, Excel for Mac's answers and Excel for Windows's): those compare numbers within 1e-9·max(1, |x|),
 * which takes 5.55E-17 for 0, so here a number Excel answered 0 is 0 in the engine, any other is within 1e-9 of Excel's
 * relatively, in both execution modes.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { Val } from '../bench/coverage/formulajs.js'
import { compute, type Execution } from '../src/index.js'
import { CLOCK, toVal } from './excel-cases.js'
import { caseBook, loadSet } from './excel-sets.js'

const SETS = ['near-zero-held', 'sum-zero-held']

for (const name of SETS) {
  const mod = await loadSet(name)
  const probes = mod.buildProbes()
  for (const execution of ['closures', 'eval'] as Execution[]) {
    const cells = compute(caseBook(probes, mod.SHEETS), { clock: CLOCK, execution }).sheets.at(-1)!.cells
    for (const [excel, dir] of [['Excel for Mac', 'cases'], ['Excel for Windows', 'win/cases']]) {
      const file = new URL(`./excel/${dir}/${name}.json`, import.meta.url)
      describe.skipIf(!existsSync(file))(`${name} against ${excel} exactly, ${execution}`, () => {
        it('a number is 0 where Excel answered 0, else within 1e-9 of Excel relatively', () => {
          const values = (JSON.parse(readFileSync(file, 'utf8')) as { values: Record<string, Val> }).values
          const off = probes.flatMap((p, i) => {
            const x = values[p.id]
            const v = toVal(cells[`A${i + 1}`].value)
            if (x.t !== 'n' || v.t !== 'n') return []
            const ok = x.v === 0 ? v.v === 0 : Math.abs(v.v - x.v) <= 1e-9 * Math.abs(x.v)
            return ok ? [] : [`${p.id} ${p.formula}: Excel ${x.v}, engine ${v.v}`]
          })
          expect(off).toEqual([])
        })
      })
    }
  }
}
