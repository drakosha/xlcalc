/**
 * The functions' arguments without their implementations (`src/function-arity.ts`) are the engine's own: the file is the
 * one `test/function-arity.ts` makes from the definitions, and each function's arity is its spec's.
 */
import { readFileSync } from 'node:fs'
import { expect, it } from 'vitest'
import { arityOf } from '../src/arity.js'
import { FUNCTION_ARITY } from '../src/function-arity.js'
import { FUNCTIONS } from '../src/functions.js'
import { functionArity } from './function-arity.js'

it('src/function-arity.ts is made from the definitions (npx tsx test/function-arity.ts --write)', () => {
  expect(readFileSync(new URL('../src/function-arity.ts', import.meta.url), 'utf8')).toBe(functionArity())
})

it("each function's arity is the engine's: the counts it takes and the arguments it lets a formula leave out", () => {
  expect(Object.keys(FUNCTION_ARITY).sort()).toEqual(Object.keys(FUNCTIONS).sort())
  for (const [name, spec] of Object.entries(FUNCTIONS)) {
    const a = arityOf(FUNCTION_ARITY[name])
    expect([a.min, a.max, a.counts], name).toEqual([spec.min, spec.max, spec.counts])
    for (let k = 0; k <= 256; k++) expect([k, a.takes(k), a.omits(k)], name).toEqual([k, spec.takes(k), spec.omits(k)])
  }
})
