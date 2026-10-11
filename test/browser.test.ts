/**
 * Outside Node: the library bundled by esbuild for the browser has no dependencies and no Node built-ins,
 * and the bundle runs in an empty `node:vm` context (only the ECMAScript globals: no `process`, `Buffer`,
 * `require`, `performance` or `crypto`) and gives the same hash as Node in both execution modes. With code
 * generation from strings disabled, as under a Content Security Policy without 'unsafe-eval', the bundle
 * computes in the default `closures` mode with the same hash, and `execution: 'eval'` fails with a clear
 * EvalError.
 */
import { createContext, runInContext } from 'node:vm'
import { describe, expect, it } from 'vitest'
import { compute } from '../src/index.js'
import { PACKAGE_NAME } from '../src/config.js'
import { synthBook } from '../src/synth.js'
import { bundle } from './fixtures/bundle.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

describe('outside Node', () => {
  it('the browser bundle has no Node built-ins and runs in an empty vm context with the same hash', async () => {
    const { code, inputs } = await bundle()
    expect(inputs.filter((p) => !/^(src|dist)\/|<stdin>/.test(p))).toEqual([])
    expect(code).not.toMatch(/\brequire\s*\(|\bprocess\.|\bBuffer\b|["']node:/)

    const ctx = createContext({})
    expect(runInContext('[typeof process, typeof require, typeof Buffer, typeof performance, typeof crypto, typeof TextEncoder]', ctx)).toEqual(Array(6).fill('undefined'))
    runInContext(code, ctx)
    const out = runInContext(
      `const s = FormulaEngine.synthBook({ seed: 1, rows: 340 })
       const session = FormulaEngine.createSession(s.book, { clock: ${JSON.stringify(clock)} })
       const evalSession = FormulaEngine.createSession(s.book, { clock: ${JSON.stringify(clock)}, execution: 'eval' })
       JSON.stringify({
         hash: FormulaEngine.compute(s.book, { clock: ${JSON.stringify(clock)} }).hash,
         eval: FormulaEngine.compute(s.book, { clock: ${JSON.stringify(clock)}, execution: 'eval' }).hash,
         session: session.getResult().hash,
         evalSession: evalSession.getResult().hash,
         seed: typeof FormulaEngine.createSession(s.book, { live: true }).clock.seed,
       })`,
      ctx,
    ) as string
    const expected = compute(synthBook({ seed: 1, rows: 340 }).book, { clock }).hash
    expect(JSON.parse(out)).toEqual({ hash: expected, eval: expected, session: expected, evalSession: expected, seed: 'number' })
  }, 60_000)

  it('without code generation from strings (strict CSP) the bundle computes in the closures mode, and the eval mode explains the failure', async () => {
    const { code } = await bundle()
    const ctx = createContext({}, { codeGeneration: { strings: false } })
    expect(() => runInContext('new Function("return 1")', ctx)).toThrow()
    runInContext(code, ctx)
    expect(runInContext('FormulaEngine.validateFormula("SUM(A1:A3)*2").ok', ctx)).toBe(true)
    expect(runInContext('FormulaEngine.toLocal("SUM(A1:A3,1.5)", FormulaEngine.ruRU).formula', ctx)).toBe('СУММ(A1:A3;1,5)')
    const out = runInContext(
      `const s = FormulaEngine.synthBook({ seed: 1, rows: 340 })
       const session = FormulaEngine.createSession(s.book, { clock: ${JSON.stringify(clock)} })
       JSON.stringify({ hash: FormulaEngine.compute(s.book, { clock: ${JSON.stringify(clock)} }).hash, session: session.getResult().hash })`,
      ctx,
    ) as string
    const expected = compute(synthBook({ seed: 1, rows: 340 }).book, { clock }).hash
    expect(JSON.parse(out)).toEqual({ hash: expected, session: expected })
    const failure = runInContext(
      `try { FormulaEngine.compute({ sheets: [{ name: 'S', cells: { A1: { formula: '1+1' } } }] }, { execution: 'eval' }); 'computed' } catch (e) { e.name + ': ' + e.message }`,
      ctx,
    ) as string
    expect(failure).toBe(
      `EvalError: ${PACKAGE_NAME} compiles formulas with new Function for execution 'eval', and this environment blocks it (a Content Security Policy without 'unsafe-eval'?); execution 'closures' does not need it`,
    )
  }, 60_000)
})
