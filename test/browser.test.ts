/**
 * Outside Node: the library bundled by esbuild for the browser (with formulajs) has no Node built-ins,
 * and the bundle runs in an empty `node:vm` context (only the ECMAScript globals: no `process`, `Buffer`,
 * `require`, `performance` or `crypto`) and gives the same hash as Node. With code generation from strings
 * disabled, as under a Content Security Policy without 'unsafe-eval', the bundle loads and validates
 * formulas, and computing fails with a clear EvalError.
 */
import { build } from 'esbuild'
import { fileURLToPath } from 'node:url'
import { createContext, runInContext } from 'node:vm'
import { describe, expect, it } from 'vitest'
import { compute } from '../src/index.js'
import { synthBook } from '../src/synth.js'
import { libUrl } from './fixtures/target.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

async function bundle(): Promise<{ code: string; inputs: string[] }> {
  const r = await build({
    stdin: { contents: `export * from ${JSON.stringify(fileURLToPath(libUrl('index.js')))}\nexport { synthBook } from ${JSON.stringify(fileURLToPath(libUrl('synth.js')))}`, loader: 'ts', resolveDir: fileURLToPath(new URL('..', import.meta.url)) },
    bundle: true,
    platform: 'browser',
    format: 'iife',
    globalName: 'FormulaEngine',
    target: 'es2020',
    write: false,
    metafile: true,
    logLevel: 'silent',
  })
  return { code: r.outputFiles[0].text, inputs: Object.keys(r.metafile.inputs) }
}

describe('outside Node', () => {
  it('the browser bundle has no Node built-ins and runs in an empty vm context with the same hash', async () => {
    const { code, inputs } = await bundle()
    expect(inputs.filter((p) => p.startsWith('node:') || !/^(src|dist|node_modules)\/|<stdin>/.test(p))).toEqual([])
    expect(code).not.toMatch(/\brequire\s*\(|\bprocess\.|\bBuffer\b|["']node:/)

    const ctx = createContext({})
    expect(runInContext('[typeof process, typeof require, typeof Buffer, typeof performance, typeof crypto, typeof TextEncoder]', ctx)).toEqual(Array(6).fill('undefined'))
    runInContext(code, ctx)
    const out = runInContext(
      `const s = FormulaEngine.synthBook({ seed: 1, rows: 340 })
       const session = FormulaEngine.openSession(s.book, { clock: ${JSON.stringify(clock)} })
       JSON.stringify({ hash: FormulaEngine.compute(s.book, { clock: ${JSON.stringify(clock)} }).hash, session: session.result().hash, seed: typeof FormulaEngine.openSession(s.book, { live: true }).clock.seed })`,
      ctx,
    ) as string
    const expected = compute(synthBook({ seed: 1, rows: 340 }).book, { clock }).hash
    expect(JSON.parse(out)).toEqual({ hash: expected, session: expected, seed: 'number' })
  }, 60_000)

  it('without code generation from strings (strict CSP) the bundle loads, validates, and compute explains the failure', async () => {
    const { code } = await bundle()
    const ctx = createContext({}, { codeGeneration: { strings: false } })
    expect(() => runInContext('new Function("return 1")', ctx)).toThrow()
    runInContext(code, ctx)
    expect(runInContext('FormulaEngine.validateFormula("SUM(A1:A3)*2").ok', ctx)).toBe(true)
    const out = runInContext(
      `try { FormulaEngine.compute({ sheets: [{ name: 'S', cells: { A1: { formula: '1+1' } } }] }); 'computed' } catch (e) { e.name + ': ' + e.message }`,
      ctx,
    ) as string
    expect(out).toBe("EvalError: formula-engine compiles formulas with new Function, and this environment blocks it (a Content Security Policy without 'unsafe-eval'?)")
  }, 60_000)
})
