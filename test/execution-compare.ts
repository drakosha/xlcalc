/**
 * The two execution modes side by side, for the tests that compare them: `eval` here, and `closures` in the
 * browser bundle inside a `node:vm` context where code generation from strings is disabled, so the closures
 * mode cannot be running generated code (the eval mode fails there). Results are cloned out of the context.
 */
import { createContext, runInContext } from 'node:vm'
import { expect } from 'vitest'
import { compute, openSession, type BookInput, type CellValue, type ComputeOptions, type ComputeResult, type Execution } from '../src/index.js'
import { synthBook } from '../src/synth.js'
import { bundle } from './fixtures/bundle.js'
import { apply, edits } from './session-edits.js'

export const clock = { now: '2026-10-05T12:00:00Z', seed: 1 }

export const ctx = createContext({}, { codeGeneration: { strings: false } })
runInContext((await bundle()).code, ctx)
/** The library inside the context without code generation from strings; its results are cloned out of it. */
export const strict = runInContext('FormulaEngine', ctx) as typeof import('../src/index.js')
// the context's own registry of locales, loaded as the tests' (`locales.ts`)
runInContext('FormulaEngine.registerLocale(FormulaEngine.ruRU)', ctx)

export const byEval = (book: BookInput, opts: ComputeOptions = {}) => compute(book, { clock, ...opts, execution: 'eval' })
export const byClosures = (book: BookInput, opts: ComputeOptions = {}) => structuredClone(strict.compute(book, { clock, ...opts, execution: 'closures' }))

const show = (v: CellValue | undefined) => (v !== null && typeof v === 'object' ? `#${v.error}` : Object.is(v, -0) ? '-0' : JSON.stringify(v))
const sameValue = (a: CellValue | undefined, b: CellValue | undefined) => Object.is(a, b) || (typeof a === 'object' && typeof b === 'object' && a?.error === b?.error)

/** The results of the two modes: cell by cell (values, formula text), then everything else but the timing. */
export function expectSame(evalResult: ComputeResult, closuresResult: ComputeResult) {
  const off: string[] = []
  evalResult.sheets.forEach((s, i) => {
    const other = closuresResult.sheets[i]
    if (other?.name !== s.name || Object.keys(other.cells).length !== Object.keys(s.cells).length) off.push(`sheet ${s.name}: other cells`)
    for (const [addr, c] of Object.entries(s.cells)) {
      const o = other?.cells[addr]
      if (!sameValue(c.value, o?.value) || c.formula !== o?.formula) off.push(`${s.name}!${addr} ${c.formula ?? ''}: eval ${show(c.value)}, closures ${show(o?.value)}`)
    }
  })
  expect(off).toEqual([])
  const rest = ({ sheets: _s, ms: _m, ...r }: ComputeResult) => r
  expect(rest(closuresResult)).toEqual(rest(evalResult))
}

/** 200 seeded session edits in the mode `session`; after each one the session equals a full computation in the other mode. */
export function sessionAgainstFull(session: Execution) {
  const synth = synthBook({ seed: 1, rows: 340 })
  const s = session === 'closures' ? strict.openSession(synth.book, { clock, execution: 'closures' }) : openSession(synth.book, { clock, execution: 'eval' })
  let book = synth.book
  for (const batch of edits(3, 200, synth)) {
    s.setCells(batch)
    book = apply(book, batch)
    const r = structuredClone(s.result())
    if (session === 'closures') expectSame(byEval(book), r)
    else expectSame(r, byClosures(book))
  }
}
