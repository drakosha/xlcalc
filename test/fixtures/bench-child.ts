/**
 * Benchmark of the package on the synthetic book in a fresh process (for `budgets.test.ts`; `--expose-gc`),
 * in one execution mode (`closures` or `eval`, the second argument), with the cells' display or without it (`off`, the
 * third argument): the first compute, then 1 warm-up run
 * and 5 measured runs; memory is the heapUsed and rss growth over an empty process with an open session;
 * 100 class A edits; 5 class B edits; validating all formulas; compiling the book's parsed formulas alone
 * (5 runs).
 */
import type { Execution } from '../../src/contract.js'

const m0 = process.memoryUsage()
const rows = Number(process.argv[2] ?? 3400)
const execution = (process.argv[3] ?? 'closures') as Execution
const display = process.argv[4] !== 'off'
const { libUrl } = await import('./target.js')
const { compute, openSession, validateFormula } = (await import(libUrl('index.js'))) as typeof import('../../src/index.js')
const { synthBook, classAEdits, classBEdit } = (await import(libUrl('synth.js'))) as typeof import('../../src/synth.js')
const { codegenContext, createState } = (await import(libUrl('engine.js'))) as typeof import('../../src/engine.js')
const { makeCompiler } = (await import(libUrl('compile.js'))) as typeof import('../../src/compile.js')
const { parse } = (await import(libUrl('syntax.js'))) as typeof import('../../src/syntax.js')

const gc = () => (globalThis as { gc?: () => void }).gc?.()
const median = (xs: number[]) => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)]
const p95 = (xs: number[]) => [...xs].sort((a, b) => a - b)[Math.ceil(0.95 * xs.length) - 1]
const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const synth = synthBook({ seed: 1, rows })
const time = (f: () => unknown) => {
  const t = performance.now()
  f()
  return performance.now() - t
}

const first = time(() => compute(synth.book, { clock, execution, display }))
compute(synth.book, { clock, execution, display })
const cold = [0, 1, 2, 3, 4].map(() => time(() => compute(synth.book, { clock, execution, display })))
gc()
const session = openSession(synth.book, { clock, execution, display })
gc()
const m1 = process.memoryUsage()
const editA = classAEdits(synth, 101, 100).map((e) => time(() => session.setCells([e])))
const editB = [0, 1, 2, 3, 4].map((i) => time(() => session.setCells([classBEdit(synth, 200 + i)])))
const hash = session.result().hash
session.close()
const formulas = synth.book.sheets.flatMap((s) => Object.values(s.cells).flatMap((c) => (c.formula === undefined ? [] : [c.formula])))
const sheets = synth.book.sheets.map((s) => s.name)
const names = (synth.book.names ?? []).map((n) => n.name)
const validate = [0, 1, 2, 3, 4].map(() => time(() => formulas.forEach((f) => validateFormula(f, { sheets, names }))))
const state = createState(synth.book, { now: 0, seed: 1, execution })
const parsed = state.formulas.map((f) => ({ f: f!, ast: parse(f!.text) }))
const compile = [0, 1, 2, 3, 4].map(() =>
  time(() => {
    const c = makeCompiler(state.rt, execution)
    c.link(parsed.map(({ f, ast }) => c.formula(codegenContext(state, f.s), ast, f.s, f.r, f.c)))
  }),
)
const r1 = (x: number) => Math.round(x * 10) / 10
const mb = (x: number) => Math.round((x / 1048576) * 10) / 10
console.log(
  JSON.stringify({
    rows,
    execution,
    display,
    formulas: formulas.length,
    firstMs: r1(first),
    coldMedianMs: r1(median(cold)),
    coldMaxMs: r1(Math.max(...cold)),
    heapUsedMB: mb(m1.heapUsed - m0.heapUsed),
    rssMB: mb(m1.rss - m0.rss),
    editA: { p50: r1(median(editA)), p95: r1(p95(editA)), max: r1(Math.max(...editA)) },
    editB: { median: r1(median(editB)), max: r1(Math.max(...editB)) },
    validateMedianMs: r1(median(validate)),
    compileMedianMs: r1(median(compile)),
    hash: hash.slice(0, 16),
  }),
)

export {}
