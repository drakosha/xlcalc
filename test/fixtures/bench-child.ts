/**
 * Замер пакета на синтетике в свежем процессе (для `budgets.test.ts`; `--expose-gc`). Методика как в
 * WB1: первый расчёт, затем 1 прогон на прогрев и 5 в зачёт; память это прирост heapUsed и rss против
 * пустого процесса с открытой сессией; 100 правок класса А; 5 правок класса Б; проверка всех формул.
 */
const m0 = process.memoryUsage()
const rows = Number(process.argv[2] ?? 3400)
const { compute, openSession, validateFormula } = await import('../../src/index')
const { synthBook, classAEdits, classBEdit } = await import('./synth')

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

const first = time(() => compute(synth.book, { clock }))
compute(synth.book, { clock })
const cold = [0, 1, 2, 3, 4].map(() => time(() => compute(synth.book, { clock })))
gc()
const session = openSession(synth.book, { clock })
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
const r1 = (x: number) => Math.round(x * 10) / 10
const mb = (x: number) => Math.round((x / 1048576) * 10) / 10
console.log(
  JSON.stringify({
    rows,
    formulas: formulas.length,
    firstMs: r1(first),
    coldMedianMs: r1(median(cold)),
    coldMaxMs: r1(Math.max(...cold)),
    heapUsedMB: mb(m1.heapUsed - m0.heapUsed),
    rssMB: mb(m1.rss - m0.rss),
    editA: { p50: r1(median(editA)), p95: r1(p95(editA)), max: r1(Math.max(...editA)) },
    editB: { median: r1(median(editB)), max: r1(Math.max(...editB)) },
    validateMedianMs: r1(median(validate)),
    hash: hash.slice(0, 16),
  }),
)

export {}
