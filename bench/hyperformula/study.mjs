import { strict as assert } from 'node:assert'
import { mkdirSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createState, applyEdits } from '../../dist/engine.js'
import { run } from '../../dist/exec/budget.js'
import { HyperFormula, HF_OPTIONS } from './shared.mjs'

// Count formula executions, rather than ranking instrumented timings.
const rows = 5000, cells = { A1: { value: 1.1 }, A2: { formula: 'INT(A1)' } }
for (let r = 3; r <= rows; r++) cells[`A${r}`] = { formula: `A${r - 1}+1` }
const book = { sheets: [{ name: 'S', cells }] }, results = []
for (const execution of ['closures', 'eval']) {
  const state = createState(book, { now: 0, seed: 1, execution })
  let evaluated = 0
  for (const f of state.formulas) {
    const fn = f.fn
    f.fn = () => { evaluated++; return fn() }
  }
  for (const value of [1.9, 2.1]) {
    evaluated = 0
    run(applyEdits(state, [{ sheet: 'S', cell: 'A1', input: { value } }]))
    const lastValue = state.rt.cell(0, rows - 1, 0)
    assert.equal(lastValue, value === 1.9 ? rows - 1 : rows)
    results.push({ engine: `xl-${execution}`, value, evaluated, lastValue })
  }
}
const hf = HyperFormula.buildFromArray(Array.from({ length: rows }, (_, i) => [i === 0 ? 1.1 : i === 1 ? '=INT(A1)' : `=A${i}+1`]), HF_OPTIONS)
try {
  let evaluated = 0
  const evaluator = hf._evaluator, evaluate = evaluator.recomputeFormulaVertexValue.bind(evaluator)
  evaluator.recomputeFormulaVertexValue = (vertex) => { evaluated++; return evaluate(vertex) }
  for (const value of [1.9, 2.1]) {
    evaluated = 0
    hf.setCellContents({ sheet: 0, col: 0, row: 0 }, [[value]])
    const lastValue = hf.getCellValue({ sheet: 0, col: 0, row: rows - 1 })
    assert.equal(lastValue, value === 1.9 ? rows - 1 : rows)
    results.push({ engine: 'hf-default', value, evaluated, lastValue })
  }
} finally { hf.destroy() }

const graph = HyperFormula.buildFromArray([[1, 2, '=SUM(A1:B1)', '=C1*2']], HF_OPTIONS)
let dependencies
try {
  dependencies = {
    precedentsOfC1: graph.getCellPrecedents({ sheet: 0, row: 0, col: 2 }),
    dependentsOfA1: graph.getCellDependents({ sheet: 0, row: 0, col: 0 }),
  }
} finally { graph.destroy() }
const report = { hyperformula: HyperFormula.version, rows, results, dependencies,
  instrumentation: 'xlcalc Formula.fn and HF 3.5.0 private evaluator method; execution counts only, not public API recommendations or timed benchmark samples' }
const dir = resolve(process.argv[2] ?? 'bench/hyperformula/results')
mkdirSync(dir, { recursive: true })
writeFileSync(resolve(dir, 'hf-architecture-probes.json'), JSON.stringify(report, null, 2) + '\n')
console.log(JSON.stringify(report, null, 2))
