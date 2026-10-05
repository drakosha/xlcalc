// Потребитель собранного пакета: чистый node без tsx, импорт по имени пакета (self-reference через
// `exports`), как его видит проект, поставивший тарбол.
import { compute, CONFIG_VERSION, engine, FUNCTION_NAMES, openSession, validateFormula } from '@mike/formula'
import * as contract from '@mike/formula/contract'
import { createState, parse, valuesHash } from '@mike/formula/internal'
import { classAEdits, synthBook } from '@mike/formula/synth'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const synth = synthBook({ seed: 1, rows: 340 })
const hash = compute(synth.book, { clock }).hash
const session = openSession(synth.book, { clock })
for (const e of classAEdits(synth, 3, 10)) session.setCells([e])
const edited = session.result().hash
session.close()
let deep = 'resolved'
try {
  await import('@mike/formula/dist/engine.js')
} catch (e) {
  deep = e.code
}
console.log(
  JSON.stringify({
    hash,
    edited,
    configVersion: CONFIG_VERSION,
    functions: FUNCTION_NAMES.length,
    engine: Object.keys(engine).sort(),
    validate: validateFormula('SUM(A1:A3)').ok,
    contract: Object.keys(contract).sort(),
    internal: [typeof createState, typeof parse, typeof valuesHash],
    deep,
  }),
)
