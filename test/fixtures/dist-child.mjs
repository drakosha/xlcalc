// A consumer of the built package: plain node without tsx, importing by package name (self-reference
// through `exports`), the way a project that installed the tarball sees it.
import { compute, CONFIG_VERSION, engine, FUNCTION_NAMES, InputError, openSession, validateFormula } from 'formula-engine'
import * as contract from 'formula-engine/contract'
import { createState, parse, valuesHash } from 'formula-engine/unstable/internal'
import { classAEdits, synthBook } from 'formula-engine/unstable/synth'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const synth = synthBook({ seed: 1, rows: 340 })
const hash = compute(synth.book, { clock }).hash
const session = openSession(synth.book, { clock })
for (const e of classAEdits(synth, 3, 10)) session.setCells([e])
const edited = session.result().hash
session.close()
/** Error code of importing a path that `exports` does not expose. */
async function closed(spec) {
  try {
    await import(spec)
    return 'resolved'
  } catch (e) {
    return e.code
  }
}
console.log(
  JSON.stringify({
    hash,
    edited,
    configVersion: CONFIG_VERSION,
    functions: FUNCTION_NAMES.length,
    engine: Object.keys(engine).sort(),
    inputError: typeof InputError,
    validate: validateFormula('SUM(A1:A3)').ok,
    contract: Object.keys(contract).sort(),
    internal: [typeof createState, typeof parse, typeof valuesHash],
    closed: await Promise.all(['formula-engine/dist/engine.js', 'formula-engine/internal', 'formula-engine/synth'].map(closed)),
  }),
)
