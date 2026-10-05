// A consumer of the built package: plain node without tsx, importing by package name (self-reference
// through `exports`), the way a project that installed the tarball sees it. The name comes from
// package.json, so renaming the package does not touch this file.
import { readFileSync } from 'node:fs'

const { name } = JSON.parse(readFileSync(new URL('../../package.json', import.meta.url), 'utf8'))
const { compute, CONFIG_VERSION, engine, FUNCTION_NAMES, InputError, openSession, validateFormula } = await import(name)
const contract = await import(`${name}/contract`)
const { createState, parse, valuesHash } = await import(`${name}/unstable/internal`)
const { classAEdits, synthBook } = await import(`${name}/unstable/synth`)

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
    closed: await Promise.all([`${name}/dist/engine.js`, `${name}/internal`, `${name}/synth`].map(closed)),
  }),
)
