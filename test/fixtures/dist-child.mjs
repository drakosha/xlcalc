// A consumer of the built package: plain node without tsx, importing by package name (self-reference
// through `exports`), the way a project that installed the tarball sees it. The name comes from
// package.json, so renaming the package does not touch this file.
import { readFileSync } from 'node:fs'

const { name } = JSON.parse(readFileSync(new URL('../../package.json', import.meta.url), 'utf8'))
const root = await import(name)
const { compute, CONFIG_VERSION, createSession, FUNCTION_NAMES, InputError, validateFormula } = root
const contract = await import(`${name}/contract`)
const text = await import(`${name}/text`)
const { createState, parse, valuesHash } = await import(`${name}/unstable/internal`)
const { classAEdits, synthBook } = await import(`${name}/unstable/synth`)
const { loadLocale, readXlsx, writeXlsx } = await import(`${name}/xlsx`)
const { createCalcPool } = await import(`${name}/worker`)

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const synth = synthBook({ seed: 1, rows: 340 })
const hash = compute(synth.book, { clock }).hash
const session = createSession(synth.book, { clock })
for (const e of classAEdits(synth, 3, 10)) session.setCells([e])
const edited = session.getResult().hash
session.close()
const xlsx = await readXlsx((await writeXlsx({ book: synth.book, result: compute(synth.book, { clock }) })).data)
// the pool's threads from their default place next to the subpath; left open, it must not keep the process alive
const pool = createCalcPool({ threads: 1 })
const inThread = await pool.createSession(synth.book, { clock })
for (const e of classAEdits(synth, 3, 10)) await inThread.setCells([e])
const threads = { hash: (await pool.compute(synth.book, { clock })).hash, edited: (await inThread.getResult()).hash }
// a package locale loaded by its tag from the xlsx subpath: its module found in dist, a file written in it, the pool's thread given it
const ru = await loadLocale('ru-RU')
const ruBook = { locale: 'ru-RU', sheets: [{ name: 'S', cells: { A1: { value: 2 }, B1: { formula: 'A1*2' } } }] }
const loaded = { tag: ru.tag, same: (await root.loadLocale('ru-RU')) === ru, read: (await readXlsx((await writeXlsx({ book: ruBook })).data)).book.locale, thread: (await pool.compute(ruBook, { clock })).sheets[0].cells.B1.value }
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
    // the old names, removed: none of them is exported
    removed: ['engine', 'openSession', 'restructure', 'fromLegacy'].filter((k) => k in root),
    inputError: typeof InputError,
    validate: validateFormula('SUM(A1:A3)').ok,
    contract: Object.keys(contract).sort(),
    internal: [typeof createState, typeof parse, typeof valuesHash],
    text: { keys: Object.keys(text).sort(), same: Object.entries(text).every(([k, f]) => root[k] === f) },
    xlsx: compute(xlsx.book, { clock }).hash,
    threads,
    loaded,
    closed: await Promise.all([`${name}/dist/engine.js`, `${name}/internal`, `${name}/synth`].map(closed)),
  }),
)
