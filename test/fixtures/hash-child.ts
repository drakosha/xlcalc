/** Child process for the determinism tests: the synthetic hash (rows from the argument) and the process time zone. */
import { libUrl } from './target.js'

const { compute } = (await import(libUrl('index.js'))) as typeof import('../../src/index.js')
const { synthBook } = (await import(libUrl('synth.js'))) as typeof import('../../src/synth.js')

const rows = Number(process.argv[2] ?? 340)
const r = compute(synthBook({ seed: 1, rows }).book, { clock: { now: '2026-10-04T12:00:00Z', seed: 7 } })
console.log(JSON.stringify({ hash: r.hash, tz: Intl.DateTimeFormat().resolvedOptions().timeZone, offset: new Date(2025, 0, 1).getTimezoneOffset() }))
