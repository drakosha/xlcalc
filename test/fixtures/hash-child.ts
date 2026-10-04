/** Дочерний процесс тестов детерминизма: хэш синтетики (строк по аргументу) и часовая зона процесса. */
import { compute } from '../../src/index'
import { synthBook } from './synth'

const rows = Number(process.argv[2] ?? 340)
const r = compute(synthBook({ seed: 1, rows }).book, { clock: { now: '2026-10-04T12:00:00Z', seed: 7 } })
console.log(JSON.stringify({ hash: r.hash, tz: Intl.DateTimeFormat().resolvedOptions().timeZone, offset: new Date(2025, 0, 1).getTimezoneOffset() }))
