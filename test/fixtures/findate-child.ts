/** Child process for test/functions-findate-tz.test.ts: date and financial formulas computed in the process time zone. */
import { libUrl } from './target.js'

const { compute } = (await import(libUrl('index.js'))) as typeof import('../../src/index.js')

const book = JSON.parse(process.argv[2]) as Parameters<typeof compute>[0]
const cells = compute(book, { clock: { now: '2026-10-04T23:30:00Z', seed: 7 } }).sheets[0].cells
const values = Object.keys(book.sheets[0].cells).map((a) => cells[a].value)
console.log(JSON.stringify({ values, offset: new Date(2025, 0, 1).getTimezoneOffset() }))
