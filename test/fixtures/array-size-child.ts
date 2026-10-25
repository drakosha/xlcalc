/**
 * Formulas whose arrays hold more values than one JavaScript array, in a fresh process (for `review-1021-arrays.test.ts`):
 * before 25.10 V8 stopped the process (`invalid table size`), which no test could catch. Computes them in a book and in a
 * session, with the execution of the first argument, and prints the values by cell as JSON.
 */
const execution = (process.argv[2] ?? 'closures') as 'closures' | 'eval'
const { libUrl } = await import('./target.js')
const { compute, createSession } = (await import(libUrl('index.js'))) as typeof import('../../src/index.js')

const FORMULAS = ['ROWS(SEQUENCE(2^31))', 'ROWS(RANDARRAY(2^31))', 'ROWS(SEQUENCE(1048576,1024))', 'ROWS(MUNIT(46341))', 'ROWS(EXPAND(1,2^31))', 'ROWS(SEQUENCE(2^15)*SEQUENCE(1,2^14))', 'SUM(SEQUENCE(2^16)*TRANSPOSE(SEQUENCE(2^16)))', 'COLUMNS(SEQUENCE(1,2^28))', 'ROWS(SEQUENCE(1048576))']
const cells = Object.fromEntries(FORMULAS.map((formula, i) => [`A${i + 1}`, { formula }]))
const book = { sheets: [{ name: 'S', cells }] }
const r = compute(book, { execution })
const s = createSession({ sheets: [{ name: 'S', cells: {} }] }, { execution })
s.setCells(FORMULAS.map((formula, i) => ({ sheet: 'S', cell: `B${i + 1}`, input: { formula } })))
const typed = s.getCells('S', FORMULAS.map((_, i) => `B${i + 1}`))
console.log(JSON.stringify({ computed: FORMULAS.map((_, i) => r.sheets[0].cells[`A${i + 1}`].value), typed: FORMULAS.map((_, i) => typed[`B${i + 1}`]?.value) }))
