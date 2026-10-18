/**
 * Whether a sheet's values are kept indexed, in a fresh process (for `sheet-columns.test.ts`; `--allow-natives-syntax`):
 * a column of 34,000 values with one cell far below them (row 100,000), and a column of a cell every 50,000 rows to the
 * sheet's last. Prints, as JSON, the arrays of the sheet's storage the JavaScript engine keeps as dictionaries rather than
 * by index, and the values read back at a few rows.
 */
const { libUrl } = await import('./target.js')
const { createState } = (await import(libUrl('engine.js'))) as typeof import('../../src/engine.js')

const cells: Record<string, { value: number | string }> = {}
for (let r = 1; r <= 34_000; r++) cells[`A${r}`] = { value: r * 1.5 }
cells.A100000 = { value: 'note' }
for (let r = 1; r <= 1_048_576; r += 50_000) cells[`B${r}`] = { value: r }
const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 0, seed: 1 })

const dictionary = new Function('x', 'return %HasDictionaryElements(x)') as (x: object) => boolean
let dictionaries = 0
const seen = new Set<object>()
const walk = (x: unknown) => {
  if (x === null || typeof x !== 'object' || seen.has(x)) return
  seen.add(x)
  if (Array.isArray(x)) {
    if (dictionary(x)) dictionaries++
    x.forEach(walk)
  } else Object.values(x).forEach(walk)
}
walk(state.ctx.sheets[0])
const read = [[0, 0], [33_999, 0], [99_999, 0], [50_000, 1], [1_000_000, 1], [1_000_001, 1]].map(([r, c]) => state.rt.cell(0, r, c))
console.log(JSON.stringify({ dictionaries, read }))
