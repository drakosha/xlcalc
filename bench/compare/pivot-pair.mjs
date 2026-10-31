// node --expose-gc --max-old-space-size=8192 pivot-pair.mjs <build> <build> [build ...]: a PivotTable's refresh in two or
// more builds of xlcalc in one process, taking turns, the order turned round each round: a cache of 10,000 and of 100,000
// records, Region and Product on the rows, Quarter on the columns, Sum of Qty and Average of Price, both row fields sorted
// by Sum of Qty descending (equal sums among the products). Median and p95 of 15 refreshes after one to warm up, each
// build against the first, ratios to three decimals.
import { at } from '../stats.mjs'

const dirs = process.argv.slice(2)
const ROUNDS = Number(process.env.ROUNDS ?? 15)
const libs = await Promise.all(dirs.map(async (d) => ({ ...(await import(d + '/dist/index.js')), pivot: await import(d + '/dist/pivot/index.js') })))

/** The source: 40 regions, 600 products, 4 quarters; Qty 1 to 20 so that sums tie, Price with cents. */
function book(records) {
  const cells = { A1: { value: 'Region' }, B1: { value: 'Product' }, C1: { value: 'Quarter' }, D1: { value: 'Qty' }, E1: { value: 'Price' } }
  let seed = 7
  const rnd = (n) => (seed = (Math.imul(seed, 1103515245) + 12345) >>> 0) % n
  for (let r = 2; r <= records + 1; r++) {
    cells[`A${r}`] = { value: `Region ${rnd(40)}` }
    cells[`B${r}`] = { value: `Product ${rnd(600)}` }
    cells[`C${r}`] = { value: `Q${1 + rnd(4)}` }
    cells[`D${r}`] = { value: 1 + rnd(20) }
    cells[`E${r}`] = { value: rnd(100000) / 100 }
  }
  return { sheets: [{ name: 'Data', cells }, { name: 'Report', cells: {} }] }
}

const byQty = { order: 'descending', by: { value: 0 } }
const definition = { rows: [0, 1], columns: [2], values: [{ field: 3 }, { field: 4, function: 'average' }], fields: [{ field: 0, sort: byQty }, { field: 1, sort: byQty }] }
const f = (v) => (v < 10 ? v.toFixed(2) : v.toFixed(1))

for (const records of [10_000, 100_000]) {
  const input = book(records)
  const sessions = libs.map((l) => {
    l.registerPivot(l.pivot.pivotEngine)
    const s = l.createSession(input, { display: false })
    const r = s.editStructure({ op: 'addPivotTable', sheet: 'Report', cell: 'A3', source: { type: 'range', sheet: 'Data', range: `A1:E${records + 1}` }, definition })
    if (!r.ok) throw new Error(`addPivotTable: ${JSON.stringify(r)}`)
    return s
  })
  const ms = libs.map(() => [])
  for (let round = 0; round <= ROUNDS; round++)
    for (let k = 0; k < libs.length; k++) {
      const i = (k + round) % libs.length
      global.gc()
      const t = performance.now()
      const r = sessions[i].editStructure({ op: 'refreshPivots' })
      const dt = performance.now() - t
      if (!r.ok) throw new Error(`refreshPivots: ${JSON.stringify(r)}`)
      if (round) ms[i].push(dt)
    }
  sessions.forEach((s) => s.close())
  for (const q of [0.5, 0.95])
    console.log(`${records} records, refresh ${q === 0.5 ? 'p50' : 'p95'}: ${dirs.map((d, i) => `${d.split('/').pop()} ${f(at(ms[i], q))}${i ? ` (${(at(ms[0], q) / at(ms[i], q)).toFixed(3)}x)` : ''}`).join(' / ')} ms`)
}
