/**
 * The heap a session keeps after operations that build its book again, in a fresh process (for
 * `session-rebuild-memory.test.ts`; `--expose-gc`): a table of `rows` rows and 16 columns with a calculated column, a
 * formula over it on another sheet; its header renamed four times, then four rows inserted. Prints the heap after the
 * session opened and after each operation, in MB, as JSON.
 */
const rows = Number(process.argv[2] ?? 3400)
const { libUrl } = await import('./target.js')
const { createSession } = (await import(libUrl('index.js'))) as typeof import('../../src/index.js')
type Cells = Record<string, { value?: string | number; formula?: string }>

const col = (c: number) => String.fromCharCode(65 + c)
const names = ['Amount', ...Array.from({ length: 14 }, (_, i) => `C${i + 2}`), 'Double']
const cells: Cells = {}
names.forEach((n, c) => (cells[`${col(c)}1`] = { value: n }))
for (let r = 2; r <= rows + 1; r++) {
  for (let c = 0; c < 15; c++) cells[`${col(c)}${r}`] = { value: (r * 7 + c) % 1000 }
  cells[`P${r}`] = { formula: 'T[[#This Row],[Amount]]*2' }
}
const columns = names.map((name, c) => (c === 15 ? { name, formula: 'T[[#This Row],[Amount]]*2' } : { name }))
const s = createSession({ sheets: [{ name: 'Data', cells, tables: [{ name: 'T', range: `A1:P${rows + 1}`, columns }] }, { name: 'Sum', cells: { A1: { formula: 'SUM(T[Double])' } } }] }, { undo: { depth: 0 } })
const gc = (globalThis as { gc: () => void }).gc
const mb = () => {
  gc()
  gc()
  return process.memoryUsage().heapUsed / 2 ** 20
}
const heap = [mb()]
for (let k = 1; k <= 4; k++) {
  s.setCellText('Data', 'A1', `Amount${k}`)
  heap.push(mb())
}
for (let k = 1; k <= 4; k++) {
  s.editStructure({ op: 'insertRows', sheet: 'Data', row: 10 })
  heap.push(mb())
}
console.log(JSON.stringify({ heap, sum: s.getCell('Sum', 'A1')?.formula }))
