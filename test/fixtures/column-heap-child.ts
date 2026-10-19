/**
 * The heap a session over a sparse sheet keeps, in a fresh process (for `column.test.ts`; `--expose-gc`): 1,000 values
 * 97 rows apart in eight columns, each read by a formula in column KY, so that nearly every chunk of a column holds one
 * cell. Prints the heap the session holds, in MB.
 */
import { setImmediate as tick } from 'node:timers/promises'
const { libUrl } = await import('./target.js')
const { createSession, cellAddress } = (await import(libUrl('index.js'))) as typeof import('../../src/index.js')

const cells: Record<string, { value?: number; formula?: string }> = {}
for (let i = 0; i < 1000; i++) {
  const at = cellAddress(i * 97 + 1, (i % 8) * 31 + 1)
  cells[at] = { value: i + 1 }
  cells[cellAddress(i * 97 + 1, 311)] = { formula: `${at}*3` }
}
const settle = async () => {
  for (let i = 0; i < 2; i++) (await tick(), global.gc!())
  return process.memoryUsage().heapUsed
}
const before = await settle()
const session = createSession({ sheets: [{ name: 'S', cells }] }, { display: false })
const after = await settle()
console.log(JSON.stringify({ heap: (after - before) / 2 ** 20, sum: session.getCell('S', 'KY1')?.value }))
