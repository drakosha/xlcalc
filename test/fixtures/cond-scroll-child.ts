/**
 * The heap a scroll through a conditional format keeps, in a fresh process (for `cond-scroll-memory.test.ts`;
 * `--expose-gc`): the block of `rows` rows with the rule `formula` (`cond-block.ts`), its first window read, then every
 * window from the top to the bottom. Prints the heap above the one after the first window, in MB, the session still held.
 */
const [rows, formula] = [Number(process.argv[2]), process.argv[3]]
const { libUrl } = await import('./target.js')
const { createSession } = (await import(libUrl('index.js'))) as typeof import('../../src/index.js')
const { condBlock, windows } = await import('./cond-block.js')
const { heap } = await import('./heap.js')

const s = createSession(condBlock(rows, formula), { clock: { now: '2026-10-04T12:00:00Z', seed: 7 }, display: false })
s.getViewport('S', 'A1:J50')
const start = await heap()
for (const w of windows(rows)) s.getViewport('S', w)
const above = (await heap()) - start
s.getHash()
console.log(above.toFixed(2))
