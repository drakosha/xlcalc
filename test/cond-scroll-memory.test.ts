/**
 * A scroll through a conditional format keeps no code per cell. A rule whose formula moves from cell to cell by its
 * references (`A1>50`) or by the functions of its cell (`MOD(ROW(),2)=0`) makes its code once however many cells a scroll
 * computes it at; the heap a scroll of 200,000 cells leaves stays under 30 MB and within 5 MB of the heap 100,000 leave
 * (code made per cell once kept about 1 KB a cell: 205 MB); a pool's thread of 256 MB scrolls the 200,000 cells (it ran
 * out of memory after 387 windows). The heap in fresh processes with `--expose-gc` (`fixtures/cond-scroll-child.ts`).
 */
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { expect, it, vi } from 'vitest'
import { createSession } from '../src/index.js'
import { condBlock, windows } from './fixtures/cond-block.js'
import { closedAfter } from './fixtures/pool.js'

// the code a rule's formula makes: each piece of it its own backend of closures
const made = vi.hoisted(() => ({ backends: 0 }))
vi.mock('../src/exec/closures.js', async (importOriginal) => {
  const m = await importOriginal<typeof import('../src/exec/closures.js')>()
  return { ...m, closureBackend: (...a: Parameters<typeof m.closureBackend>) => (made.backends++, m.closureBackend(...a)) }
})

const child = fileURLToPath(new URL('./fixtures/cond-scroll-child.ts', import.meta.url))
const opts = { clock: { now: '2026-10-04T12:00:00Z', seed: 7 }, display: false }
const { made: pool } = closedAfter()

it.each(['A1>50', 'MOD(ROW(),2)=0'])('the rule %s makes its code once however many cells a scroll computes it at', (formula) => {
  const backends = (rows: number) => {
    const s = createSession(condBlock(rows, formula), opts)
    const before = made.backends
    for (const w of windows(rows)) s.getViewport('S', w)
    s.close()
    return made.backends - before
  }
  expect(backends(2000)).toBe(backends(1000))
})

it.each(['A1>50', 'MOD(ROW(),2)=0'])('a scroll with the rule %s keeps under 30 MB, not growing with the cells', (formula) => {
  const above = (rows: number) => Number(execFileSync(process.execPath, ['--expose-gc', '--import', 'tsx', child, String(rows), formula], { encoding: 'utf8' }))
  const [half, all] = [above(10_000), above(20_000)]
  expect(all).toBeLessThan(30)
  expect(all - half).toBeLessThanOrEqual(5)
}, 120_000)

it("a pool's thread of 256 MB scrolls 200,000 cells with a rule twice", async () => {
  const calc = await pool({ threads: 1, maxHeapMb: 256 })
  const s = await calc.createSession(condBlock(20_000, 'A1>50'), opts)
  for (let pass = 0; pass < 2; pass++) for (const w of windows(20_000)) await s.getViewport('S', w)
}, 120_000)
