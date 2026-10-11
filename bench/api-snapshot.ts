/** Compare snapshot/export cost with a separate checkout: node --expose-gc --import tsx bench/api-snapshot.ts <root> <sync|inline|worker> <closures|eval>. */
import { performance } from 'node:perf_hooks'
import { resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import type { AsyncSession, BookInput, Execution, Session } from '../src/contract.js'

const [root, kind, mode] = process.argv.slice(2)
if (!root || !['sync', 'inline', 'worker'].includes(kind) || !['closures', 'eval'].includes(mode) || !globalThis.gc)
  throw new Error('Use node --expose-gc --import tsx bench/api-snapshot.ts <root> <sync|inline|worker> <closures|eval>')
const load = (file: string) => import(pathToFileURL(resolve(root, file)).href)
const api = await load('src/index.ts') as typeof import('../src/index.js')
const { synthBook } = await load('src/synth.ts') as typeof import('../src/synth.js')
const { writeXlsx } = await load('src/xlsx.ts') as typeof import('../src/xlsx.js')
const calc = kind === 'inline' ? api.createInlineCalc() : kind === 'worker' ? await (await load('test/fixtures/pool.ts')).pool({ threads: 1 }) as import('../src/contract.js').Calc : null
const execution = mode as Execution
const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const open = (book: BookInput): Session | Promise<AsyncSession> => calc ? calc.createSession(book, { clock, execution }) : api.createSession(book, { clock, execution })
const s = await open(synthBook({ seed: 1, rows: 34_000 }).book)
const gc = globalThis.gc

async function sample() {
  gc()
  const before = process.memoryUsage().heapUsed
  const start = performance.now()
  const saved = await s.getSnapshot()
  const snapshotMs = performance.now() - start
  gc()
  const snapshotHeapMb = (process.memoryUsage().heapUsed - before) / 2 ** 20
  const writing = performance.now()
  const written = await writeXlsx(saved)
  return { snapshotMs, writeMs: performance.now() - writing, snapshotHeapMb, bytes: written.data.length, hash: saved.result.hash, inputHash: saved.result.inputHash }
}

try {
  await sample() // One warmup, followed by exactly five measured samples, on each checkout.
  const samples: Awaited<ReturnType<typeof sample>>[] = []
  for (let i = 0; i < 5; i++) samples.push(await sample())
  const median = (field: 'snapshotMs' | 'writeMs' | 'snapshotHeapMb') => samples.map((s) => s[field]).sort((a, b) => a - b)[2]
  if (samples.some((r) => !r.hash.startsWith('e64d5f4d0f0ac632'))) throw new Error('The synthetic values hash changed')
  console.log(JSON.stringify({ root, kind, execution, node: process.version, medians: { snapshotMs: median('snapshotMs'), writeMs: median('writeMs'), snapshotHeapMb: median('snapshotHeapMb') }, samples }))
} finally {
  await s.close()
  await calc?.close()
}
