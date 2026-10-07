/**
 * The thread pool under test. With the built `dist` the threads run `dist/worker/thread.js`; with the sources, Node
 * cannot run `src/worker/thread.ts` in a thread (its imports name `.js` files), so the thread is bundled by esbuild
 * into one module first. `peek` sees the pool's threads and sessions; `inside` also crashes a thread.
 */
import { build } from 'esbuild'
import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import type { Calc, CalcPoolOptions } from '../../src/contract.js'
import { checkPoolOptions } from '../../src/async/options.js'
import { nodeSpawn, workerThreads, type Thread, type ThreadLoss } from '../../src/worker/platform.js'
import { createPool, type PoolInside } from '../../src/worker/pool.js'
import { libUrl, TARGET } from './target.js'

let script: URL | undefined

async function threadScript(): Promise<URL> {
  if (TARGET === 'dist') return new URL(libUrl('worker/thread.js'))
  if (!script) {
    const r = await build({ entryPoints: [fileURLToPath(libUrl('worker/thread.js')).replace(/\.js$/, '.ts')], bundle: true, platform: 'node', format: 'esm', write: false, logLevel: 'silent' })
    const file = join(mkdtempSync(join(tmpdir(), 'xlcalc-thread-')), 'thread.mjs')
    writeFileSync(file, r.outputFiles[0].text)
    script = pathToFileURL(file)
  }
  return script
}

/** A pool as `createCalcPool` makes it in Node, over the thread under test. */
export async function pool(opts?: CalcPoolOptions, peek?: (inside: PoolInside) => void): Promise<Calc> {
  const o = checkPoolOptions(opts)
  return createPool(o, nodeSpawn(workerThreads()!, await threadScript(), o.maxHeapMb), peek)
}

/** A pool with its inside, and a crash of thread `i` as Node reports a thread that exited by itself. */
export async function inside(opts?: CalcPoolOptions): Promise<{ calc: Calc; inside: PoolInside; crash(i: number): void }> {
  const o = checkPoolOptions(opts)
  const base = nodeSpawn(workerThreads()!, await threadScript(), o.maxHeapMb)
  const losses = new Map<Thread, (l: ThreadLoss) => void>()
  let seen: PoolInside | undefined
  const calc = createPool(
    o,
    (onMessage, onLoss) => {
      const t = base(onMessage, onLoss)
      losses.set(t, onLoss)
      return t
    },
    (x) => (seen = x),
  )
  return {
    calc,
    inside: seen!,
    crash(i) {
      const t = seen!.slots[i].thread!
      t.terminate()
      losses.get(t)!({ code: 'thread_failed', message: 'the thread exited with code 1' })
    },
  }
}
