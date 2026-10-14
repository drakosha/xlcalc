/**
 * The thread pool under test. With the built `dist` the threads run `dist/worker/thread.js`; with the sources, Node
 * cannot run `src/worker/thread.ts` in a thread (its imports name `.js` files), so the thread is bundled by esbuild
 * into one module first. `peek` sees the pool's threads and sessions; `inside` also crashes a thread.
 */
import { build } from 'esbuild'
import { afterAll } from 'vitest'
import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import type { Calc, CalcPoolOptions } from '../../src/contract.js'
import type { Thread, ThreadLoss } from '../../src/worker/platform.js'
import type { PoolInside } from '../../src/worker/pool.js'
import { libUrl, TARGET } from './target.js'

// the modules of the build under test, so that their CalcError is the one the tests import
const { checkPoolOptions } = (await import(libUrl('async/options.js'))) as typeof import('../../src/async/options.js')
const { nodeSpawn, workerThreads } = (await import(libUrl('worker/platform.js'))) as typeof import('../../src/worker/platform.js')
const { createPool } = (await import(libUrl('worker/pool.js'))) as typeof import('../../src/worker/pool.js')

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

/**
 * Calcs closed after a test file's tests (`open`, in the order kept): `keep` one made otherwise, `made` a pool (`pool`). Called at the top of a file,
 * where `afterAll` registers the closing.
 */
export function closedAfter() {
  const open: Calc[] = []
  afterAll(() => Promise.all(open.map((c) => c.close())))
  const keep = <T extends Calc>(c: T): T => (open.push(c), c)
  return { open, keep, made: async (...a: Parameters<typeof pool>) => keep(await pool(...a)) }
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
