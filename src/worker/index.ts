/**
 * The subpath `xlcalc/worker`: `createCalcPool`, computation in a pool of threads (`docs/async.md`). In Node the
 * threads are `worker_threads`, elsewhere Web Workers; the root entry imports nothing of this.
 */
import { checkPoolOptions } from '../async/options.js'
import type { Calc, CalcPoolOptions, WorkerLike } from '../contract.js'
import { InputError } from '../input-error.js'
import { strict } from '../strict-args.js'
import { nodeSpawn, webSpawn, workerThreads, type ScriptUrl } from './platform.js'
import { createPool } from './pool.js'

export { CalcError } from '../exec/budget.js'
export type * from '../contract.js'

// the platform's globals, declared here: the package's types have no DOM; written as bundlers recognise a worker
declare const URL: new (url: string, base: string) => ScriptUrl
declare const Worker: new (url: ScriptUrl, opts: { type: 'module' }) => WorkerLike

/**
 * A pool of threads. Node runs `worker_threads` with `maxHeapMb` as each thread's heap; a browser runs Web Workers,
 * started by `spawn` or by the thread's module next to this one, and has no memory limit.
 */
export const createCalcPool = /* @__PURE__ */ strict(function createCalcPool(opts?: CalcPoolOptions): Calc {
  const o = checkPoolOptions(opts)
  const wt = workerThreads()
  if (o.maxHeapMb !== undefined && (o.spawn || !wt)) throw new InputError('opts.maxHeapMb', "a thread's memory limit needs the pool's own Node worker_threads; Web Workers and spawned workers have none")
  if (o.spawn) return createPool(o, webSpawn(o.spawn))
  if (wt) return createPool(o, nodeSpawn(wt, new URL('./thread.js', (import.meta as { url: string }).url), o.maxHeapMb))
  return createPool(o, webSpawn(() => new Worker(new URL('./thread.js', (import.meta as { url: string }).url), { type: 'module' })))
})
