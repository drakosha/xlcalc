/**
 * Threads of the pool on each platform: Node's `worker_threads` (taken with `process.getBuiltinModule`, so the module
 * has no `node:` import and a browser bundle takes nothing of Node), else Web Workers.
 */
import type { WorkerLike } from '../contract.js'
import type { FromThread, ToThread } from './protocol.js'

/** Why a thread went away by itself. */
export interface ThreadLoss {
  code: 'out_of_memory' | 'thread_failed'
  message: string
}

export interface Thread {
  post(m: ToThread): void
  terminate(): void
  /** Node: a thread keeps the process alive only while it works. */
  hold(on: boolean): void
}

export type Spawn = (onMessage: (m: FromThread) => void, onLoss: (l: ThreadLoss) => void) => Thread

interface NodeWorker {
  postMessage(m: unknown): void
  terminate(): Promise<number>
  ref(): void
  unref(): void
  on(event: 'message', f: (m: FromThread) => void): void
  on(event: 'error', f: (e: Error & { code?: string }) => void): void
  on(event: 'exit', f: (code: number) => void): void
}

/** A URL object of the platform (`new URL(...)`), for the thread's script. */
export interface ScriptUrl {
  readonly href: string
}

interface WorkerThreads {
  Worker: new (url: ScriptUrl, opts: { execArgv: string[]; resourceLimits?: { maxOldGenerationSizeMb?: number } }) => NodeWorker
}

/**
 * A word of memory shared with a thread, where the platform shares memory between threads: Node, and a browser page
 * that is cross-origin isolated. Elsewhere none (`docs/async.md`).
 */
export function sharedWord(): Int32Array | undefined {
  const g = globalThis as { SharedArrayBuffer?: typeof SharedArrayBuffer; crossOriginIsolated?: boolean }
  return g.SharedArrayBuffer && g.crossOriginIsolated !== false ? new Int32Array(new g.SharedArrayBuffer(4)) : undefined
}

/** Node's `worker_threads`, or undefined outside Node. */
export const workerThreads = (): WorkerThreads | undefined =>
  (globalThis as { process?: { getBuiltinModule?(id: string): unknown } }).process?.getBuiltinModule?.('node:worker_threads') as WorkerThreads | undefined

/** A Node thread running `script`, its heap limited to `maxHeapMb`. */
export function nodeSpawn(wt: WorkerThreads, script: ScriptUrl, maxHeapMb: number | undefined): Spawn {
  return (onMessage, onLoss) => {
    // the host's own flags (--input-type, --eval, loaders) are no flags of the thread's module
    const w = new wt.Worker(script, { execArgv: [], ...(maxHeapMb ? { resourceLimits: { maxOldGenerationSizeMb: maxHeapMb } } : {}) })
    let ended = false
    const lost = (l: ThreadLoss) => {
      if (ended) return
      ended = true
      onLoss(l)
    }
    w.on('message', (m) => !ended && onMessage(m))
    w.on('error', (e) => lost(e.code === 'ERR_WORKER_OUT_OF_MEMORY' ? { code: 'out_of_memory', message: `the thread reached its maxHeapMb of ${maxHeapMb}` } : { code: 'thread_failed', message: `the thread failed: ${e.message}` }))
    w.on('exit', (code) => lost({ code: 'thread_failed', message: `the thread exited with code ${code}` }))
    w.unref()
    return {
      post: (m) => w.postMessage(m),
      terminate() {
        ended = true
        void w.terminate()
      },
      hold: (on) => (on ? w.ref() : w.unref()),
    }
  }
}

/** A Web Worker from `make` (by default the thread's module next to this one). */
export function webSpawn(make: () => WorkerLike): Spawn {
  return (onMessage, onLoss) => {
    const w = make()
    let ended = false
    w.addEventListener('message', (e) => !ended && onMessage(e.data as FromThread))
    w.addEventListener('error', (e) => {
      if (ended) return
      ended = true
      onLoss({ code: 'thread_failed', message: `the thread failed: ${e.message ?? 'an error'}` })
    })
    return {
      post: (m) => w.postMessage(m),
      terminate() {
        ended = true
        w.terminate()
      },
      hold() {},
    }
  }
}
