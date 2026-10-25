/**
 * A session's subscribers (`Session.subscribe`): after a call that changed the book, each listener gets what it answered
 * of `EditResult` and the call's name. The three kinds of session share this: each tells it a changing call's answer,
 * and a change is told once, by its revision. A listener's own call is told after the change it answers, so changes
 * reach every listener in the order of their revisions; a listener that throws stops neither the call nor the others.
 */
import type { ChangingMethod, EditResult, SessionChange } from '../contract.js'
import { InputError, what } from '../input-error.js'
import { strict } from '../strict-args.js'

export type Listener = (change: SessionChange) => void

/** The platform's `queueMicrotask` (both Node and browsers have it). */
const g = globalThis as unknown as { queueMicrotask(f: () => void): void }

/** A listener as a call's `path` argument takes it, refused where it is not a function. */
export const listenerOf = (path: string, listener: unknown): Listener => {
  if (typeof listener !== 'function') throw new InputError(path, `listener must be a function, got ${what(listener)}`)
  return listener as Listener
}

/** Tells `listener` the `change`: its own throw reaches the host as a microtask, stopping neither the call nor the other listeners. */
export const tell = (listener: Listener, change: SessionChange) => {
  try {
    listener(change)
  } catch (e) {
    g.queueMicrotask(() => {
      throw e
    })
  }
}

/** The fields of `EditResult` a change carries, out of a call's answer (`CellTextResult`, `UndoResult`…). */
function changeOf(method: ChangingMethod, r: EditResult): SessionChange {
  const c: SessionChange = { method, revision: r.revision, changes: r.changes }
  if (r.layout) c.layout = r.layout
  if (r.styles) c.styles = r.styles
  if (r.moved) c.moved = r.moved
  if (r.repaint) c.repaint = r.repaint
  return c
}

/** The subscribers of a session whose revision is `revision` now. */
export function subscribers(revision: number) {
  const listeners = new Set<Listener>()
  const waiting: SessionChange[] = []
  let telling = false
  return {
    subscribe(listener: Listener): () => void {
      listenerOf('subscribe', listener)
      listeners.add(listener)
      return () => {
        listeners.delete(listener)
      }
    },
    /** A call's answer: told when it is a change of the book (its revision past the last one told). */
    after<T>(method: ChangingMethod, answer: T): T {
      const r = answer as Partial<EditResult> | null
      if (!r || typeof r.revision !== 'number' || r.revision <= revision || !Array.isArray(r.changes)) return answer
      revision = r.revision
      if (!listeners.size) return answer
      waiting.push(changeOf(method, r as EditResult))
      // a change a listener makes waits until the one it answers has reached every listener
      if (telling) return answer
      telling = true
      try {
        for (let c = waiting.shift(); c; c = waiting.shift()) {
          for (const listener of [...listeners]) tell(listener, c)
        }
      } finally {
        telling = false
      }
      return answer
    },
    clear() {
      listeners.clear()
      waiting.length = 0
    },
  }
}

/** `api` (an asynchronous session, its methods strict and answering promises) with `subscribe` strict as a synchronous call. */
export function withSubscribe<T extends { subscribe: (listener: Listener) => () => void }>(api: T, subs: ReturnType<typeof subscribers>): T {
  api.subscribe = strict((listener: Listener) => subs.subscribe(listener), 'subscribe', 1, false)
  return api
}
