/** Calls taken in order: each starts when the one before has finished; an abort rejects a waiting call at once. */
import type { AbortSignalLike } from '../contract.js'
import { CalcError } from '../exec/budget.js'

export const abortError = (signal: AbortSignalLike): CalcError => new CalcError('aborted', 'the call was aborted', { cause: signal.reason })

export class Serial {
  private tail: Promise<void> = Promise.resolve()

  run<T>(task: () => Promise<T>, signal?: AbortSignalLike): Promise<T> {
    return new Promise<T>((resolve, reject) => {
      if (signal?.aborted) return reject(abortError(signal))
      let skip = false
      const onAbort = () => {
        skip = true
        reject(abortError(signal!))
      }
      signal?.addEventListener('abort', onAbort, { once: true })
      this.tail = this.tail.then(async () => {
        signal?.removeEventListener('abort', onAbort)
        if (skip) return
        try {
          resolve(await task())
        } catch (e) {
          reject(e)
        }
      })
    })
  }
}
