/** `getCells` of an asynchronous session takes `signal` and `timeoutMs`, as a calculation does: a signal aborted before the call stops it, in the same thread and in a thread. */
import { afterAll, describe, expect, it } from 'vitest'
import { ASYNC_CALCS } from './session-kinds.js'
import type { Calc } from '../src/index.js'

describe.each(ASYNC_CALCS)('getCells %s', (_, make) => {
  let calc: Calc
  afterAll(() => calc?.close())
  it('answers as before, and stops with aborted on a signal aborted', async () => {
    calc = await make()
    const s = await calc.createSession({ sheets: [{ name: 'S', cells: { A1: { value: 2 }, B1: { formula: 'A1*3' } } }] })
    expect((await s.getCells('S', 'A1:B1', { timeoutMs: 5000 })).B1).toMatchObject({ value: 6 })
    await expect(s.getCells('S', 'A1:B1', { signal: AbortSignal.abort() })).rejects.toMatchObject({ code: 'aborted' })
    await expect(s.getCells('S', 'A1', { rows: [1] } as never)).rejects.toThrow('opts')
    await s.close()
  })
})
