/**
 * A pool's threads compute in the locales loaded where the pool runs: those loaded when a thread starts come with it,
 * those loaded later are sent to every thread, a thread started again after a loss gets them all. A locale goes as its
 * data, which a structured clone carries, but its messages, which are functions: a thread imports the package's
 * module of the tag for them, and a locale outside the package speaks English there, as it does here. A session's
 * explanation, a window's formulas in the locale (`getViewport`) and its refusals are a synchronous session's. Before
 * ru-RU is loaded a pool refuses a ru-RU book as the same thread does. The file does not import `locales.ts`.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, InputError, createSession, registerLocale, type BookInput, type Calc } from '../src/index.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { sentLocale } from '../src/worker/protocol.js'
import { inside, pool } from './fixtures/pool.js'

const clock = { now: '2026-10-05T12:00:00Z', seed: 1 }
const cells = { A1: { formula: 'VALUE("1,5")+VALUE("1 234,5")' }, A2: { formula: '"a"&1.5&TRUE' }, A3: { formula: 'TEXT(45688.75,"ДД ММММ ГГГГ ч:мм")' } }
const book = (locale: string): BookInput => ({ locale, sheets: [{ name: 'S', cells }] })
const HINT = "book.locale: the locale ru-RU is not loaded: import { ruRU } from 'xlcalc/locales/ru-RU' and call registerLocale(ruRU) (loaded: en-US)"

const open: Calc[] = []
const made = async (...a: Parameters<typeof pool>) => {
  const p = await pool(...a)
  open.push(p)
  return p
}
afterAll(() => Promise.all(open.map((p) => p.close())))

/** What a synchronous session answers: its values, its explanation, a window's formula in the locale and its refusal of deleting the last sheet. */
function sync(b: BookInput, execution: 'closures' | 'eval') {
  const s = createSession(b, { clock, execution })
  const local = s.getViewport('S', 'A1:A3', { formulaLocal: true }).cells.A1.formulaLocal
  return { hash: s.getResult().hash, explain: s.explain('S', 'A1').text, local, refusal: s.editStructure({ op: 'deleteSheet', sheet: 'S' }) }
}

/** The same of a session in a pool. */
async function threaded(calc: Calc, b: BookInput, execution: 'closures' | 'eval') {
  const s = await calc.createSession(b, { clock, execution })
  const local = (await s.getViewport('S', 'A1:A3', { formulaLocal: true })).cells.A1.formulaLocal
  return { hash: (await s.getResult()).hash, explain: (await s.explain('S', 'A1')).text, local, refusal: await s.editStructure({ op: 'deleteSheet', sheet: 'S' }) }
}

describe('locales in a pool', () => {
  it('a locale is plain data but its messages: a structured clone carries what the pool sends', () => {
    expect(() => structuredClone(ruRU)).toThrow()
    const sent = sentLocale(ruRU)
    expect(sent.messages).toBeUndefined()
    expect(structuredClone(sent)).toEqual(sent)
  })

  it('a pool started before ru-RU is loaded refuses a ru-RU book as the same thread does', async () => {
    const calc = await made({ threads: 2 })
    const e = await calc.compute(book('ru-RU'), { clock }).catch((x: unknown) => x)
    expect(e).toBeInstanceOf(InputError)
    expect((e as InputError).message).toBe(HINT)
    await expect(calc.createSession(book('ru-RU'), { clock })).rejects.toThrow(HINT)
  })

  describe.each(['closures', 'eval'] as const)('execution %s', (execution) => {
    it('once ru-RU is loaded, the threads of a running pool compute in it, with its messages', async () => {
      const calc = open[0]
      registerLocale(ruRU)
      expect((await calc.compute(book('ru-RU'), { clock, execution })).hash).toBe(compute(book('ru-RU'), { clock, execution }).hash)
      // two sessions, one on each thread
      const want = sync(book('ru-RU'), execution)
      expect(want).toMatchObject({ local: 'ЗНАЧЕН("1,5")+ЗНАЧЕН("1 234,5")', refusal: { ok: false, message: 'В книге должен остаться хотя бы один лист' } })
      expect(await Promise.all([threaded(calc, book('ru-RU'), execution), threaded(calc, book('ru-RU'), execution)])).toEqual([want, want])
    })

    it('a pool started after ru-RU is loaded has it from the start, and a thread started again after a loss too', async () => {
      const calc = await made({ threads: 1 })
      expect(await threaded(calc, book('ru-RU'), execution)).toEqual(sync(book('ru-RU'), execution))
      const { calc: lost, crash } = await inside({ threads: 1 })
      open.push(lost)
      const s = await lost.createSession(book('ru-RU'), { clock, execution })
      crash(0)
      expect((await s.explain('S', 'A1')).text).toBe(sync(book('ru-RU'), execution).explain)
      expect((await lost.compute(book('ru-RU'), { clock, execution })).hash).toBe(compute(book('ru-RU'), { clock, execution }).hash)
    })
  })

  it('a locale outside the package goes as its data, and speaks English in a thread as it does here', async () => {
    const calc = open[0]
    registerLocale({ ...ruRU, tag: 'xx-XX', messages: undefined })
    const want = sync(book('xx-XX'), 'closures')
    expect(want.refusal).toMatchObject({ ok: false, message: 'a book keeps at least one sheet' })
    expect(await threaded(calc, book('xx-XX'), 'closures')).toEqual(want)
    expect((await calc.compute(book('xx-XX'), { clock })).sheets[0].cells).toEqual(compute(book('ru-RU'), { clock }).sheets[0].cells)
  })
})
