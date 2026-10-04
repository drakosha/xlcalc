/**
 * Оба варианта на короткой синтетической книге (340 строк данных): поячеечное совпадение, хэш не
 * зависит от порядка входа и числа прогонов, правки в сессии дают тот же хэш, что полный расчёт.
 */
import { describe, expect, it } from 'vitest'
import { cEngine } from '../../src/probe/c-engine'
import { compareValues } from '../../src/probe/compare'
import type { ProbeEngine } from '../../src/probe/engine'
import { hfEngine, makeHfEngine } from '../../src/probe/hf-engine'
import type { BookInput } from '../../src/probe/model'
import { classAEdits, classBEdit, mulberry32, S, synthBook } from '../../src/probe/synth'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const synth = synthBook({ seed: 1, rows: 340 })

function shuffled(book: BookInput, seed: number): BookInput {
  const rnd = mulberry32(seed)
  const mix = <T>(xs: T[]) => {
    const a = [...xs]
    for (let i = a.length - 1; i > 0; i--) {
      const j = Math.floor(rnd() * (i + 1))
      ;[a[i], a[j]] = [a[j], a[i]]
    }
    return a
  }
  return {
    sheets: mix(book.sheets).map((s) => ({ name: s.name, cells: Object.fromEntries(mix(Object.entries(s.cells))) })),
    names: mix(book.names ?? []),
  }
}

describe('варианты H и C на синтетике', () => {
  it('поячеечно совпадают, кроме TEXT от даты (formulajs)', () => {
    const h = hfEngine.compute(synth.book, { clock })
    const c = cEngine('idx').compute(synth.book, { clock })
    expect(h.invalid).toEqual([])
    expect(c.invalid).toEqual([])
    const cmp = compareValues(synth.book, h.values, c.values)
    expect(cmp.diffs.map((d) => `${d.sheet}!${d.addr}:${d.kind}`)).toEqual([`${S.params}!B15:TEXT`])
    expect(cmp.cells).toBeGreaterThan(9000)
  }, 60_000)

  it('H со своей SUMIFS совпадает с H', () => {
    const h = hfEngine.compute(synth.book, { clock })
    const hi = makeHfEngine(true).compute(synth.book, { clock })
    expect(compareValues(synth.book, h.values, hi.values).diffs).toEqual([])
  }, 60_000)

  for (const engine of [hfEngine, cEngine('idx')] as ProbeEngine[]) {
    it(`${engine.name}: один хэш при перемешанном входе и повторных прогонах`, () => {
      const hashes = new Set([0, 1, 2].map((i) => engine.compute(i === 0 ? synth.book : shuffled(synth.book, i), { clock }).hash))
      expect(hashes.size).toBe(1)
    }, 60_000)

    it(`${engine.name}: правки в сессии дают хэш полного расчёта`, () => {
      const session = engine.open(synth.book, { clock })
      let book = synth.book
      const edits = [...classAEdits(synth, 11, 15), classBEdit(synth, 12)]
      for (const e of edits) {
        session.setCells([e])
        book = { ...book, sheets: book.sheets.map((s) => (s.name === e.sheet ? { name: s.name, cells: { ...s.cells, [e.addr]: e.input } } : s)) }
        expect(session.hash()).toBe(engine.compute(book, { clock }).hash)
      }
      session.close()
    }, 120_000)
  }
})
