/**
 * The round trip: every formula of the coverage measurement, the Excel case sets and the probes, in the
 * engine's canonical form, shown in a locale and read back, is the same text. A formula the parser does
 * not take goes round as it is written.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { enUS, ruRU, toInvariant, toLocal, type Locale } from '../src/index.js'
import { canonical, parse } from '../src/syntax.js'
import { CASES, engineFormula } from './excel-cases.js'
import { SETS } from './excel-sets.js'

const probes = (JSON.parse(readFileSync(new URL('../docs/excel-probes.json', import.meta.url), 'utf8')) as { cases: { formula: string }[] }).cases

function canonicalOf(f: string): string {
  try {
    return canonical(parse(f))
  } catch {
    return f
  }
}

const FORMULAS = [...new Set([...CASES, ...SETS.flatMap((s) => s.cases), ...probes].map((c) => canonicalOf(engineFormula(c.formula))))]

describe('round trip through a locale', () => {
  it('the formulas: measurement, case sets, probes', () => {
    expect(FORMULAS.length).toBeGreaterThan(14_000)
  })

  for (const loc of [ruRU, enUS] as Locale[]) {
    it(`${loc.tag}: toInvariant(toLocal(f)) is f for every formula`, () => {
      const off = FORMULAS.flatMap((f) => {
        const shown = toLocal(f, loc)
        if (!shown.ok) return [`${f}: ${shown.message} at ${shown.at}`]
        const back = toInvariant(shown.formula, loc)
        return back.ok && back.formula === f ? [] : [`${f} → ${shown.formula} → ${back.ok ? back.formula : back.message}`]
      })
      expect(off.slice(0, 10)).toEqual([])
    })
  }

  it('ru-RU shows almost every formula differently (the round trip is not the identity)', () => {
    const same = FORMULAS.filter((f) => {
      const r = toLocal(f, ruRU)
      return r.ok && r.formula === f
    })
    expect(same.length / FORMULAS.length).toBeLessThan(0.1)
  })
})
