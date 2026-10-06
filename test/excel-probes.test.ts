/**
 * Acceptance against Excel for rules outside single functions: numbers as text (`&`, CONCAT, text
 * arguments, TEXT with General), number literals and numeric text longer than 15 significant digits,
 * criteria over whole columns of sheets with different numbers of rows. The cases come from
 * `bench/coverage/probes.ts`; Excel 16.113's answers (en-US locale, English UI) are stored in
 * `docs/excel-probes.json` (`npx tsx bench/coverage/run.ts probes` refreshes them). Comparison as in the
 * coverage measurement.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { compare } from '../bench/coverage/compare.js'
import type { Val } from '../bench/coverage/formulajs.js'
import { buildProbes, PROBE_SHEETS, type Probe } from '../bench/coverage/probes.js'
import { compute } from '../src/index.js'
import { CLOCK, show, toVal } from './excel-cases.js'
import { caseBook } from './excel-sets.js'

const stored = JSON.parse(readFileSync(new URL('../docs/excel-probes.json', import.meta.url), 'utf8')) as { cases: Probe[]; values: Record<string, Val> }
const probes = buildProbes()

const result = compute(caseBook(probes, PROBE_SHEETS), { clock: CLOCK }).sheets.at(-1)!.cells

describe('acceptance against Excel en-US: probes', () => {
  it('the stored Excel answers are for the current probes', () => {
    expect(stored.cases).toEqual(probes)
    expect(probes.filter((p) => !stored.values[p.id])).toEqual([])
  })

  for (const group of [...new Set(probes.map((p) => p.group))]) {
    it(`${group}: all probes match Excel`, () => {
      const off = probes
        .map((p, i) => ({ p, engine: toVal(result[`A${i + 1}`].value), excel: stored.values[p.id] }))
        .filter((o) => o.p.group === group && !compare(o.excel, o.engine).ok)
        .map((o) => `${o.p.id} ${o.p.formula}: Excel ${show(o.excel)}, engine ${show(o.engine)}`)
      expect(off).toEqual([])
    })
  }
})
