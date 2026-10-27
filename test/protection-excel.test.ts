/**
 * Protection against Excel 16.114 for Mac en-US (`test/excel/protect/`, from `run.ts protect <set>`): what the object
 * model read after each case's steps (every sheet's protection, the read cells' locked and formula hidden) against the
 * engine's same steps over the book the reader reads from the file sent, and the file Excel saved read back: its sheets'
 * protection and the read cells' as the engine's book has them. Known differences are pinned with their reasons.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { ProtectAnswer, ProtectCases } from '../bench/coverage/protect/cases.js'
import { answerOf, caseFile, stepsBook } from '../bench/coverage/protect/cases.js'
import * as files from '../bench/coverage/protect/files.js'
import * as held from '../bench/coverage/protect/held.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import { readXlsx } from '../src/xlsx.js'
import { storedSets } from './stored-sets.js'
import { currentCases, pinnedCases } from './stored-cases.js'

type Stored = ProtectCases & { sent: Record<string, string>; answers: Record<string, ProtectAnswer>; rejected?: Record<string, string> }

/** Differences pinned by `<set>:<case>`. */
const KNOWN: Record<string, string> = {}

const sets = storedSets<Stored>('protect/')

/** The sets' modules by name. */
const MODULES: Record<string, { buildProtectCases(): ProtectCases }> = { files, held }

describe('protection against Excel for Mac', () => {
  it('every set is answered', () => expect(sets.map((s) => s.name)).toEqual(['files', 'held']))
  it('every pinned difference names a case', () => expect(pinnedCases(sets, KNOWN)).toEqual([]))

  for (const set of sets)
    describe(set.name, () => {
      const built = MODULES[set.name].buildProtectCases()
      currentCases(set, built.cases)
      it('Excel loaded every book, sent as the writer writes it now', async () => {
        expect(set.rejected ?? {}).toEqual({})
        for (const c of built.cases) expect(set.sent[c.id], c.id).toBe(await partsHash(await caseFile(built, c)))
      })
      for (const c of built.cases) {
        const a = set.answers[c.id]
        it.skipIf(KNOWN[`${set.name}:${c.id}`] !== undefined)(`${c.id}: Excel reads the protection the engine has after the steps, and saves it so`, async () => {
          const book = stepsBook((await readXlsx(await caseFile(built, c))).book, c.steps)
          const engine = answerOf(book, c.read)
          expect(a.error ?? '').toBe('')
          // an unprotected sheet's contents alone: the object model reads what it allows anyway
          const excel = Object.fromEntries(Object.entries(a.sheets).map(([name, g]) => [name, g.contents ? g : { contents: false }]))
          expect(excel).toEqual(engine.sheets)
          expect(a.cells).toEqual(engine.cells)
          const saved = (await readXlsx(new Uint8Array(readFileSync(new URL(`${c.id}.xlsx`, set.dir))))).book
          expect(saved.sheets.map((s) => [s.name, s.protection])).toEqual(book.sheets.map((s) => [s.name, s.protection]))
          expect(answerOf(saved, c.read).cells).toEqual(engine.cells)
        })
      }
    })
})
