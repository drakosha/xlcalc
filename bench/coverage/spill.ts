/**
 * Spilling against Excel en-US: each case of `spill/<set>.ts` is a small book whose formulas are written as
 * dynamic array formulas (Excel 365); Excel computes it and the values of the case's rectangles are read
 * back. `npx tsx bench/coverage/run.ts spill <set>` stores the answers in `test/excel/spill/<set>.json`;
 * `test/spill-excel.test.ts` compares the engine with them without Excel. The book and its reading through
 * mirror sheets: `case-book.ts`.
 * `FSPILL_ONLY=id,id` runs only those cases, `FCOV_FRESH=1` runs every case again.
 */
import { join } from 'node:path'
import { answeredLine, inCaseOrder, keptSame, loadStored, runCases, secondsSince, setModule, took, unpredicted, writeAnswers } from './stored.js'
import { answerOf, BookRejected, editOf, MIRRORS, refusedOf, runOnBook, type FileParts, type RectsAnswer } from './case-book.js'
import { ensureExcel, excelVersion } from './excel.js'
import { localeMeta } from './locales.js'
import type { SpillCase, SpillSet } from './spill/cases.js'
import { HERE, ROOT } from './paths.js'

const SCRIPT = join(HERE, 'excel-spill.applescript')

export interface SpillAnswer extends RectsAnswer {
  /** A case with cells of the old kind (`legacy`): the formulas as Excel 365 shows them (`formula2`, `@` where it intersects). */
  formulas2?: Record<string, string>
  /** The edits Excel refused, by their index in `edits`: its error. */
  refused?: Record<number, string>
  /** Excel did not load the book ("We found a problem with some content", declined): nothing else is answered. */
  rejected?: true
}

export interface StoredSpill {
  meta: Record<string, string | number>
  cases: SpillCase[]
  answers: Record<string, SpillAnswer>
}

/** The file of a case with cells of the old kind: those as plain formulas, the others as dynamic array formulas. */
export function partsOf(c: SpillCase): FileParts | undefined {
  if (!c.legacy) return undefined
  const old = new Set(c.legacy)
  return { plain: { arrays: c.book.sheets.flatMap((s) => Object.keys(s.cells).map((a) => `${s.name}!${a}`).filter((k) => !old.has(k))) } }
}

/** One case in Excel: its book opened and computed, its edits applied, its rectangles read. */
async function runCase(c: SpillCase): Promise<SpillAnswer> {
  const edits = (c.edits ?? []).map((e) => [e.sheet, e.cell, ...editOf(e.value)].join('\t')).join('\n')
  let out: string
  try {
    out = await runOnBook(c.id, SCRIPT, c.book, c.read, [edits, c.legacy ? '1' : '', MIRRORS.join('\t')], partsOf(c))
  } catch (e) {
    if (!(e instanceof BookRejected)) throw e
    console.log(e.message)
    return { values: {}, formulas: {}, rejected: true }
  }
  const raw = JSON.parse(out) as { rects: [unknown, unknown, unknown, unknown, unknown][]; edits: string[] }
  const answer: SpillAnswer = answerOf(c.read, raw.rects, !!c.legacy)
  const refused = refusedOf(raw.edits)
  return refused ? { ...answer, refused } : answer
}

/** Runs the set's cases without a stored answer for the same case (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepSpill(set: string): Promise<void> {
  const cases = (await setModule<SpillSet>('spill', set)).buildSpillCases()
  const target = join(ROOT, `test/excel/spill/${set}.json`)
  const answers = keptSame({ cases }, loadStored<StoredSpill>(target), unpredicted)
  const t0 = Date.now()
  const save = () => writeAnswers(target, { meta: { version: excelVersion(), locale: localeMeta('en-US'), seconds: secondsSince(t0) }, cases, answers: inCaseOrder(cases, answers) })
  await ensureExcel('en-US')
  await runCases(cases.filter((c) => !answers[c.id]), 'FSPILL_ONLY', async (c) => void (answers[c.id] = await runCase(c)), (done) => {
    save()
    console.log(`${answeredLine(`spill ${set}`, done, answers, cases)}, ${took(t0)} s`)
  })
}
