/**
 * Spilling against Excel en-US: each case of `spill/<set>.ts` is a small book whose formulas are written as
 * dynamic array formulas (Excel 365); Excel computes it and the values of the case's rectangles are read
 * back. `npx tsx bench/coverage/run.ts spill <set>` stores the answers in `test/excel/spill/<set>.json`;
 * `test/spill-excel.test.ts` compares the engine with them without Excel. The book and its reading through
 * mirror sheets: `case-book.ts`.
 * `FSPILL_ONLY=id,id` runs only those cases, `FCOV_FRESH=1` runs every case again.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { inCaseOrder } from './stored.js'
import { answerOf, BookRejected, runOnBook, type FileParts, type RectsAnswer } from './case-book.js'
import { ensureExcel, excelVersion } from './excel.js'
import type { SpillCase, SpillSet } from './spill/cases.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
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
function partsOf(c: SpillCase): FileParts | undefined {
  if (!c.legacy) return undefined
  const old = new Set(c.legacy)
  return { plain: { arrays: c.book.sheets.flatMap((s) => Object.keys(s.cells).map((a) => `${s.name}!${a}`).filter((k) => !old.has(k))) } }
}

/** One case in Excel: its book opened and computed, its edits applied, its rectangles read. */
async function runCase(c: SpillCase): Promise<SpillAnswer> {
  const kind = (v: NonNullable<SpillCase['edits']>[number]['value']) => (v === null ? 'x\t' : typeof v === 'number' ? `n\t${v}` : typeof v === 'string' ? `s\t${v}` : `f\t=${v.formula}`)
  const edits = (c.edits ?? []).map((e) => `${e.sheet}\t${e.cell}\t${kind(e.value)}`).join('\n')
  let out: string
  try {
    out = await runOnBook(c.id, SCRIPT, c.book, c.read, [edits, c.legacy ? '1' : ''], partsOf(c))
  } catch (e) {
    if (!(e instanceof BookRejected)) throw e
    console.log(e.message)
    return { values: {}, formulas: {}, rejected: true }
  }
  const raw = JSON.parse(out) as { rects: [unknown, unknown, unknown, unknown, unknown][]; edits: string[] }
  const answer: SpillAnswer = answerOf(c.read, raw.rects, !!c.legacy)
  raw.edits.forEach((e, i) => e && ((answer.refused ??= {})[i] = e))
  return answer
}

/** Runs the set's cases without a stored answer for the same case (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepSpill(set: string): Promise<void> {
  if (!/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
  const mod = (await import(`./spill/${set}.js`)) as SpillSet
  const cases = mod.buildSpillCases()
  const target = join(ROOT, `test/excel/spill/${set}.json`)
  const prev = existsSync(target) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(target, 'utf8')) as StoredSpill) : null
  const was = new Map(prev?.cases.map((c) => [c.id, JSON.stringify({ ...c, predicted: undefined })]) ?? [])
  const answers: Record<string, SpillAnswer> = {}
  for (const c of cases) if (prev?.answers[c.id] && was.get(c.id) === JSON.stringify({ ...c, predicted: undefined })) answers[c.id] = prev.answers[c.id]
  const t0 = Date.now()
  mkdirSync(join(ROOT, 'test/excel/spill'), { recursive: true })
  const save = () => {
    const meta = { version: excelVersion(), locale: 'en-US (-AppleLanguages (en) -AppleLocale en_US)', seconds: Math.round((Date.now() - t0) / 100) / 10 }
    writeFileSync(target, JSON.stringify({ meta, cases, answers: inCaseOrder(cases, answers) }, null, 1) + '\n')
  }
  await ensureExcel('en')
  let done = 0
  try {
    const only = process.env.FSPILL_ONLY?.split(',')
    for (const c of cases) {
      if (answers[c.id] || (only && !only.includes(c.id))) continue
      answers[c.id] = await runCase(c)
      done++
    }
  } finally {
    save()
    console.log(`spill ${set}: ${done} cases run, ${Object.keys(answers).length} of ${cases.length} answered, ${((Date.now() - t0) / 1000).toFixed(1)} s`)
  }
}
