/**
 * Hidden sheets against Excel en-US: `npx tsx bench/coverage/run.ts hidden <set>` writes each book of `hidden/<set>.ts`
 * with `writeXlsx` (changed by hand for a raw case), Excel opens it, sets the visibility of sheets, applies the case's
 * sheet operation, reads each sheet's visibility, the active sheet and the read cells' values, and saves the book
 * (`excel-hidden.applescript`). The answers go to `test/excel/hidden/<set>.json` and the saved files to
 * `test/excel/hidden/<set>/<id>.xlsx`; `test/hidden-sheets-excel.test.ts` compares them with the engine and the reader
 * without Excel. A book Excel would repair is recorded in `rejected`. `FHIDDEN_ONLY=id,id` runs only those cases,
 * `FCOV_FRESH=1` every case again.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { ensureExcel, excelVersion, runSavingScript } from './excel.js'
import { hiddenFile, type HiddenAnswer, type HiddenCase, type HiddenCases, type HiddenSet } from './hidden/cases.js'
import { inCaseOrder, keptAnswers } from './stored.js'
import { opArgs } from './struct.js'
import { partsHash } from './xlsx/sent.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-hidden.applescript')

export interface StoredHidden extends HiddenCases {
  meta: Record<string, string | number>
  /** What was sent for each case: `partsHash` of its file. */
  sent: Record<string, string>
  answers: Record<string, HiddenAnswer>
  /** Cases whose book Excel did not load: why. */
  rejected?: Record<string, string>
}

/** The files of a set's cases and what identifies them. */
export async function hiddenFiles(built: HiddenCases): Promise<{ files: Map<string, Uint8Array>; sent: Record<string, string> }> {
  const files = new Map(await Promise.all(built.cases.map(async (c) => [c.id, await hiddenFile(c, built)] as const)))
  const sent = Object.fromEntries(await Promise.all(built.cases.map(async (c) => [c.id, await partsHash(files.get(c.id)!)] as const)))
  return { files, sent }
}

/** One case in Excel: its answer and the saved file, or the repair window's text. A window that is not a book stops the run. */
async function runCase(c: HiddenCase, data: Uint8Array, read: HiddenCases['read']): Promise<{ answer: HiddenAnswer; saved: Uint8Array } | { rejected: string }> {
  const op = c.op ? opArgs({ id: c.id, book: c.book, op: c.op }) : []
  const makes = (c.make ?? []).map((m) => [m.sheet, m.visible])
  const r = await runSavingScript(c.id, SCRIPT, data, [JSON.stringify(makes), JSON.stringify(c.op ? [c.op.op, ...op] : []), read.sheet, JSON.stringify(read.cells)], 120_000)
  if ('rejected' in r) return r
  const { error, ...rest } = JSON.parse(r.out) as Omit<HiddenAnswer, 'error'> & { error: string }
  return { answer: { ...(error ? { error } : {}), ...rest, values: rest.values.map((v) => (v === '' ? null : v)) }, saved: r.saved }
}

/** Runs the set's cases that have no stored answer for the same case and file (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepHidden(set: string): Promise<void> {
  if (!/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
  const built = ((await import(`./hidden/${set}.js`)) as HiddenSet).buildHiddenCases()
  const target = join(ROOT, `test/excel/hidden/${set}.json`)
  const dir = join(ROOT, `test/excel/hidden/${set}`)
  const prev = existsSync(target) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(target, 'utf8')) as StoredHidden) : null
  const { files, sent } = await hiddenFiles(built)
  const { answers, rejected } = keptAnswers(built.cases, sent, prev, (id) => existsSync(join(dir, `${id}.xlsx`)))
  mkdirSync(dir, { recursive: true })
  const t0 = Date.now()
  const save = () => {
    const meta = { version: excelVersion(), locale: 'en-US (-AppleLanguages (en) -AppleLocale en_US)', seconds: Math.round((Date.now() - t0) / 100) / 10 }
    const refused = inCaseOrder(built.cases, rejected)
    writeFileSync(target, JSON.stringify({ meta, ...built, sent: inCaseOrder(built.cases, sent), answers: inCaseOrder(built.cases, answers), ...(Object.keys(refused).length ? { rejected: refused } : {}) }, null, 1) + '\n')
  }
  await ensureExcel('en-US')
  let done = 0
  try {
    const only = process.env.FHIDDEN_ONLY?.split(',')
    for (const c of built.cases) {
      if (answers[c.id] || rejected[c.id] || (only && !only.includes(c.id))) continue
      const r = await runCase(c, files.get(c.id)!, built.read)
      if ('rejected' in r) {
        rejected[c.id] = r.rejected
        console.log(`${c.id}: ${r.rejected}`)
      } else {
        answers[c.id] = r.answer
        writeFileSync(join(dir, `${c.id}.xlsx`), r.saved)
      }
      done++
      save()
    }
  } finally {
    save()
    console.log(`hidden ${set}: ${done} cases run, ${Object.keys(answers).length} of ${built.cases.length} answered, ${Object.keys(rejected).length} not loaded, ${((Date.now() - t0) / 1000).toFixed(1)} s`)
  }
}
