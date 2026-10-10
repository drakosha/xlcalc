/**
 * Hidden sheets against Excel en-US: `npx tsx bench/coverage/run.ts hidden <set>` writes each book of `hidden/<set>.ts`
 * with `writeXlsx` (changed by hand for a raw case), Excel opens it, sets the visibility of sheets, applies the case's
 * sheet operation, reads each sheet's visibility, the active sheet and the read cells' values, and saves the book
 * (`excel-hidden.applescript`). The answers go to `test/excel/hidden/<set>.json` and the saved files to
 * `test/excel/hidden/<set>/<id>.xlsx`; `test/hidden-sheets-excel.test.ts` compares them with the engine and the reader
 * without Excel. A book Excel would repair is recorded in `rejected`. `FHIDDEN_ONLY=id,id` runs only those cases,
 * `FCOV_FRESH=1` every case again.
 */
import { spawn } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { declineRepair, dialogText, ensureExcel, excelVersion } from './excel.js'
import { hiddenFile, type HiddenAnswer, type HiddenCase, type HiddenCases, type HiddenSet } from './hidden/cases.js'
import { inCaseOrder } from './stored.js'
import { opArgs } from './struct.js'
import { partsHash } from './xlsx/sent.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-hidden.applescript')
const SANDBOX = join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/tmp')

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

/** The stored answers kept for a run: a case's, while it and its file are the same. */
export function keptAnswers(cases: HiddenCase[], sent: Record<string, string>, prev: StoredHidden | null, saved: (id: string) => boolean) {
  const key = (c: HiddenCase, hash: string | undefined) => JSON.stringify([c, hash])
  const same = new Map(prev?.cases.map((c) => [c.id, key(c, prev.sent[c.id])]) ?? [])
  const answers: Record<string, HiddenAnswer> = {}
  const rejected: Record<string, string> = {}
  for (const c of cases)
    if (same.get(c.id) === key(c, sent[c.id])) {
      if (prev!.answers[c.id] && saved(c.id)) answers[c.id] = prev!.answers[c.id]
      else if (prev!.rejected?.[c.id]) rejected[c.id] = prev!.rejected[c.id]
    }
  return { answers, rejected }
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

/** One case in Excel: its answer and the saved file, or the repair window's text. A window that is not a book stops the run. */
async function runCase(c: HiddenCase, data: Uint8Array, read: HiddenCases['read']): Promise<{ answer: HiddenAnswer; saved: Uint8Array } | { rejected: string }> {
  const before = dialogText()
  if (before) throw new Error(`${c.id}: Excel has a window that is not a book: ${before}`)
  const dir = join(SANDBOX, `fhidden-${process.pid}`)
  mkdirSync(dir, { recursive: true })
  try {
    const file = join(dir, 'book.xlsx')
    writeFileSync(file, data)
    const op = c.op ? opArgs({ id: c.id, book: c.book, op: c.op }) : []
    const makes = (c.make ?? []).map((m) => [m.sheet, m.visible])
    const child = spawn('osascript', [SCRIPT, file, join(dir, 'saved.xlsx'), 'saved.xlsx', JSON.stringify(makes), JSON.stringify(c.op ? [c.op.op, ...op] : []), read.sheet, JSON.stringify(read.cells)])
    let out = ''
    let err = ''
    child.stdout.on('data', (b: Buffer) => (out += b.toString('utf8')))
    child.stderr.on('data', (b: Buffer) => (err += b.toString('utf8')))
    const exited = new Promise<number | null>((res) => child.on('close', res))
    const until = Date.now() + 120_000
    for (;;) {
      const code = await Promise.race([exited, sleep(1000).then(() => undefined)])
      if (code !== undefined) {
        const after = dialogText()
        if (after) throw new Error(`${c.id}: Excel window after the case: ${after}`)
        if (code !== 0) throw new Error(`${c.id}: osascript ${code}: ${err.trim().split('\n')[0]}`)
        const raw = JSON.parse(out) as Omit<HiddenAnswer, 'error'> & { error: string }
        const { error, ...rest } = raw
        const answer: HiddenAnswer = { ...(error ? { error } : {}), ...rest, values: rest.values.map((v) => (v === '' ? null : v)) }
        return { answer, saved: new Uint8Array(readFileSync(join(dir, 'saved.xlsx'))) }
      }
      if (declineRepair()) {
        child.kill()
        await exited
        await sleep(1500)
        declineRepair()
        return { rejected: 'Excel did not accept the book (repair window declined)' }
      }
      const open = dialogText()
      if (open) {
        child.kill()
        throw new Error(`${c.id}: Excel shows a window, stopped without clicking: ${open}`)
      }
      if (Date.now() > until) {
        child.kill()
        throw new Error(`${c.id}: Excel did not answer in 120 s`)
      }
    }
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
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
