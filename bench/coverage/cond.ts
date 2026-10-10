/**
 * Conditional formats against Excel for Mac: `npx tsx bench/coverage/run.ts cond <set>` writes each book of
 * `cond/<set>.ts` with `writeXlsx`, Excel opens it in the set's locale, applies the case's operation, computes the book,
 * reads what the listed cells show by their display format and saves the book (`excel-cond.applescript`). The answers go
 * to `test/excel/cond/<set>.json` (another locale's to `cond/<tag>/`) and the saved files beside them;
 * `test/cond-excel.test.ts` compares them with the engine and the reader without Excel. Cases whose rules Excel makes
 * through its object model are Windows's alone (`bench/windows/cond.ts`): Excel for Mac's `make new format condition`
 * stopped Excel. `FCOND_ONLY=id,id` runs only those cases, `FCOV_FRESH=1` every case again.
 */
import { spawn } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { writeXlsx } from '../../src/xlsx.js'
import { fromRgb, type CondAnswer, type CondCase, type CondCases, type CondSet, type Look } from './cond/cases.js'
import { declineRepair, dialogText, ensureExcel, excelVersion } from './excel.js'
import { localeDir, localeMeta } from './locales.js'
import { inCaseOrder } from './stored.js'
import { opArgs } from './struct.js'
import { partsHash } from './xlsx/sent.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-cond.applescript')
const SANDBOX = join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/tmp')

export interface StoredCond extends CondCases {
  meta: Record<string, string | number>
  /** What was sent for each case: `partsHash` of the file `writeXlsx` wrote. */
  sent: Record<string, string>
  answers: Record<string, CondAnswer>
  /** Cases whose book Excel did not load. */
  rejected?: Record<string, string>
}

/** The set's cases and books, and where its answers go (`test/excel/<base>cond/[<tag>/]<set>`). */
export async function condSet(set: string): Promise<{ built: CondCases; tag: string; files: Map<string, Uint8Array>; sent: Record<string, string> }> {
  if (!/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
  const built = ((await import(`./cond/${set}.js`)) as CondSet).buildCondCases()
  const files = new Map(await Promise.all(built.cases.map(async (c) => [c.id, (await writeXlsx(built.books[c.book])).data] as const)))
  const sent = Object.fromEntries(await Promise.all(built.cases.map(async (c) => [c.id, await partsHash(files.get(c.id)!)] as const)))
  return { built, tag: built.locale ?? 'en-US', files, sent }
}

type RawLook = [address: string, text: string, fill: number[], colorIndex: string, font: number[], bold: boolean, italic: boolean]

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

/** One case in Excel: its answer and the saved file, or the repair window's text. A window that is not a book stops the run. */
async function runCase(c: CondCase, data: Uint8Array): Promise<{ answer: CondAnswer; saved: Uint8Array } | { rejected: string }> {
  const before = dialogText()
  if (before) throw new Error(`${c.id}: Excel has a window that is not a book: ${before}`)
  const dir = join(SANDBOX, `fcond-${process.pid}`)
  mkdirSync(dir, { recursive: true })
  try {
    const file = join(dir, 'book.xlsx')
    writeFileSync(file, data)
    const op = c.op ? opArgs({ id: c.id, book: c.book, op: c.op }) : []
    const child = spawn('osascript', [SCRIPT, file, join(dir, 'saved.xlsx'), 'saved.xlsx', JSON.stringify(c.op ? [c.op.op, ...op] : []), JSON.stringify(c.read)])
    let out = ''
    let err = ''
    child.stdout.on('data', (b: Buffer) => (out += b.toString('utf8')))
    child.stderr.on('data', (b: Buffer) => (err += b.toString('utf8')))
    const exited = new Promise<number | null>((res) => child.on('close', res))
    const until = Date.now() + 180_000
    for (;;) {
      const code = await Promise.race([exited, sleep(1000).then(() => undefined)])
      if (code !== undefined) {
        const after = dialogText()
        if (after) throw new Error(`${c.id}: Excel window after the case: ${after}`)
        if (code !== 0) throw new Error(`${c.id}: osascript ${code}: ${err.trim().split('\n')[0]}`)
        const raw = JSON.parse(out) as { error: string; looks: [string, RawLook[]][] }
        const looks: CondAnswer['looks'] = {}
        for (const [sheet, cells] of raw.looks)
          looks[sheet] = Object.fromEntries(
            cells.map(([a, text, fill, ci, font, bold, italic]): [string, Look] => [a.replaceAll('$', ''), { text, fill: ci === 'color index none' ? null : fromRgb(fill), font: fromRgb(font), bold, italic }]),
          )
        return { answer: { ...(raw.error ? { error: raw.error } : {}), looks }, saved: new Uint8Array(readFileSync(join(dir, 'saved.xlsx'))) }
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
        throw new Error(`${c.id}: Excel did not answer in 180 s`)
      }
    }
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

/** Runs the set's cases (but those whose rules Excel makes) that have no stored answer for the same case and file, and stores the answers. */
export async function stepCond(set: string): Promise<void> {
  const { built, tag, files, sent } = await condSet(set)
  const { books, cases } = built
  const base = `test/excel/cond/${localeDir(tag)}`
  const target = join(ROOT, `${base}${set}.json`)
  const dir = join(ROOT, `${base}${set}`)
  const prev = existsSync(target) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(target, 'utf8')) as StoredCond) : null
  const key = (c: CondCase, hash: string | undefined) => JSON.stringify([c, hash])
  const same = new Map(prev?.cases.map((c) => [c.id, key(c, prev.sent[c.id])]) ?? [])
  const answers: Record<string, CondAnswer> = {}
  const rejected: Record<string, string> = {}
  for (const c of cases)
    if (same.get(c.id) === key(c, sent[c.id])) {
      if (prev!.answers[c.id] && existsSync(join(dir, `${c.id}.xlsx`))) answers[c.id] = prev!.answers[c.id]
      else if (prev!.rejected?.[c.id]) rejected[c.id] = prev!.rejected[c.id]
    }
  const todo = cases.filter((c) => !c.make && !answers[c.id] && !rejected[c.id])
  if (!todo.length) return console.log(`cond ${set}: nothing to run on the Mac`)
  mkdirSync(dir, { recursive: true })
  const t0 = Date.now()
  const day = new Date().toLocaleDateString('sv-SE')
  const save = () => {
    const meta = { version: excelVersion(), locale: localeMeta(tag), day, seconds: Math.round((Date.now() - t0) / 100) / 10 }
    const refused = inCaseOrder(cases, rejected)
    writeFileSync(target, JSON.stringify({ meta, books, cases, ...(built.locale ? { locale: built.locale } : {}), sent: inCaseOrder(cases, sent), answers: inCaseOrder(cases, answers), ...(Object.keys(refused).length ? { rejected: refused } : {}) }, null, 1) + '\n')
  }
  await ensureExcel(tag)
  let done = 0
  try {
    const only = process.env.FCOND_ONLY?.split(',')
    for (const c of todo) {
      if (only && !only.includes(c.id)) continue
      const r = await runCase(c, files.get(c.id)!)
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
    console.log(`cond ${set}: ${done} cases run, ${Object.keys(answers).length} of ${cases.length} answered, ${Object.keys(rejected).length} not loaded, ${((Date.now() - t0) / 1000).toFixed(1)} s`)
  }
}
