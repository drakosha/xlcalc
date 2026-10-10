/**
 * Notes, hyperlinks and frozen panes against Excel en-US: `npx tsx bench/coverage/run.ts parts <set>` writes each book
 * of `parts/<set>.ts` with `writeXlsx`, Excel opens it, makes what the case makes, applies its operation, reads every
 * sheet's parts and saves the book (`excel-parts.applescript`). The answers go to `test/excel/parts/<set>.json` and the
 * saved files to `test/excel/parts/<set>/<id>.xlsx`; `test/sheet-parts-excel.test.ts` compares them with the writer,
 * the reader and `restructure` without Excel. `FPARTS_ONLY=id,id` runs only those cases, `FCOV_FRESH=1` every case again.
 */
import { spawn } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { writeXlsx } from '../../src/xlsx.js'
import { declineRepair, dialogText, ensureExcel, excelVersion } from './excel.js'
import type { PartsAnswer, PartsCase, PartsCases, PartsMake, PartsSet } from './parts/cases.js'
import { inCaseOrder } from './stored.js'
import { threadedFile } from './parts/threaded.js'
import { opArgs } from './struct.js'
import { partsHash } from './xlsx/sent.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-parts.applescript')
const SANDBOX = join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/tmp')

export interface StoredParts extends PartsCases {
  meta: Record<string, string | number>
  /** What was sent for each case: `partsHash` of the file `writeXlsx` wrote. */
  sent: Record<string, string>
  answers: Record<string, PartsAnswer>
  /** Cases whose book Excel did not load: the repair window's text. */
  rejected?: Record<string, string>
}

/** What a case makes, as the script takes it. */
export function makeArgs(make: PartsMake[] = []): (string | number)[][] {
  return make.map((m) => {
    if ('note' in m) return ['note', m.sheet, m.note, m.text]
    if ('link' in m) return ['link', m.sheet, m.link, m.url ?? '', m.location ?? '', m.tooltip ?? '']
    return ['freeze', m.freeze, m.rows, m.columns]
  })
}

/** The file a case's book is written to, the same on both platforms: by `writeXlsx`, with the book's threads put in. */
export async function partsFile(c: PartsCase, set: PartsCases): Promise<Uint8Array> {
  const threads = set.threads?.[c.book]
  return threads ? threadedFile(set.books[c.book], threads) : (await writeXlsx(set.books[c.book])).data
}

type RawSheet = [string, [string, string, string][], [string, string, string, string, string][], [boolean, number, number, number, number, number, number]]

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))
const plain = (a: string) => a.replaceAll('$', '')

/** One case in Excel: its answer and the saved file, or the repair window's text. A window that is not a book stops the run. */
async function runCase(c: PartsCase, data: Uint8Array): Promise<{ answer: PartsAnswer; saved: Uint8Array } | { rejected: string }> {
  const before = dialogText()
  if (before) throw new Error(`${c.id}: Excel has a window that is not a book: ${before}`)
  const dir = join(SANDBOX, `fparts-${process.pid}`)
  mkdirSync(dir, { recursive: true })
  try {
    const file = join(dir, 'book.xlsx')
    writeFileSync(file, data)
    const op = c.op ? opArgs({ id: c.id, book: c.book, op: c.op }) : []
    const child = spawn('osascript', [SCRIPT, file, join(dir, 'saved.xlsx'), 'saved.xlsx', JSON.stringify(makeArgs(c.make)), JSON.stringify(c.op ? [c.op.op, ...op] : [])])
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
        const raw = JSON.parse(out) as { error: string; sheets: RawSheet[] }
        const answer: PartsAnswer = {
          ...(raw.error ? { error: raw.error } : {}),
          sheets: raw.sheets.map(([name, notes, links, [freeze, splitRow, splitColumn, , , scrollRow, scrollColumn]]) => ({
            name,
            notes: Object.fromEntries(notes.map(([a, author, text]) => [plain(a), { author, text }])),
            links: links.map(([range, address, subAddress, tip, display]) => ({ range: plain(range), address, subAddress, tip, display })),
            panes: { freeze, splitRow, splitColumn, scrollRow, scrollColumn },
          })),
        }
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

/** Runs the set's cases that have no stored answer for the same case and book (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepParts(set: string): Promise<void> {
  if (!/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
  const built = ((await import(`./parts/${set}.js`)) as PartsSet).buildPartsCases()
  const { books, cases, threads } = built
  const target = join(ROOT, `test/excel/parts/${set}.json`)
  const dir = join(ROOT, `test/excel/parts/${set}`)
  const prev = existsSync(target) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(target, 'utf8')) as StoredParts) : null
  const files = new Map(await Promise.all(cases.map(async (c) => [c.id, await partsFile(c, built)] as const)))
  const sent = Object.fromEntries(await Promise.all(cases.map(async (c) => [c.id, await partsHash(files.get(c.id)!)] as const)))
  // an answer is kept while its case and the file the writer writes for it are the same
  const key = (c: PartsCase, hash: string | undefined) => JSON.stringify([c, hash])
  const same = new Map(prev?.cases.map((c) => [c.id, key(c, prev.sent[c.id])]) ?? [])
  const answers: Record<string, PartsAnswer> = {}
  const rejected: Record<string, string> = {}
  for (const c of cases)
    if (same.get(c.id) === key(c, sent[c.id])) {
      if (prev!.answers[c.id] && existsSync(join(dir, `${c.id}.xlsx`))) answers[c.id] = prev!.answers[c.id]
      else if (prev!.rejected?.[c.id]) rejected[c.id] = prev!.rejected[c.id]
    }
  mkdirSync(dir, { recursive: true })
  const t0 = Date.now()
  const save = () => {
    const meta = { version: excelVersion(), locale: 'en-US (-AppleLanguages (en) -AppleLocale en_US)', seconds: Math.round((Date.now() - t0) / 100) / 10 }
    const refused = inCaseOrder(cases, rejected)
    writeFileSync(target, JSON.stringify({ meta, books, cases, ...(threads ? { threads } : {}), sent: inCaseOrder(cases, sent), answers: inCaseOrder(cases, answers), ...(Object.keys(refused).length ? { rejected: refused } : {}) }, null, 1) + '\n')
  }
  await ensureExcel('en-US')
  let done = 0
  try {
    const only = process.env.FPARTS_ONLY?.split(',')
    for (const c of cases) {
      if (answers[c.id] || rejected[c.id] || (only && !only.includes(c.id))) continue
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
    console.log(`parts ${set}: ${done} cases run, ${Object.keys(answers).length} of ${cases.length} answered, ${Object.keys(rejected).length} not loaded, ${((Date.now() - t0) / 1000).toFixed(1)} s`)
  }
}
