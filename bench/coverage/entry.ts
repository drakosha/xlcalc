/**
 * Typed text against Excel for Mac: the cases of `entry/<set>.ts` are written by `writeXlsx` into books of up to PART
 * cells (`entryBook`), Excel types each text into its cell in the user's language (`excel-entry.applescript`, `formula2
 * local`), reads what each cell shows, its format codes and the number it holds, and saves the book; the saved file is
 * read back by `readXlsx` for what each cell holds. `npx tsx bench/coverage/run.ts entry <set> [ru|<tag>]` stores the answers in
 * `test/excel/entry/[<tag>/]<set>.json` with the date of the run (a date typed without a year takes its year);
 * `test/entry-excel.test.ts` compares `parseEntry` with them without Excel. `FENTRY_ONLY=id,id` runs only those cases,
 * `FCOV_FRESH=1` every case again.
 */
import { spawn } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { CellInput } from '../../src/contract.js'
import { readXlsx, writeXlsx } from '../../src/xlsx.js'
import { declineRepair, dialogText, ensureExcel, excelDialogs, excelVersion, settleRejected } from './excel.js'
import { localeDir, localeMeta } from './locales.js'
import { inCaseOrder } from './stored.js'
import { entryBook, type EntryAnswer, type EntryCase, type EntrySet } from './entry/cases.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-entry.applescript')
const SANDBOX = join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/tmp')
const PART = 100
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

export interface StoredEntry {
  meta: Record<string, string | number>
  cases: EntryCase[]
  answers: Record<string, EntryAnswer>
}

/** The stored answers of a set: `test/excel/entry/<set>.json`, in a directory of its tag for a locale other than en-US. */
export const entryFile = (set: string, tag = 'en-US', win = false) => join(ROOT, 'test/excel', win ? 'win' : '', 'entry', localeDir(tag), `${set}.json`)

/** What a saved cell holds, as the engine takes it: its value or formula and its format (styles are not compared). */
export const heldOf = (c: CellInput | undefined): CellInput | null => {
  if (!c) return null
  const { value, formula, format } = c
  const held = { ...(formula !== undefined ? { formula } : value !== undefined && value !== null ? { value } : {}), ...(format !== undefined ? { format } : {}) }
  return Object.keys(held).length ? held : null
}

/** The answers of a part from Excel's reading of its cells and the file it saved. */
export async function partAnswers(part: EntryCase[], refused: string[], shown: [string, string, string, string?, unknown?, string?, string?, string?][], saved: Uint8Array): Promise<Record<string, EntryAnswer>> {
  const read = await readXlsx(saved)
  const cells = read.book.sheets.find((s) => s.name === 'S')!.cells
  const cached = read.cached.find((s) => s.name === 'S')!.cells
  const out: Record<string, EntryAnswer> = {}
  part.forEach((c, i) => {
    const at = `A${i + 1}`
    const [text, code, local, prefix, value2, formula, formulaLocal, bar] = shown[i]
    const cell = heldOf(cells[at])
    // a formula Excel does not take leaves the cell as it was, blank, with no error from AppleScript
    const why = refused[i] || (c.text !== '' && cell?.value === undefined && cell?.formula === undefined && !prefix ? 'not taken: the cell stayed blank' : '')
    out[c.id] = { ...(why ? { refused: why } : {}), cell, ...(at in cached ? { computed: cached[at] } : {}), text, code, local, ...(prefix ? { prefix } : {}), ...(typeof value2 === 'number' ? { value2 } : {}), ...(formula === undefined ? {} : { formula, formulaLocal }), ...(bar === undefined ? {} : { bar }) }
  })
  return out
}

/** One part in Excel. A window that is not a book stops the run (nothing is clicked); a book Excel repairs stops it too. */
async function runPart(part: EntryCase[], label: string, keys: boolean, edit: boolean): Promise<Record<string, EntryAnswer>> {
  if (declineRepair()) console.log('closed a leftover book repair window')
  const before = dialogText()
  if (before) throw new Error(`entry ${label}: Excel has a window that is not a book: ${before}`)
  const dir = join(SANDBOX, `fentry-${process.pid}`)
  mkdirSync(dir, { recursive: true })
  try {
    const file = join(dir, 'book.xlsx')
    const savedName = 'saved.xlsx'
    const savedFile = join(dir, savedName)
    const { data, issues } = await writeXlsx(entryBook(part))
    if (issues.length) throw new Error(`entry ${label}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
    writeFileSync(file, data)
    const child = spawn('osascript', [SCRIPT, file, savedFile, savedName, JSON.stringify(part.map((c) => c.text)), ...(keys ? ['keys', join(HERE, 'excel-keys.applescript')] : []), ...(edit ? ['edit'] : [])])
    let out = ''
    let err = ''
    child.stdout.on('data', (d: Buffer) => (out += d.toString('utf8')))
    child.stderr.on('data', (d: Buffer) => (err += d.toString('utf8')))
    const exited = new Promise<number | null>((res) => child.on('close', res))
    const until = Date.now() + 170_000
    for (;;) {
      const code = await Promise.race([exited, sleep(1000).then(() => undefined)])
      if (code !== undefined) {
        const after = dialogText()
        if (after) throw new Error(`entry ${label}: Excel window after the part: ${after}`)
        if (code !== 0) throw new Error(`entry ${label}: osascript ${code}: ${err.trim().split('\n')[0]}`)
        const raw = JSON.parse(out) as { refused: string[]; cells: [string, string, string, string?, unknown?, string?, string?, string?][] }
        return await partAnswers(part, raw.refused, raw.cells, new Uint8Array(readFileSync(savedFile)))
      }
      if (declineRepair()) {
        child.kill()
        await exited
        await settleRejected()
        throw new Error(`entry ${label}: Excel would repair the book`)
      }
      // typed on the keyboard, the script answers the window a refused formula brings itself
      if (!keys && excelDialogs().length) {
        child.kill()
        throw new Error(`entry ${label}: Excel shows a window, stopped without clicking: ${dialogText()}`)
      }
      if (Date.now() > until) {
        child.kill()
        throw new Error(`entry ${label}: Excel did not answer in 170 s`)
      }
    }
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

/** Runs the set's cases without a stored answer for the same case (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepEntry(set: string, tag: string): Promise<void> {
  if (!/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
  const mod = (await import(`./entry/${set}.js`)) as EntrySet
  const cases = mod.buildEntryCases()
  const target = entryFile(set, tag)
  const prev = existsSync(target) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(target, 'utf8')) as StoredEntry) : null
  const was = new Map(prev?.cases.map((c) => [c.id, JSON.stringify(c)]) ?? [])
  const answers: Record<string, EntryAnswer> = {}
  for (const c of cases) if (prev?.answers[c.id] && was.get(c.id) === JSON.stringify(c)) answers[c.id] = prev.answers[c.id]
  const only = process.env.FENTRY_ONLY?.split(',')
  const todo = cases.filter((c) => !answers[c.id] && (!only || only.includes(c.id)))
  const t0 = Date.now()
  mkdirSync(join(target, '..'), { recursive: true })
  const save = () => {
    const meta = { version: excelVersion(), locale: localeMeta(tag), date: new Date().toISOString().slice(0, 10), seconds: Math.round((Number(prev?.meta.seconds ?? 0) * 1000 + Date.now() - t0) / 100) / 10 }
    writeFileSync(target, JSON.stringify({ meta, cases, answers: inCaseOrder(cases, answers) }, null, 1) + '\n')
  }
  await ensureExcel(tag)
  try {
    for (let k = 0; k * PART < todo.length; k++) {
      const part = todo.slice(k * PART, (k + 1) * PART)
      Object.assign(answers, await runPart(part, `${set} ${tag} part ${k + 1}`, mod.keys === true, mod.edit === true))
      save()
      console.log(`entry ${set} ${tag}: part ${k + 1}, ${part.length} cases, ${((Date.now() - t0) / 1000).toFixed(1)} s`)
    }
  } finally {
    save()
    console.log(`entry ${set} ${tag}: ${todo.length} cases run, ${Object.keys(answers).length} of ${cases.length} answered`)
  }
}
