/**
 * Tables against Excel: each case of `tables/<set>.ts` is a small book with tables, written by `writeXlsx` with three
 * empty mirror sheets; Excel computes it, takes the case's steps (`excel-tables.applescript`) and its rectangles are
 * read through mirror formulas written after the steps, with the formulas and the tables of every sheet.
 * `npx tsx bench/coverage/run.ts tables <set> [ru]` stores the answers in `test/excel/tables/[ru-RU/]<set>.json`;
 * `test/tables-excel.test.ts` compares the engine with them without Excel.
 * `FTABLES_ONLY=id,id` runs only those cases, `FTABLES_AGAIN=id,id` runs those again, `FCOV_FRESH=1` runs every case again.
 */
import { execFileSync, spawn } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { inCaseOrder } from './stored.js'
import { answerOf, MIRRORS, mirrorRanges } from './case-book.js'
import { declineRepair, dialogText, ensureExcel, excelDialogs, excelVersion, settleRejected, type Locale } from './excel.js'
import type { TableAnswer, TableCase, TableSet, TableShape, TableStep } from './tables/cases.js'
import type { BookInput } from '../../src/contract.js'
import { colToIndex, indexToCol } from '../../src/model.js'
import { sheetKey } from '../../src/syntax.js'
import { writeXlsx } from '../../src/xlsx.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-tables.applescript')
const SANDBOX = join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/tmp')
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

export interface StoredTables {
  meta: Record<string, string | number>
  cases: TableCase[]
  answers: Record<string, TableAnswer>
}

/** The stored answers of a set: `test/excel/tables/<set>.json`, in `ru-RU/` for ru. */
export const tablesFile = (set: string, locale: Locale = 'en') => join(ROOT, 'test/excel/tables', locale === 'ru' ? 'ru-RU' : '', `${set}.json`)

/** The case's steps as AppleScript lines; a table operation finds its table's sheet by name, renames followed. */
export function stepLines(c: TableCase): string[] {
  const where = new Map<string, string>()
  for (const s of c.book.sheets) for (const t of s.tables ?? []) where.set(sheetKey(t.name), s.name)
  const sheetOf = (table: string) => {
    const s = where.get(sheetKey(table))
    if (!s) throw new Error(`${c.id}: no table ${table}`)
    return s
  }
  const span = (col: string, n = 1) => `${col}1:${indexToCol(colToIndex(col) + n - 1)}1`
  return (c.steps ?? []).map((step: TableStep) => {
    if ('edit' in step) {
      const { sheet, cell, value: v } = step.edit
      const kind = v === null ? 'x\t' : typeof v === 'number' ? `n\t${v}` : typeof v === 'string' ? `s\t${v}` : 'local' in v ? `l\t${v.local}` : `f\t=${v.formula}`
      return `e\t${sheet}\t${cell}\t${kind}`
    }
    const o = step.op
    switch (o.op) {
      case 'insertRows':
      case 'deleteRows':
        return `${o.op === 'insertRows' ? 'insr' : 'delr'}\t${o.sheet}\tA${o.row}:A${o.row + (o.count ?? 1) - 1}`
      case 'insertColumns':
      case 'deleteColumns':
        return `${o.op === 'insertColumns' ? 'insc' : 'delc'}\t${o.sheet}\t${span(o.column, o.count)}`
      case 'moveRange':
        return `cut\t${o.from.sheet}\t${o.from.range}\t${o.to.sheet}\t${o.to.cell}`
      case 'copyRange':
        return `copy\t${o.from.sheet}\t${o.from.range}\t${o.to.sheet}\t${o.to.range}`
      case 'renameSheet':
        return `renamesheet\t${o.sheet}\t${o.name}`
      case 'deleteSheet':
        return `delsheet\t${o.sheet}`
      case 'renameTable': {
        const sheet = sheetOf(o.table)
        where.set(sheetKey(o.name), sheet)
        return `rename\t${sheet}\t${o.table}\t${o.name}`
      }
      case 'deleteTable':
        return `deltable\t${sheetOf(o.table)}\t${o.table}`
      case 'convertTable':
        return `convert\t${sheetOf(o.table)}\t${o.table}`
      default:
        throw new Error(`${c.id}: the tables step has no ${o.op}`)
    }
  })
}

/** The book with the three empty mirror sheets the script writes its mirror formulas into. */
export const withMirrors = (b: BookInput): BookInput => ({ ...b, sheets: [...b.sheets, ...MIRRORS.map((name) => ({ name, cells: {} }))] })

/** A book Excel did not open: the repair window, declined (activating Excel first: on another Space it is not listed). */
async function rejectedBook(child: ReturnType<typeof spawn>, exited: Promise<unknown>, id: string): Promise<TableAnswer | null> {
  if (!declineRepair()) {
    execFileSync('osascript', ['-e', 'tell application "Microsoft Excel" to activate'])
    await sleep(2000)
    if (!declineRepair()) return null
  }
  child.kill()
  await exited
  await settleRejected()
  console.log(`${id}: Excel did not load the book (repair window declined)`)
  return { values: {}, formulas: {}, tables: [], rejected: true }
}

/** One case in Excel. A window that is not a book stops the run (nothing is clicked); the repair window is declined. */
async function runCase(c: TableCase, local: boolean): Promise<TableAnswer> {
  const before = dialogText()
  if (before) throw new Error(`${c.id}: Excel has a window that is not a book: ${before}`)
  const dir = join(SANDBOX, `ftables-${process.pid}`)
  mkdirSync(dir, { recursive: true })
  try {
    const file = join(dir, 'book.xlsx')
    const { data, issues } = await writeXlsx(withMirrors(c.book))
    if (issues.length) throw new Error(`${c.id}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
    writeFileSync(file, data)
    const at = mirrorRanges(c.read)
    const rects = c.read.map((g, i) => `${g.sheet}\t${g.range}\t${at[i]}`).join('\n')
    const child = spawn('osascript', [SCRIPT, file, stepLines(c).join('\n'), rects, local ? '1' : ''])
    let out = ''
    let err = ''
    child.stdout.on('data', (d: Buffer) => (out += d.toString('utf8')))
    child.stderr.on('data', (d: Buffer) => (err += d.toString('utf8')))
    const exited = new Promise<number | null>((res) => child.on('close', res))
    const until = Date.now() + 120_000
    for (;;) {
      const code = await Promise.race([exited, sleep(1000).then(() => undefined)])
      if (code !== undefined) {
        const after = dialogText()
        if (after) throw new Error(`${c.id}: Excel window after the case: ${after}`)
        if (code !== 0) {
          // the book did not open (`wb` undefined): a repair window Excel shows on another Space
          if (/-2753|wb/.test(err)) {
            const r = await rejectedBook(child, exited, c.id)
            if (r) return r
          }
          throw new Error(`${c.id}: osascript ${code}: ${err.trim().split('\n')[0]}`)
        }
        const raw = JSON.parse(out) as { steps: string[]; rects: [unknown, unknown, unknown, unknown, unknown][]; tables: [string, string, string, boolean, boolean, string[]][] }
        const read = answerOf(c.read, raw.rects, local)
        const answer: TableAnswer = { values: read.values, formulas: read.formulas, ...(local ? { local: read.formulas2 } : {}), tables: raw.tables.map(([sheet, name, range, headers, totals, columns]): TableShape => ({ sheet, name, range: range.replaceAll('$', ''), headers, totals, columns })) }
        raw.steps.forEach((e, i) => e && ((answer.refused ??= {})[i] = e))
        return answer
      }
      const r = declineRepair() ? await rejectedBook(child, exited, c.id) : null
      if (r) return r
      if (excelDialogs().length) {
        child.kill()
        throw new Error(`${c.id}: Excel shows a window, stopped without clicking: ${dialogText()}`)
      }
      if (Date.now() > until) {
        child.kill()
        const hidden = await rejectedBook(child, exited, c.id)
        if (hidden) return hidden
        throw new Error(`${c.id}: Excel did not answer in 120 s`)
      }
    }
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

/** Runs the set's cases without a stored answer for the same case (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepTables(set: string, locale: Locale): Promise<void> {
  if (!/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
  const mod = (await import(`./tables/${set}.js`)) as TableSet
  const cases = mod.buildTableCases()
  const local = mod.local === true || locale === 'ru'
  const target = tablesFile(set, locale)
  const prev = existsSync(target) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(target, 'utf8')) as StoredTables) : null
  const was = new Map(prev?.cases.map((c) => [c.id, JSON.stringify(c)]) ?? [])
  const answers: Record<string, TableAnswer> = {}
  // FTABLES_AGAIN=id,id runs those cases again (a fix in the harness), keeping the other answers
  const again = new Set(process.env.FTABLES_AGAIN?.split(',') ?? [])
  for (const c of cases) if (prev?.answers[c.id] && was.get(c.id) === JSON.stringify(c) && !again.has(c.id)) answers[c.id] = prev.answers[c.id]
  const t0 = Date.now()
  mkdirSync(join(target, '..'), { recursive: true })
  const save = () => {
    const meta = { version: excelVersion(), locale: locale === 'ru' ? 'ru-RU (the system language and locale)' : 'en-US (-AppleLanguages (en) -AppleLocale en_US)', seconds: Math.round((Date.now() - t0) / 100) / 10 }
    writeFileSync(target, JSON.stringify({ meta, cases, answers: inCaseOrder(cases, answers) }, null, 1) + '\n')
  }
  await ensureExcel(locale)
  let done = 0
  try {
    const only = process.env.FTABLES_ONLY?.split(',')
    for (const c of cases) {
      if (answers[c.id] || (only && !only.includes(c.id))) continue
      answers[c.id] = await runCase(c, local)
      done++
    }
  } finally {
    save()
    console.log(`tables ${set}: ${done} cases run, ${Object.keys(answers).length} of ${cases.length} answered, ${((Date.now() - t0) / 1000).toFixed(1)} s`)
  }
}
