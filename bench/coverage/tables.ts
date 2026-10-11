/**
 * Tables against Excel: each case of `tables/<set>.ts` is a small book with tables, written by `writeXlsx` with three
 * empty mirror sheets; Excel computes it, takes the case's steps (`excel-tables.applescript`) and its rectangles are
 * read through mirror formulas written after the steps, with the formulas and the tables of every sheet.
 * `npx tsx bench/coverage/run.ts tables <set> [ru|<tag>]` stores the answers in `test/excel/tables/[ru-RU/ or <tag>/]<set>.json`;
 * `test/tables-excel.test.ts` compares the engine with them without Excel.
 * `FTABLES_ONLY=id,id` runs only those cases, `FTABLES_AGAIN=id,id` runs those again, `FCOV_FRESH=1` runs every case again.
 */
import { execFileSync } from 'node:child_process'
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { answeredLine, inCaseOrder, keptSame, loadStored, runCases, secondsSince, setModule, took, writeAnswers } from './stored.js'
import { answerOf, MIRRORS, mirrorRanges } from './case-book.js'
import { declineRepair, ensureExcel, excelVersion, runWatched, settleRejected } from './excel.js'
import { sleep } from './proc.js'
import { localeDir, localeMeta } from './locales.js'
import type { TableAnswer, TableCase, TableSet, TableShape, TableStep } from './tables/cases.js'
import type { BookInput } from '../../src/contract.js'
import { colToIndex, indexToCol } from '../../src/model.js'
import { sheetKey } from '../../src/syntax.js'
import { writeXlsx } from '../../src/xlsx.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-tables.applescript')

export interface StoredTables {
  meta: Record<string, string | number>
  cases: TableCase[]
  answers: Record<string, TableAnswer>
}

/** The stored answers of a set: `test/excel/tables/<set>.json`, in a directory of its tag for a locale other than en-US (`ru-RU/`). */
export const tablesFile = (set: string, tag = 'en-US') => join(ROOT, 'test/excel/tables', localeDir(tag), `${set}.json`)

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
    if ('name' in step) return `name\t${step.name.scope ?? ''}\t${step.name.name}\t=${step.name.ref}`
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

/** A case that adds names: its answer lists the names Excel has after the steps. */
export const addsNames = (c: TableCase): boolean => (c.steps ?? []).some((s) => 'name' in s)

/** The book with the three empty mirror sheets the script writes its mirror formulas into. */
export const withMirrors = (b: BookInput): BookInput => ({ ...b, sheets: [...b.sheets, ...MIRRORS.map((name) => ({ name, cells: {} }))] })

/**
 * A book Excel did not open: the repair window, declined (activating Excel first: on another Space it is not listed),
 * or, with `refused`, an error from `open workbook` with no window at all.
 */
async function rejectedBook(stop: () => Promise<void>, id: string, refused = false): Promise<TableAnswer | undefined> {
  let declined = declineRepair()
  if (!declined) {
    execFileSync('osascript', ['-e', 'tell application "Microsoft Excel" to activate'])
    await sleep(2000)
    declined = declineRepair()
    if (!declined && !refused) return undefined
  }
  await stop()
  await settleRejected()
  console.log(`${id}: Excel did not load the book (${declined ? 'repair window declined' : 'open workbook refused it, no window'})`)
  return { values: {}, formulas: {}, tables: [], rejected: true }
}

/** One case in Excel. A window that is not a book stops the run (nothing is clicked); the repair window is declined. */
async function runCase(c: TableCase, local: boolean): Promise<TableAnswer> {
  const at = mirrorRanges(c.read)
  const rects = c.read.map((g, i) => `${g.sheet}\t${g.range}\t${at[i]}`).join('\n')
  return runWatched<TableAnswer>({
    script: SCRIPT,
    label: c.id,
    folder: 'ftables',
    book: async (file) => {
      const { data, issues } = await writeXlsx({ book: withMirrors(c.book) }, { values: 'omit' })
      if (issues.length) throw new Error(`${c.id}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
      writeFileSync(file, data)
    },
    args: () => [stepLines(c).join('\n'), rects, local ? '1' : ''],
    timeout: 120_000,
    poll: 1000,
    after: 'case',
    fatal: true,
    // declined already, the window is looked for once more after Excel is activated; the script then fails on the book
    repair: (stop) => rejectedBook(stop, c.id),
    // the book did not open (`wb` undefined, an error from `open workbook`, a timeout): a repair window on another Space
    failed: (err, stop) => rejectedBook(stop, c.id, /the book did not open/.test(err)),
    late: (stop) => rejectedBook(stop, c.id),
    done: (out) => {
      const raw = JSON.parse(out) as { steps: string[]; rects: [unknown, unknown, unknown, unknown, unknown][]; tables: [string, string, string, boolean, boolean, string[]][]; names: string[] }
      const read = answerOf(c.read, raw.rects, local)
      const answer: TableAnswer = { values: read.values, formulas: read.formulas, ...(local ? { local: read.formulas2 } : {}), tables: raw.tables.map(([sheet, name, range, headers, totals, columns]): TableShape => ({ sheet, name, range: range.replaceAll('$', ''), headers, totals, columns })) }
      if (addsNames(c)) answer.names = raw.names
      raw.steps.forEach((e, i) => e && ((answer.refused ??= {})[i] = e))
      return answer
    },
  })
}

/** Runs the set's cases without a stored answer for the same case (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepTables(set: string, tag: string): Promise<void> {
  const mod = await setModule<TableSet>('tables', set)
  const cases = mod.buildTableCases()
  // formulas typed and read in the user's language in every locale but en-US
  const local = mod.local === true || tag !== 'en-US'
  const target = tablesFile(set, tag)
  const answers = keptSame({ cases }, loadStored<StoredTables>(target))
  // FTABLES_AGAIN=id,id runs those cases again (a fix in the harness), keeping the other answers
  for (const id of process.env.FTABLES_AGAIN?.split(',') ?? []) delete answers[id]
  const t0 = Date.now()
  const save = () => writeAnswers(target, { meta: { version: excelVersion(), locale: localeMeta(tag), seconds: secondsSince(t0) }, cases, answers: inCaseOrder(cases, answers) })
  await ensureExcel(tag)
  await runCases(cases.filter((c) => !answers[c.id]), 'FTABLES_ONLY', async (c) => void (answers[c.id] = await runCase(c, local)), (done) => {
    save()
    console.log(`${answeredLine(`tables ${set}`, done, answers, cases)}, ${took(t0)} s`)
  })
}
