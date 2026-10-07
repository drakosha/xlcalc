/**
 * Book-level features against Excel en-US: each case of `workbook/<set>.ts` is a small book whose file holds what
 * the model may not (hidden rows and columns, filters, the 1904 date system, iterative calculation); Excel opens it,
 * takes the case's steps (calculations, edits, hiding rows) and its rectangles are read at each `read` step.
 * `npx tsx bench/coverage/run.ts workbook <set>` stores the answers in `test/excel/workbook/<set>.json`;
 * `test/workbook-excel.test.ts` compares the engine with them. The book and its reading through mirror sheets:
 * `case-book.ts`. `FBOOK_ONLY=id,id` runs only those cases, `FCOV_FRESH=1` runs every case again.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { answerOf, runOnBook } from './case-book.js'
import { ensureExcel, excelVersion } from './excel.js'
import type { Step, WorkbookAnswer, WorkbookCase, WorkbookSet } from './workbook/cases.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-workbook.applescript')

export interface StoredWorkbook {
  meta: Record<string, string | number>
  cases: WorkbookCase[]
  answers: Record<string, WorkbookAnswer>
}

/** Where the answers of a set are stored. */
export const workbookFile = (set: string) => join(ROOT, `test/excel/workbook/${set}.json`)

/** A step as a line of the AppleScript's argument. */
function stepLine(s: Step): string {
  if (typeof s === 'string') return s
  if ('set' in s) {
    const { sheet, cell, value } = s.set
    return ['set', sheet, cell, value === null ? 'x' : typeof value === 'number' ? 'n' : 's', value ?? ''].join('\t')
  }
  if ('hide' in s) {
    const { sheet, rows, columns, hidden } = s.hide
    return rows !== undefined ? ['hide', sheet, `A${rows.replace(':', ':A')}`, hidden ? 1 : 0].join('\t') : ['hidecols', sheet, `${columns!.split(':')[0]}1:${columns!.split(':').at(-1)}1`, hidden ? 1 : 0].join('\t')
  }
  const { sheet, range, field, criterion } = s.filter
  return ['filter', sheet, range, field, criterion].join('\t')
}

/** One case in Excel: its book opened, its steps taken. */
async function runCase(c: WorkbookCase): Promise<WorkbookAnswer> {
  const steps = (c.steps ?? ['full', 'read']).map(stepLine).join('\n')
  const raw = JSON.parse(await runOnBook(c.id, SCRIPT, c.book, c.read, [steps], c.file)) as { settings: [boolean, number, number]; reads: [unknown, unknown, unknown, unknown][][] }
  const [iteration, count, delta] = raw.settings
  return { settings: { iteration, count, delta }, reads: raw.reads.map((r) => answerOf(c.read, r)) }
}

/** Runs the set's cases without a stored answer for the same case (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepWorkbook(set: string): Promise<void> {
  if (!/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
  const mod = (await import(`./workbook/${set}.js`)) as WorkbookSet
  const cases = mod.buildWorkbookCases()
  const target = workbookFile(set)
  const same = (c: WorkbookCase) => JSON.stringify({ ...c, predicted: undefined })
  const prev = existsSync(target) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(target, 'utf8')) as StoredWorkbook) : null
  const was = new Map(prev?.cases.map((c) => [c.id, same(c)]) ?? [])
  const answers: Record<string, WorkbookAnswer> = {}
  for (const c of cases) if (prev?.answers[c.id] && was.get(c.id) === same(c)) answers[c.id] = prev.answers[c.id]
  const t0 = Date.now()
  mkdirSync(join(target, '..'), { recursive: true })
  const save = () => {
    const meta = { version: excelVersion(), locale: 'en-US (-AppleLanguages (en) -AppleLocale en_US)', seconds: Math.round((Date.now() - t0) / 100) / 10 }
    writeFileSync(target, JSON.stringify({ meta, cases, answers }, null, 1) + '\n')
  }
  await ensureExcel('en')
  let done = 0
  try {
    const only = process.env.FBOOK_ONLY?.split(',')
    for (const c of cases) {
      if (answers[c.id] || (only && !only.includes(c.id))) continue
      answers[c.id] = await runCase(c)
      done++
    }
  } finally {
    save()
    console.log(`workbook ${set}: ${done} cases run, ${Object.keys(answers).length} of ${cases.length} answered, ${((Date.now() - t0) / 1000).toFixed(1)} s`)
  }
}
