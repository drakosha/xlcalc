/**
 * Book-level features against Excel en-US: each case of `workbook/<set>.ts` is a small book whose file holds what
 * the model may not (hidden rows and columns, filters, the 1904 date system, iterative calculation); Excel opens it,
 * takes the case's steps (calculations, edits, hiding rows) and its rectangles are read at each `read` step.
 * `npx tsx bench/coverage/run.ts workbook <set>` stores the answers in `test/excel/workbook/<set>.json`;
 * `test/workbook-excel.test.ts` compares the engine with them. The book and its reading through mirror sheets:
 * `case-book.ts`. `FBOOK_ONLY=id,id` runs only those cases, `FCOV_FRESH=1` runs every case again.
 */
import { join } from 'node:path'
import { answeredLine, inCaseOrder, keptSame, loadStored, runCases, secondsSince, setModule, takenBook, took, unpredicted, writeAnswers } from './stored.js'
import { answerOf, cellsOf, editOf, fileFormula, MIRRORS, runOnBook } from './case-book.js'
import { ensureExcel, excelVersion } from './excel.js'
import { localeMeta } from './locales.js'
import { compute } from '../../src/index.js'
import type { Step, WorkbookAnswer, WorkbookCase, WorkbookSet } from './workbook/cases.js'
import { HERE, ROOT } from './paths.js'

const SCRIPT = join(HERE, 'excel-workbook.applescript')

export interface StoredWorkbook {
  meta: Record<string, string | number>
  cases: WorkbookCase[]
  answers: Record<string, WorkbookAnswer>
}

/** Where the answers of a set are stored. */
export const workbookFile = (set: string) => join(ROOT, `test/excel/workbook/${set}.json`)

/** The cells of the rows (their first column) or the columns (their first row) a step hides or shows. */
export const hideRange = ({ rows, columns }: { rows?: string; columns?: string }): string => (rows !== undefined ? `A${rows.replace(':', ':A')}` : `${columns!.split(':')[0]}1:${columns!.split(':').at(-1)}1`)

/** A step as a line of the AppleScript's argument. */
function stepLine(s: Step): string {
  if (s === 'save') throw new Error('the workbook step save is Excel for Windows\'s alone')
  if (typeof s === 'string') return s
  if ('set' in s) {
    const { sheet, cell, value } = s.set
    return ['set', sheet, cell, ...editOf(value)].join('\t')
  }
  if ('formula' in s) {
    const { sheet, cell, formula, dynamic, array } = s.formula
    // typed as Excel 365 enters it, in the user's spelling: Excel for Mac leaves a cell empty when typed `_xlpm.` names reach it
    if (dynamic) return ['formula2', sheet, cell, `=${formula}`].join('\t')
    return array ? ['formulaarray', sheet, array, `=${fileFormula(formula)}`].join('\t') : ['formula', sheet, cell, `=${fileFormula(formula)}`].join('\t')
  }
  if ('hide' in s) {
    const { sheet, rows, hidden } = s.hide
    return [rows !== undefined ? 'hide' : 'hidecols', sheet, hideRange(s.hide), hidden ? 1 : 0].join('\t')
  }
  const { sheet, range, field, criterion } = s.filter
  return ['filter', sheet, range, field, criterion].join('\t')
}

/**
 * A book with a circular reference and without iterative calculation makes Excel for Mac show a warning window on
 * opening, which stops the run: such a case is refused before Excel sees it (the engine finds the cycle), and so is a
 * case marked `cycle`, which Excel for Windows alone runs.
 */
export function checkCycles(c: WorkbookCase) {
  if (c.cycle) throw new Error(`${c.id}: a circular reference without iterative calculation, run by Excel for Windows alone (Excel for Mac would show a warning window)`)
  if (c.file?.iterate) return
  const cycle = compute(takenBook(c.book), { display: false }).diagnostics.find((i) => i.kind === 'cell' && i.error === 'CYCLE')
  if (cycle?.kind === 'cell') throw new Error(`${c.id}: ${cycle.sheet}!${cycle.cell} is in a circular reference and the case has no iterative calculation (Excel would show a warning window)`)
}

/** One case in Excel: its book opened, its steps taken. */
async function runCase(c: WorkbookCase): Promise<WorkbookAnswer> {
  checkCycles(c)
  const steps = (c.steps ?? ['full', 'read']).map(stepLine).join('\n')
  const shows = (c.show ?? []).flatMap((g) => cellsOf(g.range).addrs.map((a) => `${g.sheet}!${a}`))
  const mirrors = (c.file?.mirrors ?? MIRRORS).join('\t')
  const out = await runOnBook(c.id, SCRIPT, c.book, c.read, [steps, shows.map((k) => k.replace('!', '\t')).join('\n'), mirrors], c.file)
  const raw = JSON.parse(out) as { settings: [boolean, number, number]; reads: [[unknown, unknown, unknown, unknown][], string[]][] }
  const [iteration, count, delta] = raw.settings
  return {
    settings: { iteration, count, delta },
    reads: raw.reads.map(([r, texts]) => ({ ...answerOf(c.read, r), ...(shows.length ? { shown: Object.fromEntries(shows.map((k, i) => [k, texts[i]])) } : {}) })),
  }
}

/** Runs the set's cases without a stored answer for the same case (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepWorkbook(set: string): Promise<void> {
  const cases = (await setModule<WorkbookSet>('workbook', set)).buildWorkbookCases()
  const target = workbookFile(set)
  const answers = keptSame({ cases }, loadStored<StoredWorkbook>(target), unpredicted)
  const t0 = Date.now()
  const save = () => writeAnswers(target, { meta: { version: excelVersion(), locale: localeMeta('en-US'), seconds: secondsSince(t0) }, cases, answers: inCaseOrder(cases, answers) })
  await ensureExcel('en-US')
  await runCases(cases.filter((c) => !answers[c.id]), 'FBOOK_ONLY', async (c) => void (answers[c.id] = await runCase(c)), (done) => {
    save()
    console.log(`${answeredLine(`workbook ${set}`, done, answers, cases)}, ${took(t0)} s`)
  })
}
