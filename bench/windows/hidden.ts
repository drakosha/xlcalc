/**
 * Hidden sheets in Excel for Windows: each case of `bench/coverage/hidden/<set>.ts` in the file the Mac step sends, the
 * visibility it sets through COM (`Worksheet.Visible`), its sheet operation, each sheet's visibility, the active sheet
 * and the read cells' values (`Value2`) read, and the book saved. Answers go to `test/excel/win/hidden/<set>.json` in the
 * Mac's shape, the saved files beside them; `test/hidden-sheets-win.test.ts` reads them.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import type { CellValue } from '../../src/contract.js'
import { hiddenFiles, type StoredHidden } from '../coverage/hidden.js'
import { keptAnswers } from '../coverage/stored.js'
import type { HiddenSet, Visible } from '../coverage/hidden/cases.js'
import { inCaseOrder } from '../coverage/stored.js'
import { opArgs } from '../coverage/struct.js'
import { caseJobs } from './books.js'
import { winFile } from './functions.js'
import { metaOf, type Step, type Tag, type WinEnv } from './remote.js'

/** COM's error codes of the values read (`0x800A0000` plus Excel's error number): none is expected, the read formulas catch them. */
const value = (v: unknown): CellValue => (v !== null && typeof v === 'object' ? { error: 'VALUE' } : (v as CellValue))

export async function stepHidden(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const built = ((await import(`../coverage/hidden/${set}.js`)) as HiddenSet).buildHiddenCases()
  const file = winFile(`hidden/${set}.json`)
  const dir = winFile(`hidden/${set}`)
  const prev = existsSync(file) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(file, 'utf8')) as StoredHidden) : null
  const { files, sent } = await hiddenFiles(built)
  const { answers, rejected } = keptAnswers(built.cases, sent, prev, (id) => existsSync(join(dir, `${id}.xlsx`)))
  mkdirSync(dir, { recursive: true })
  const t0 = Date.now()
  const save = () => {
    const refused = inCaseOrder(built.cases, rejected)
    mkdirSync(dirname(file), { recursive: true })
    writeFileSync(file, JSON.stringify({ meta: { ...metaOf(e, tag), seconds: Math.round((Date.now() - t0) / 100) / 10 }, ...built, sent: inCaseOrder(built.cases, sent), answers: inCaseOrder(built.cases, answers), ...(Object.keys(refused).length ? { rejected: refused } : {}) }, null, 1) + '\n')
  }
  const { sheet: readSheet, cells } = built.read
  const todo = built.cases.filter((c) => !answers[c.id] && !rejected[c.id])
  await caseJobs(
    `hidden-${set}`,
    todo,
    async (c, f) => {
      writeFileSync(f, files.get(c.id)!)
      const [sheet, range, target, to] = c.op ? opArgs({ id: c.id, book: c.book, op: c.op }) : []
      const steps: Step[] = [{ do: 'book' }, ...(c.make ?? []).map((m) => ({ do: 'visible', sheet: m.sheet, visible: m.visible })), ...(c.op ? [{ do: 'struct', op: c.op.op, sheet, range, target, to }] : [])]
      return { file: f, steps: [...steps, { do: 'calc' }, { do: 'book' }, ...cells.map((cell) => ({ do: 'values', sheet: readSheet, range: cell })), { do: 'saveAs' }] }
    },
    (c, a) => {
      const n = 1 + (c.make ?? []).length + (c.op ? 1 : 0)
      // the first step Excel refused, where the Mac stops
      const errors = (a.answers.slice(1, n) as ({ error?: string } | null)[]).map((x, i) => (x?.error ? `${i < (c.make ?? []).length ? 'make' : 'op'}: ${x.error}` : '')).filter(Boolean)
      const [opened] = a.answers[0] as [string, [string, number][]]
      const [active, sheets] = a.answers[n + 1] as [string, [string, number][]]
      const values = cells.map((_, i) => value((a.answers[n + 2 + i] as unknown[][])[0][0]))
      answers[c.id] = { ...(errors.length ? { error: errors[0] } : {}), opened, active, sheets: sheets.map(([name, v]) => [name, v as Visible]), values }
      if (a.saved) writeFileSync(join(dir, `${c.id}.xlsx`), a.saved)
      else {
        delete answers[c.id]
        rejected[c.id] = 'Excel did not save the book'
      }
    },
    (c, why) => {
      console.log(`${c.id}: Excel did not load the book (${why})`)
      rejected[c.id] = why
    },
    save,
  )
  save()
  console.log(`hidden ${set}: ${todo.length} cases run, ${Object.keys(answers).length} of ${built.cases.length} answered, ${Object.keys(rejected).length} not loaded`)
}
