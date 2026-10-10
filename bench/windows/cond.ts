/**
 * Conditional formats in Excel for Windows: each case of `bench/coverage/cond/<set>.ts` in the file the Mac step sends
 * (`writeXlsx`), the rules it makes through COM (`FormatConditions.Add` and its kin), its operation, the book computed,
 * the read cells' display format (`DisplayFormat`: Range.Text, fill and font colours, bold, italic), the rules each
 * sheet lists, and the book saved. Answers go to `test/excel/win/cond/[<tag>/]<set>.json` in the Mac's shape, the saved
 * files beside them; `test/cond-win.test.ts` reads them.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { condSet, type StoredCond } from '../coverage/cond.js'
import { fromBgr, type CondAnswer, type CondCase, type Look } from '../coverage/cond/cases.js'
import { localeDir } from '../coverage/locales.js'
import { inCaseOrder } from '../coverage/stored.js'
import { opArgs } from '../coverage/struct.js'
import { caseJobs } from './books.js'
import { winFile } from './functions.js'
import { metaOf, type Step, type Tag, type WinEnv } from './remote.js'

type RawLook = [address: string, text: string, fill: number, colorIndex: number, font: number, bold: boolean, italic: boolean]
type RawRules = [sheet: string, rules: [number, string, string, number, boolean][]][]

/** xlNone: no fill. */
const NO_FILL = -4142

export async function stepCond(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { built, files, sent } = await condSet(set)
  const { books, cases } = built
  const base = `cond/${localeDir(tag)}`
  const file = winFile(`${base}${set}.json`)
  const dir = winFile(`${base}${set}`)
  const prev = existsSync(file) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(file, 'utf8')) as StoredCond) : null
  const key = (c: CondCase, hash: string | undefined) => JSON.stringify([c, hash])
  const same = new Map(prev?.cases.map((c) => [c.id, key(c, prev.sent[c.id])]) ?? [])
  const answers: Record<string, CondAnswer> = {}
  const rejected: Record<string, string> = {}
  for (const c of cases)
    if (same.get(c.id) === key(c, sent[c.id])) {
      if (prev!.answers[c.id] && existsSync(join(dir, `${c.id}.xlsx`))) answers[c.id] = prev!.answers[c.id]
      else if (prev!.rejected?.[c.id]) rejected[c.id] = prev!.rejected[c.id]
    }
  mkdirSync(dir, { recursive: true })
  const t0 = Date.now()
  const day = new Date().toLocaleDateString('sv-SE')
  const save = () => {
    const refused = inCaseOrder(cases, rejected)
    mkdirSync(dirname(file), { recursive: true })
    writeFileSync(file, JSON.stringify({ meta: { ...metaOf(e, tag), day, seconds: Math.round((Date.now() - t0) / 100) / 10 }, books, cases, ...(built.locale ? { locale: built.locale } : {}), sent: inCaseOrder(cases, sent), answers: inCaseOrder(cases, answers), ...(Object.keys(refused).length ? { rejected: refused } : {}) }, null, 1) + '\n')
  }
  const todo = cases.filter((c) => !answers[c.id] && !rejected[c.id])
  await caseJobs(
    `cond-${set}`,
    todo,
    async (c, f) => {
      writeFileSync(f, files.get(c.id)!)
      const [sheet, range, target, to] = c.op ? opArgs({ id: c.id, book: c.book, op: c.op }) : []
      const make: Step[] = (c.make ?? []).map((m) => ({ do: 'cond', kind: 'make', ...m }))
      const op: Step[] = c.op ? [{ do: 'struct', op: c.op.op, sheet, range, target, to }] : []
      return { file: f, steps: [...make, ...op, { do: 'full' }, { do: 'cond', kind: 'looks', read: c.read }, { do: 'cond', kind: 'rules' }, { do: 'saveAs' }] }
    },
    (c, a) => {
      const n = (c.make ?? []).length + (c.op ? 1 : 0)
      const errors = (a.answers.slice(0, n) as ({ error?: string } | null)[]).flatMap((x, i) => (x?.error ? [`${i < (c.make ?? []).length ? 'make' : 'op'}: ${x.error}`] : []))
      const looks: CondAnswer['looks'] = {}
      for (const [sheet, cells] of a.answers[n + 1] as [string, RawLook[]][])
        looks[sheet] = Object.fromEntries(cells.map(([addr, text, fill, ci, font, bold, italic]): [string, Look] => [addr, { text, fill: ci === NO_FILL ? null : fromBgr(fill), font: fromBgr(font), bold, italic }]))
      const rules = Object.fromEntries((a.answers[n + 2] as RawRules).map(([sheet, list]) => [sheet, list]))
      answers[c.id] = { ...(errors.length ? { error: errors.join('; ') } : {}), looks, rules }
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
  console.log(`cond ${set}: ${todo.length} cases run, ${Object.keys(answers).length} of ${cases.length} answered, ${Object.keys(rejected).length} not loaded`)
}
