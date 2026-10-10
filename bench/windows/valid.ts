/**
 * Data validation and outlines in Excel for Windows: each case of `bench/coverage/valid/<set>.ts` in the file the Mac
 * step sends (`writeXlsx`), what it makes through COM (`Validation.Add`, `Group`, `ShowDetail`, `Outline`), its operation,
 * every sheet's rules over the window (each cell selected, with `Validation.Value`), its outline and the book saved.
 * Answers go to `test/excel/win/valid/[<tag>/]<set>.json` in the Mac's shape (COM's numbers as the Mac's names), the saved
 * files beside them; `test/validation-win.test.ts` reads them.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { writeXlsx } from '../../src/xlsx.js'
import { areasOf } from '../../src/validation/rules.js'
import { formatAddress } from '../../src/model.js'
import { localeDir } from '../coverage/locales.js'
import { inCaseOrder } from '../coverage/stored.js'
import { opArgs } from '../coverage/struct.js'
import { formulaForExcel, type StoredValid } from '../coverage/valid.js'
import { lineRange, WINDOW, type ExcelLine, type ExcelRule, type ValidAnswer, type ValidCase, type ValidMake, type ValidSet } from '../coverage/valid/cases.js'
import { partsHash } from '../coverage/xlsx/sent.js'
import { caseJobs } from './books.js'
import { winFile } from './functions.js'
import { metaOf, type Step, type Tag, type WinEnv } from './remote.js'

const TYPES = ['validate input only', 'validate whole number', 'validate decimal', 'validate list', 'validated date', 'validate time', 'validate text length', 'validate custom']
const OPERATORS = ['', 'operator between', 'operator not between', 'operator equal', 'operator not equal', 'operator greater', 'operator less', 'operator greater equal', 'operator less equal']
const ALERTS = ['', 'valid alert stop', 'valid alert warning', 'valid alert information']
const TYPE_OF: Record<string, number> = { any: 0, whole: 1, decimal: 2, list: 3, date: 4, time: 5, textLength: 6, custom: 7 }
const OPERATOR_OF: Record<string, number> = { between: 1, notBetween: 2, equal: 3, notEqual: 4, greaterThan: 5, lessThan: 6, greaterThanOrEqual: 7, lessThanOrEqual: 8 }
const ALERT_OF: Record<string, number> = { stop: 1, warning: 2, information: 3 }

const flag = (v: boolean | undefined) => (v === undefined ? '' : v ? '1' : '0')

/** A step of `xl.ps1` that makes what the case makes. */
function makeStep(m: ValidMake): Step {
  if ('rule' in m) {
    const r = m.rule
    const first = areasOf(m.range)![0]
    const [f1, f2] = m.excel ?? [formulaForExcel(r.formula1), formulaForExcel(r.formula2)]
    const range = areasOf(m.range)!.map((x) => `${formatAddress(x.r1, x.c1)}:${formatAddress(x.r2, x.c2)}`).join(',')
    return {
      do: 'vmake', kind: 'rule', sheet: m.sheet, range, first: formatAddress(first.r1, first.c1), type: TYPE_OF[r.type], alert: ALERT_OF[r.errorStyle ?? 'stop'], operator: OPERATOR_OF[r.operator ?? 'between'], f1, f2: f2 ?? '',
      ignore: flag(r.allowBlank), dropdown: flag(r.dropdown), showInput: flag(r.showPrompt), showError: flag(r.showError), inputTitle: r.promptTitle ?? '', inputMessage: r.prompt ?? '', errorTitle: r.errorTitle ?? '', errorMessage: r.error ?? '',
    }
  }
  if ('summary' in m) return { do: 'vmake', kind: 'summary', sheet: m.sheet, below: flag(m.summary.summaryBelow), right: flag(m.summary.summaryRight) }
  const text = 'group' in m ? m.group : m.collapse
  return { do: 'vmake', kind: 'group' in m ? 'group' : 'collapse', sheet: m.sheet, rows: /^\d/.test(text), range: lineRange(text) }
}

type RawRule = [string, number, number, string, string, boolean, boolean, boolean, boolean, number, string, string, string, string, boolean]
type RawSheet = [string, RawRule[], ExcelLine[], ExcelLine[], [number, number]]

const plain = (a: string) => a.replaceAll('$', '')

export async function stepValid(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const built = ((await import(`../coverage/valid/${set}.js`)) as ValidSet).buildValidCases()
  const { books, cases } = built
  const window = built.window ?? WINDOW
  const file = winFile(`valid/${localeDir(tag)}${set}.json`)
  const dir = file.slice(0, -5)
  const prev = existsSync(file) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(file, 'utf8')) as StoredValid) : null
  const files = new Map(await Promise.all(cases.map(async (c) => [c.id, (await writeXlsx(books[c.book])).data] as const)))
  const sent = Object.fromEntries(await Promise.all(cases.map(async (c) => [c.id, await partsHash(files.get(c.id)!)] as const)))
  const key = (c: ValidCase, hash: string | undefined) => JSON.stringify([c, hash])
  const same = new Map(prev?.cases.map((c) => [c.id, key(c, prev.sent[c.id])]) ?? [])
  const answers: Record<string, ValidAnswer> = {}
  const rejected: Record<string, string> = {}
  for (const c of cases)
    if (same.get(c.id) === key(c, sent[c.id])) {
      if (prev!.answers[c.id] && existsSync(join(dir, `${c.id}.xlsx`))) answers[c.id] = prev!.answers[c.id]
      else if (prev!.rejected?.[c.id]) rejected[c.id] = prev!.rejected[c.id]
    }
  mkdirSync(dir, { recursive: true })
  const t0 = Date.now()
  const save = () => {
    const refused = inCaseOrder(cases, rejected)
    mkdirSync(dirname(file), { recursive: true })
    writeFileSync(file, JSON.stringify({ meta: { ...metaOf(e, tag), seconds: Math.round((Date.now() - t0) / 100) / 10 }, books, cases, ...(built.window ? { window: built.window } : {}), sent: inCaseOrder(cases, sent), answers: inCaseOrder(cases, answers), ...(Object.keys(refused).length ? { rejected: refused } : {}) }, null, 1) + '\n')
  }
  const todo = cases.filter((c) => !answers[c.id] && !rejected[c.id])
  await caseJobs(
    `valid-${set}`,
    todo,
    async (c, f) => {
      writeFileSync(f, files.get(c.id)!)
      const [sheet, range, target, to] = c.op ? opArgs({ id: c.id, book: c.book, op: c.op }) : []
      const op: Step[] = c.op ? [{ do: 'struct', op: c.op.op, sheet, range, target, to }] : []
      return { file: f, steps: [...(c.make ?? []).map(makeStep), ...op, { do: 'valid', window }, { do: 'saveAs' }] }
    },
    (c, a) => {
      const n = (c.make ?? []).length + (c.op ? 1 : 0)
      const errors = (a.answers.slice(0, n) as ({ error?: string } | null)[]).flatMap((x, i) => (x?.error ? [`${i < (c.make ?? []).length ? 'make' : 'op'}: ${x.error}`] : []))
      const raw = a.answers[n] as RawSheet[]
      answers[c.id] = {
        ...(errors.length ? { error: errors.join('; ') } : {}),
        sheets: raw.map(([name, rules, rows, columns, [sr, sc]]) => ({
          name,
          rules: rules.map((x): ExcelRule => [plain(x[0]), TYPES[x[1]], OPERATORS[x[2]] ?? '', x[3], x[4], x[5], x[6], x[7], x[8], ALERTS[x[9]] ?? '', x[10], x[11], x[12], x[13], x[14]]),
          rows,
          columns,
          summary: [sr === 0 ? 'summary above' : 'summary below', sc === -4131 ? 'summary on left' : 'summary on right'],
        })),
      }
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
  console.log(`valid ${set} (${tag}): ${todo.length} cases run, ${Object.keys(answers).length} of ${cases.length} answered, ${Object.keys(rejected).length} not loaded`)
}
