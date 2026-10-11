/**
 * Data validation and outlines in Excel for Windows: each case of `bench/coverage/valid/<set>.ts` in the file the Mac
 * step sends (`writeXlsx`), what it makes through COM (`Validation.Add`, `Group`, `ShowDetail`, `Outline`), its operation,
 * every sheet's rules over the window (each cell selected, with `Validation.Value`), its outline and the book saved.
 * Answers go to `test/excel/win/valid/[<tag>/]<set>.json` in the Mac's shape (COM's numbers as the Mac's names), the saved
 * files beside them; `test/validation-win.test.ts` reads them.
 */
import { localeDir } from '../coverage/locales.js'
import { fileAnswers, secondsSince } from '../coverage/stored.js'
import { flag, formulaForExcel, ruleRange, validSet } from '../coverage/valid.js'
import { lineRange, WINDOW, type ExcelLine, type ExcelRule, type ValidAnswer, type ValidCase, type ValidMake } from '../coverage/valid/cases.js'
import { fileJobs, makeErrors, opSteps } from './books.js'
import { winFile } from './functions.js'
import { metaOf, type Step, type Tag, type WinEnv } from './remote.js'

const TYPES = ['validate input only', 'validate whole number', 'validate decimal', 'validate list', 'validated date', 'validate time', 'validate text length', 'validate custom']
const OPERATORS = ['', 'operator between', 'operator not between', 'operator equal', 'operator not equal', 'operator greater', 'operator less', 'operator greater equal', 'operator less equal']
const ALERTS = ['', 'valid alert stop', 'valid alert warning', 'valid alert information']
const TYPE_OF: Record<string, number> = { any: 0, whole: 1, decimal: 2, list: 3, date: 4, time: 5, textLength: 6, custom: 7 }
const OPERATOR_OF: Record<string, number> = { between: 1, notBetween: 2, equal: 3, notEqual: 4, greaterThan: 5, lessThan: 6, greaterThanOrEqual: 7, lessThanOrEqual: 8 }
const ALERT_OF: Record<string, number> = { stop: 1, warning: 2, information: 3 }

/** A step of `xl.ps1` that makes what the case makes. */
function makeStep(m: ValidMake): Step {
  if ('rule' in m) {
    const r = m.rule
    const { range, first } = ruleRange(m.range)
    const [f1, f2] = m.excel ?? [formulaForExcel(r.formula1), formulaForExcel(r.formula2)]
    return {
      do: 'vmake', kind: 'rule', sheet: m.sheet, range, first, type: TYPE_OF[r.type], alert: ALERT_OF[r.errorStyle ?? 'stop'], operator: OPERATOR_OF[r.operator ?? 'between'], f1, f2: f2 ?? '',
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
  const { built, files, sent } = await validSet(set)
  const { books, cases } = built
  const window = built.window ?? WINDOW
  const store = fileAnswers<ValidCase, ValidAnswer>(winFile(`valid/${localeDir(tag)}${set}.json`), cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ ...metaOf(e, tag), seconds: secondsSince(t0) }, { books, cases, ...(built.window ? { window: built.window } : {}) })
  await fileJobs(
    `valid ${set} (${tag})`,
    `valid-${set}`,
    store,
    files,
    (c) => [...(c.make ?? []).map(makeStep), ...opSteps(c), { do: 'valid', window }, { do: 'saveAs' }],
    (c, a) => {
      const { n, errors } = makeErrors(c, a.answers)
      const raw = a.answers[n] as RawSheet[]
      return {
        ...(errors.length ? { error: errors.join('; ') } : {}),
        sheets: raw.map(([name, rules, rows, columns, [sr, sc]]) => ({
          name,
          rules: rules.map((x): ExcelRule => [plain(x[0]), TYPES[x[1]], OPERATORS[x[2]] ?? '', x[3], x[4], x[5], x[6], x[7], x[8], ALERTS[x[9]] ?? '', x[10], x[11], x[12], x[13], x[14]]),
          rows,
          columns,
          summary: [sr === 0 ? 'summary above' : 'summary below', sc === -4131 ? 'summary on left' : 'summary on right'],
        })),
      }
    },
    save,
  )
}
