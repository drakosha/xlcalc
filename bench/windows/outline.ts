/**
 * Excel's outline commands in Excel for Windows (`bench/coverage/outline/`): each case's book written by `writeXlsx` and
 * opened; Excel reads the window's lines and saves a copy, then takes each command through COM (`Group`, `Ungroup`,
 * `ShowDetail`, `Hidden`, `Outline.ShowLevels`) and reads and saves again; at the end it saves the book. The answers go
 * to `test/excel/win/outline/[<tag>/]<set>.json`, the books saved beside them; `test/outline-win.test.ts` reads them.
 */
import { writeXlsx } from '../../src/xlsx.js'
import { comLines, sentBook, windowOf, type ComLine, type OutlineAnswer, type OutlineCase, type OutlineCommand, type OutlineSet, type OutlineStep } from '../coverage/outline/cases.js'
import { localeDir } from '../coverage/locales.js'
import { fileAnswers, secondsSince, setModule } from '../coverage/stored.js'
import { caseFiles } from '../coverage/xlsx/sent.js'
import { fileJobs } from './books.js'
import { winFile } from './functions.js'
import { metaOf, type Step, type Tag, type WinEnv } from './remote.js'

/** A command as a step of `xl.ps1`. */
function commandStep(c: OutlineCommand): Step {
  if ('levels' in c) return { do: 'ocmd', kind: 'levels', rows: c.levels[0], columns: c.levels[1] }
  const [kind, lines] = Object.entries(c)[0] as [string, string]
  return { do: 'ocmd', kind, lines: comLines(lines) }
}

export async function stepOutline(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { BOOKS, buildOutlineCases } = await setModule<OutlineSet>('outline', set)
  const cases = buildOutlineCases()
  const { files, sent } = await caseFiles(cases, async (c) => (await writeXlsx({ book: sentBook(BOOKS[c.book]) }, { values: 'omit' })).data)
  const store = fileAnswers<OutlineCase, OutlineAnswer>(winFile(`outline/${localeDir(tag)}${set}.json`), cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ ...metaOf(e, tag), date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0) }, { books: BOOKS, cases })
  await fileJobs(
    `outline ${set} (${tag})`,
    `outline-${set}`,
    store,
    files,
    (c) => {
      const w = windowOf(c.window)
      const read: Step = { do: 'oread', rows: w.rows.map((i) => i + 1), columns: w.columns.map((i) => i + 1) }
      return [read, ...c.commands.flatMap((x) => [commandStep(x), read]), { do: 'saveAs' }]
    },
    (c, a) => {
      const read = (x: unknown, error?: unknown): OutlineStep => {
        const [rows, columns, saved] = x as [ComLine[], ComLine[], string[]]
        const err = error && typeof error === 'object' && 'error' in error ? String(error.error) : undefined
        return { ...(err ? { error: err } : {}), rows, columns, saved }
      }
      return { steps: [read(a.answers[0]), ...c.commands.map((_, i) => read(a.answers[2 + 2 * i], a.answers[1 + 2 * i]))] }
    },
    save,
  )
}
