/**
 * Formats of rows and columns in Excel for Windows: each case of `bench/coverage/lines/<set>.ts` in the file the Mac step
 * sends (`writeXlsx`), its steps through COM (`NumberFormat`, `Font.Bold`, `Interior.Color` of a range, `Value2` of a
 * cell), the book computed, the read cells' `Range.Text`, number format and display format (`DisplayFormat`: fill and font
 * colours, bold, italic), and the book saved. Answers go to `test/excel/win/lines/<set>.json` in the Mac's shape, the
 * saved files beside them; `test/lines-excel.test.ts` reads them.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fromBgr } from '../coverage/cond/cases.js'
import { linesSet, type StoredLines } from '../coverage/lines.js'
import type { LineLook, LinesAnswer, LineStep } from '../coverage/lines/cases.js'
import { inCaseOrder, keptAnswers } from '../coverage/stored.js'
import { caseJobs } from './books.js'
import { winFile } from './functions.js'
import { metaOf, type Step, type Tag, type WinEnv } from './remote.js'

type RawLook = [address: string, text: string, fill: number, colorIndex: number, font: number, bold: boolean, italic: boolean, format: string]

/** xlNone: no fill. */
const NO_FILL = -4142

/** `#RRGGBB` as COM's BGR. */
const bgr = (c: string) => parseInt(c.slice(5, 7) + c.slice(3, 5) + c.slice(1, 3), 16)

/** A step through COM: the range's formats, or a value typed into a cell. */
const comStep = (s: LineStep): Step =>
  s.value !== undefined
    ? { do: 'set', sheet: s.sheet, cell: s.range, kind: typeof s.value === 'number' ? 'n' : 's', value: String(s.value) }
    : { do: 'format', sheet: s.sheet, range: s.range, ...(s.format !== undefined ? { numberFormat: s.format } : {}), ...(s.bold ? { bold: true } : {}), ...(s.fill ? { fill: bgr(s.fill) } : {}) }

export async function stepLines(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { built, files, sent } = await linesSet(set)
  const file = winFile(`lines/${set}.json`)
  const dir = winFile(`lines/${set}`)
  const prev = existsSync(file) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(file, 'utf8')) as StoredLines) : null
  const { answers, rejected } = keptAnswers(built.cases, sent, prev, (id) => existsSync(join(dir, `${id}.xlsx`)))
  mkdirSync(dir, { recursive: true })
  const t0 = Date.now()
  const save = () => {
    const refused = inCaseOrder(built.cases, rejected)
    mkdirSync(dirname(file), { recursive: true })
    writeFileSync(file, JSON.stringify({ meta: { ...metaOf(e, tag), seconds: Math.round((Date.now() - t0) / 100) / 10 }, ...built, sent: inCaseOrder(built.cases, sent), answers: inCaseOrder(built.cases, answers), ...(Object.keys(refused).length ? { rejected: refused } : {}) }, null, 1) + '\n')
  }
  const todo = built.cases.filter((c) => !answers[c.id] && !rejected[c.id])
  await caseJobs(
    `lines-${set}`,
    todo,
    async (c, f) => {
      writeFileSync(f, files.get(c.id)!)
      return { file: f, steps: [...(c.steps ?? []).map(comStep), { do: 'full' }, { do: 'cond', kind: 'looks', read: c.read, codes: true }, { do: 'saveAs' }] }
    },
    (c, a) => {
      const n = (c.steps ?? []).length
      const errors = (a.answers.slice(0, n) as ({ error?: string } | null)[]).flatMap((x, i) => (x?.error ? [`${c.steps![i].range}: ${x.error}`] : []))
      const looks: LinesAnswer['looks'] = {}
      for (const [sheet, cells] of a.answers[n + 1] as [string, RawLook[]][])
        looks[sheet] = Object.fromEntries(cells.map(([addr, text, fill, ci, font, bold, italic, format]): [string, LineLook] => [addr, { text, format, fill: ci === NO_FILL ? null : fromBgr(fill), font: fromBgr(font), bold, italic }]))
      answers[c.id] = { ...(errors.length ? { error: errors.join('; ') } : {}), looks }
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
  console.log(`lines ${set}: ${todo.length} cases run, ${Object.keys(answers).length} of ${built.cases.length} answered, ${Object.keys(rejected).length} not loaded`)
}
