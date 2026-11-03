/**
 * Links to other workbooks in Excel for Windows: each case of `bench/coverage/links/<set>.ts` in the file the Mac step
 * sends, opened without updating its links (`Workbooks.Open(path, 0)`, as every book of this harness), the linked books
 * listed (`LinkSources`), the rectangles read through COM (`Value2`, and each cell's `Text`, `Formula`, `FormulaLocal`)
 * after opening and, after the case's operations (`Struct`), after a full calculation, and the book saved. Answers go to `test/excel/win/links/<set>.json` in
 * the Mac's shape, the saved files beside them.
 */
import { cellsOf, eachRectCell, refusedOf } from '../coverage/case-book.js'
import { linksSet } from '../coverage/links.js'
import type { LinksAnswer, LinksCase, LinksOp, LinksRead } from '../coverage/links/cases.js'
import { fileAnswers, secondsSince } from '../coverage/stored.js'
import { writeXlsx } from '../../src/xlsx.js'
import { fileJobs } from './books.js'
import { valOf, winFile, type WinCell } from './functions.js'
import { metaOf, type Step, type Tag, type WinEnv } from './remote.js'

/** One read: each rectangle's values, then its cells' shown text and formulas. */
const readSteps = (c: LinksCase): Step[] => c.read.flatMap((g) => [{ do: 'values', sheet: g.sheet, range: g.range }, { do: 'cells', sheet: g.sheet, cells: cellsOf(g.range).addrs }])

/**
 * An operation as a step of `xl.ps1`: a fill down or a copy on the case's sheet, rows inserted, the sheet renamed
 * (`Struct`); a formula typed (`Formula2`); the linked book written beside this one, as `writeXlsx` writes its values,
 * opened read-only or left on the disk, and closed and removed.
 */
async function opStep(o: LinksOp): Promise<Step> {
  if (o.op === 'enter') return { do: 'set', sheet: o.sheet, cell: o.cell, kind: 'f2', value: o.formula }
  if (o.op === 'source') {
    const sheets = Object.entries(o.sheets).map(([name, cells]) => ({ name, cells: Object.fromEntries(Object.entries(cells).map(([a, value]) => [a, { value }])) }))
    return { do: 'source', file: o.target, open: o.open, data: Buffer.from((await writeXlsx({ book: { sheets } })).data).toString('base64') }
  }
  if (o.op === 'dropSource') return { do: 'dropSource', file: o.target }
  if (o.op === 'fillDown') return { do: 'struct', op: 'fillDown', target: o.sheet, to: o.to }
  if (o.op === 'copyRange') return { do: 'struct', op: 'copyRange', sheet: o.sheet, range: o.range, target: o.sheet, to: o.to }
  if (o.op === 'insertRows') return { do: 'struct', op: 'insertRows', sheet: o.sheet, range: o.range }
  return { do: 'struct', op: 'renameSheet', sheet: o.sheet, target: o.to }
}

function readOf(c: LinksCase, got: unknown[]): LinksRead {
  const out: LinksRead = { values: {}, formulas: {}, local: {}, shown: {} }
  eachRectCell(c.read, (key, i, r, j) => {
    const vals = got[i * 2] as WinCell[][]
    const cell = (got[i * 2 + 1] as string[][])[r * vals[0].length + j]
    const v = valOf(vals[r][j])
    if (v) out.values[key] = v
    const [text, , , , , formula, local] = cell
    if (formula.startsWith('=')) {
      out.formulas[key] = formula
      out.local[key] = local
    }
    out.shown[key] = text
  })
  return out
}

export async function stepLinks(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { built, files, sent } = await linksSet(set)
  const store = fileAnswers<LinksCase, LinksAnswer>(winFile(`links/${set}.json`), built.cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ ...metaOf(e, tag), seconds: secondsSince(t0) }, built)
  // each case's operations as steps, and the books it put beside this one and left there, closed and removed after the save
  const ops = new Map<string, Step[]>()
  for (const c of store.todo) {
    const own = c.ops ?? []
    const left = own.flatMap((o) => (o.op === 'source' && !own.some((d) => d.op === 'dropSource' && d.target === o.target) ? [{ do: 'dropSource', file: o.target }] : []))
    ops.set(c.id, [...(await Promise.all(own.map(opStep))), ...left])
  }
  await fileJobs(
    `links ${set}`,
    `links-${set}`,
    store,
    files,
    (c) => {
      const steps = ops.get(c.id)!
      const m = (c.ops ?? []).length
      return [{ do: 'links' }, ...readSteps(c), ...steps.slice(0, m), { do: 'full' }, ...readSteps(c), { do: 'saveAs' }, ...steps.slice(m)]
    },
    (c, a) => {
      const [folder, sources] = a.answers[0] as [string, string[]]
      const n = c.read.length * 2
      const m = (c.ops ?? []).length
      const errors = refusedOf((a.answers.slice(1 + n, 1 + n + m) as ({ error?: string } | null)[]).map((x) => x?.error))
      return { sources, folder, ...(errors ? { errors } : {}), reads: [readOf(c, a.answers.slice(1, 1 + n)), readOf(c, a.answers.slice(2 + n + m, 2 + 2 * n + m))] }
    },
    save,
  )
}
