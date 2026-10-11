/**
 * What Excel for Mac does with what a person does in its window: each case of `ui/<set>.ts` is written by `writeXlsx` into
 * a book of its own, and with Excel in front (`excel-ui.applescript`, the keyboard's handlers in `excel-keys.applescript`)
 * Excel reads the formula bar of cells selected and in editing, takes text pasted from the clipboard, or a text entered
 * into a selection with Control+Return; then the cells and their columns are read and the book is saved, its file read
 * back by `readXlsx` for what each cell holds and its `<col>` elements for the widths Excel stores. `npx tsx
 * bench/coverage/run.ts ui <set> [ru|<tag>]` stores the answers in `test/excel/ui/[<tag>/]<set>.json` with the day of the
 * run; `test/ui-excel.test.ts` compares the engine with them. The clipboard's text is put back after the set. `FUI_ONLY=id,id`
 * runs only those cases, `FCOV_FRESH=1` every case again.
 */
import { execFileSync } from 'node:child_process'
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { readXlsx, writeXlsx } from '../../src/xlsx.js'
import { declineRepair, ensureExcel, excelVersion, runWatched, settleRejected } from './excel.js'
import { heldOf } from './entry.js'
import { savedColumns } from './fit.js'
import { localeDir, localeMeta } from './locales.js'
import { answeredLine, inCaseOrder, keptSame, loadStored, runCases, secondsSince, setModule, writeAnswers } from './stored.js'
import { clipboardText, readCells, readColumns, uiBook, type UiAnswer, type UiCase, type UiCell, type UiSet } from './ui/cases.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-ui.applescript')
const KEYS = join(HERE, 'excel-keys.applescript')
/** pbpaste and pbcopy read and write text in the locale's encoding. */
const UTF8 = { ...process.env, LANG: 'en_US.UTF-8' }

export interface StoredUi {
  meta: Record<string, string | number>
  cases: UiCase[]
  answers: Record<string, UiAnswer>
}

/** The stored answers of a set: `test/excel/[win/]ui/<set>.json`, in a directory of its tag for a locale other than en-US. */
export const uiFile = (set: string, tag = 'en-US', win = false) => join(ROOT, 'test/excel', win ? 'win' : '', 'ui', localeDir(tag), `${set}.json`)

/** A cell as the object models read it: what it shows, its format codes, prefix, number, formulas, its old-kind array; Windows adds its spill. */
export type RawCell = [string, string, string, string, unknown, string, string, string, string, boolean, string, boolean?, string?]

/** The answer of a case from what Excel read of its cells and columns, the windows it showed and the file it saved. */
export async function caseAnswer(c: UiCase, dialog: string, cells: RawCell[], columns: [number, number, boolean][], saved: Uint8Array, bars?: [[string, string], [string, string], string, string][]): Promise<UiAnswer> {
  const read = await readXlsx(saved)
  const held = read.book.sheets.find((s) => s.name === 'S')!.cells
  const cached = read.cached.sheets.find((s) => s.name === 'S')!.cells
  const cols = await savedColumns(saved)
  const names = readCells(c)
  const out: Record<string, UiCell> = {}
  names.forEach((a, i) => {
    const [text, code, local, prefix, value2, formula, formulaLocal, formula2, formula2Local, hasArray, formulaArray, hasSpill, spillParent] = cells[i]
    // a case of bars reads its cells first
    const b = bars?.[i]
    out[a] = {
      cell: heldOf(held[a]),
      ...(a in cached ? { computed: cached[a].value } : {}),
      text,
      code,
      local,
      ...(prefix ? { prefix } : {}),
      ...(typeof value2 === 'number' ? { value2 } : {}),
      formula,
      formulaLocal,
      formula2,
      formula2Local,
      hasArray,
      ...(hasArray ? { formulaArray } : {}),
      ...(hasSpill === undefined ? {} : { hasSpill, ...(hasSpill ? { spillParent } : {}) }),
      ...(b ? { bar: { selected: b[0], edit: b[1], focused: b[2], ...(b[3] ? { dialog: b[3] } : {}) } } : {}),
    }
  })
  const letters = readColumns(c)
  return {
    ...(dialog ? { dialog } : {}),
    cells: out,
    columns: Object.fromEntries(letters.map((k, i) => [k, { columnWidth: columns[i][0], points: columns[i][1], standard: columns[i][2], saved: cols.get(k) ?? null }])),
  }
}

/** The action as the script takes it; a paste with its clipboard's text. */
const actJson = (c: UiCase) => JSON.stringify(c.act.do === 'paste' ? { do: 'paste', at: c.act.at, text: clipboardText(c.act) } : c.act)

/** One case in Excel, its action taken (`none` reads the book alone). A book Excel repairs stops the run; a window Excel shows for the action is answered by the script. */
export async function runCase(c: UiCase, label: string, act = actJson(c)): Promise<UiAnswer> {
  if (declineRepair()) console.log('closed a leftover book repair window')
  return runWatched<UiAnswer>({
    script: SCRIPT,
    label: `ui ${label}`,
    folder: 'fui',
    book: async (file) => {
      const { data, issues } = await writeXlsx({ book: uiBook(c) }, { values: 'omit' })
      if (issues.length) throw new Error(`ui ${label}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
      writeFileSync(file, data)
    },
    args: (dir) => [join(dir, 'saved.xlsx'), 'saved.xlsx', act, JSON.stringify(readCells(c)), JSON.stringify(readColumns(c)), KEYS],
    timeout: 170_000,
    poll: 1000,
    after: 'case',
    // the script answers the windows its action brings
    fatal: false,
    repair: async (stop) => {
      await stop()
      await settleRejected()
      throw new Error(`ui ${label}: Excel would repair the book`)
    },
    done: (out, dir) => {
      const raw = JSON.parse(out) as { dialog: string; bars: [[string, string], [string, string], string, string][]; cells: RawCell[]; columns: [number, number, boolean][] }
      return caseAnswer(c, raw.dialog, raw.cells, raw.columns, new Uint8Array(readFileSync(join(dir, 'saved.xlsx'))), raw.bars.length ? raw.bars : undefined)
    },
  })
}

/** Runs the set's cases without a stored answer for the same case (FCOV_FRESH=1 runs all) and stores the answers; the clipboard's text is put back after. */
export async function stepUi(set: string, tag: string): Promise<void> {
  const cases = (await setModule<UiSet>('ui', set)).buildUiCases()
  const target = uiFile(set, tag)
  const prev = loadStored<StoredUi>(target)
  const answers = keptSame({ cases }, prev)
  const todo = cases.filter((c) => !answers[c.id])
  const t0 = Date.now()
  // the time of the runs before it too
  const save = () => writeAnswers(target, { meta: { version: excelVersion(), locale: localeMeta(tag), date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0 - Number(prev?.meta.seconds ?? 0) * 1000) }, cases, answers: inCaseOrder(cases, answers) })
  await ensureExcel(tag)
  // the clipboard's text, put back after a set that pastes
  const pastes = todo.some((c) => c.act.do === 'paste')
  const clipboard = pastes ? execFileSync('pbpaste', { encoding: 'utf8', env: UTF8 }) : ''
  try {
    await runCases(todo, 'FUI_ONLY', async (c) => {
      answers[c.id] = await runCase(c, `${set} ${tag} ${c.id}`)
      save()
    }, (n) => {
      save()
      console.log(answeredLine(`ui ${set} ${tag}`, n, answers, cases))
    })
  } finally {
    if (pastes) execFileSync('pbcopy', { input: clipboard, env: UTF8 })
  }
}
