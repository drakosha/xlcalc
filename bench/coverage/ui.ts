/**
 * What Excel for Mac does with what a person does in its window: each case of `ui/<set>.ts` is written by `writeXlsx` into
 * a book of its own, and with Excel in front (`excel-ui.applescript`, the keyboard's handlers in `excel-keys.applescript`)
 * Excel reads the formula bar of cells selected and in editing, takes text pasted from the clipboard, a text entered
 * into a selection with Control+Return, or reads the ribbon's Number Format box and applies items of its list (through
 * Accessibility, nothing typed); then the cells and their columns are read and the book is saved, its file read
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
import { escapeAttr } from '../../src/xlsx/xml/text.js'
import { declineRepair, ensureExcel, excelVersion, runWatched, settleRejected } from './excel.js'
import { heldOf } from './entry.js'
import { savedColumns } from './fit.js'
import { localeDir, localeMeta } from './locales.js'
import { answeredLine, inCaseOrder, keptSame, loadStored, runCases, secondsSince, setModule, writeAnswers } from './stored.js'
import { clipboardText, readCells, readColumns, selectedOf, uiBook, type PointStep, type UiAnswer, type UiCase, type UiCell, type UiSet } from './ui/cases.js'
import { rawFile } from './xlsx/sent.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-ui.applescript')
const KEYS = join(HERE, 'excel-keys.applescript')
const INPUT_SOURCE = join(HERE, 'input-source.swift')
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

/** The formula bar of a cell of `bars` or `edit`: selected, in editing, the element in focus, the windows Excel showed; an edit's bar once its keys were typed. */
export type RawBar = [[string, string], [string, string], string, string, [string, string]?]

/** What the ribbon's Number Format box showed for each cell an action selected, and the item applied to it (empty for none). */
export type RawRibbon = [string, string][]

/** The cells `names` of the sheet `S` as Excel read them (`cells`, in their order) and saved them (`saved`). */
export async function savedCells(names: string[], cells: RawCell[], saved: Uint8Array): Promise<Record<string, UiCell>> {
  const read = await readXlsx(saved)
  const held = read.book.sheets.find((s) => s.name === 'S')!.cells
  const cached = read.cached.sheets.find((s) => s.name === 'S')!.cells
  const out: Record<string, UiCell> = {}
  names.forEach((a, i) => {
    const [text, code, local, prefix, value2, formula, formulaLocal, formula2, formula2Local, hasArray, formulaArray, hasSpill, spillParent] = cells[i]
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
    }
  })
  return out
}

/** The answer of a case from what Excel read of its cells and columns, the windows it showed and the file it saved. */
export async function caseAnswer(c: UiCase, dialog: string, cells: RawCell[], columns: [number, number, boolean][], saved: Uint8Array, bars?: RawBar[], ribbon?: RawRibbon, point?: [string, string][]): Promise<UiAnswer> {
  const cols = await savedColumns(saved)
  const out = await savedCells(readCells(c), cells, saved)
  readCells(c).forEach((a, i) => {
    // a case of bars, of an edit or of the ribbon reads its cells first
    const b = bars?.[i]
    const r = ribbon?.[i]
    if (b) out[a].bar = { selected: b[0], edit: b[1], focused: b[2], ...(b[3] ? { dialog: b[3] } : {}), ...(b[4] ? { typed: b[4] } : {}) }
    if (r) out[a].ribbon = { box: r[1], ...(r[0] ? { item: r[0] } : {}) }
  })
  const letters = readColumns(c)
  return {
    ...(dialog ? { dialog } : {}),
    ...(point ? { point: point.map(([bar, said]) => ({ bar, ...(said ? { dialog: said } : {}) })) } : {}),
    cells: out,
    columns: Object.fromEntries(letters.map((k, i) => [k, { columnWidth: columns[i][0], points: columns[i][1], standard: columns[i][2], saved: cols.get(k) ?? null }])),
  }
}

/** The action as the script takes it; a paste with its clipboard's text, a choice with its cells and items apart. */
function actJson(c: UiCase): string {
  const act = c.act
  if (act.do === 'paste') return JSON.stringify({ do: 'paste', at: act.at, text: clipboardText(act) })
  return JSON.stringify(act.do === 'choose' ? { do: 'choose', cells: selectedOf(act), items: act.picks.map(([, k]) => k) } : act)
}

/** A marker the writer stores a code of the file's own with (an id from 164), taken out of the file after. */
const OWN = '"\u00a7own"'

/** A case's file: `writeXlsx`'s, with the codes of its `own` cells as codes of the file's own. */
export async function caseFile(c: UiCase, label: string): Promise<Uint8Array> {
  const own = new Set(c.own)
  const cells = Object.fromEntries(Object.entries(c.cells).map(([a, x]) => [a, own.has(a) ? { ...x, numberFormat: `${x.numberFormat}${OWN}` } : x]))
  const { data, issues } = await writeXlsx({ book: uiBook({ ...c, cells }) }, { values: 'omit' })
  if (issues.length) throw new Error(`ui ${label}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
  const codes = new Set([...own].map((a) => String(c.cells[a].numberFormat)))
  return rawFile(data, [...codes].map((code) => ({ part: 'xl/styles.xml', from: `formatCode="${escapeAttr(code + OWN)}"`, to: `formatCode="${escapeAttr(code)}"` })))
}

/** A case of `point` whose formula bar does not end with a text it typed, right after typing it, is no answer: the keyboard typed something else. */
function checkTyped(steps: PointStep[], bars: [string, string][], label: string): void {
  steps.forEach((s, i) => {
    if (!('type' in s)) return
    const shown = / [A-Z]+,\d+ (.*)$/s.exec(bars[i]?.[0] ?? '')?.[1] ?? ''
    if (!shown.endsWith(s.type)) throw new Error(`ui ${label}: typed ${JSON.stringify(s.type)}, the formula bar shows ${JSON.stringify(shown)}: the keyboard's layout is not ABC`)
  })
}

/** One case in Excel, its action taken (`none` reads the book alone). A book Excel repairs stops the run; a window Excel shows for the action is answered by the script. */
export async function runCase(c: UiCase, label: string, act = actJson(c)): Promise<UiAnswer> {
  if (declineRepair()) console.log('closed a leftover book repair window')
  // the keyboard types a case of `point` in the ABC layout, taken again before each: another layout typed `=ф2#` for `=E2#`
  if (c.act.do === 'point') execFileSync('swift', [INPUT_SOURCE, 'com.apple.keylayout.ABC'])
  return runWatched<UiAnswer>({
    script: SCRIPT,
    label: `ui ${label}`,
    folder: 'fui',
    book: async (file) => writeFileSync(file, await caseFile(c, label)),
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
      const raw = JSON.parse(out) as { dialog: string; bars: RawBar[]; ribbon: RawRibbon; point: [string, string][]; cells: RawCell[]; columns: [number, number, boolean][] }
      if (c.act.do === 'point') checkTyped(c.act.steps, raw.point, label)
      return caseAnswer(c, raw.dialog, raw.cells, raw.columns, new Uint8Array(readFileSync(join(dir, 'saved.xlsx'))), raw.bars.length ? raw.bars : undefined, raw.ribbon.length ? raw.ribbon : undefined, raw.point.length ? raw.point : undefined)
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
