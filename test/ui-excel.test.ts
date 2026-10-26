/**
 * The engine against what Excel did with a person's actions in its window, en-US and ru-RU (`test/excel/ui/` from Excel
 * for Mac, `test/excel/win/ui/` from Excel for Windows, the step `ui` of `bench/coverage` and `bench/windows`):
 * - a cell of an array, selected and in editing (`arrays-held`): `ViewportCell.editText` is what F2 opens it with on the
 *   Mac, and what the object model gives as `Formula2Local` on both; the formula bar of a selected cell shows the same
 *   text, but an old-kind array's formula in braces in every cell of its area and a spilled cell's anchor's formula;
 * - text pasted from the clipboard (`paste-held`, both Excels): `pasteText` of the clipboard's text; a text entered
 *   into a selection with Control+Return (`fill-held`, the Mac's keyboard): `fillCellText` from the selection's first
 *   cell, a host taking back a call with a stop alert as Excel takes no such value; a cell edited in place after F2
 *   (`edit-held`, the Mac's keyboard): `setCellText` with `inPlace`; an item of a dropdown list picked (`pick-held`):
 *   `setCells` of the item's `input` from `getValidation`. What each cell holds after (its
 *   value or formula and its format), the width the columns took, and whether the action brought a window (a refused
 *   formula, a value the active cell's rule does not take), which Excel shows for the active cell;
 * - the ribbon's Number Format box (`ribbon-*`, the Mac's screen): what it shows for a cell selected, against
 *   `numberFormatCategory` of the cell's format; an item of its list applied (`choose`), against `numberFormatChoices`:
 *   the item's name, the format the saved file holds, what the box shows after.
 * Known differences are in `test/ui-known.ts`; the predictions committed before Excel ran each set
 * (`bench/coverage/ui/<set>.predicted.json`) are kept with how many cases Excel confirmed.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredUi } from '../bench/coverage/ui.js'
import { clipboardText, entriesOf, macOnly, pressInEngine, readCells, readColumns, ribbonOf, uiBook, windowOver, type RibbonCell, type UiAnswer, type UiCase, type UiSet } from '../bench/coverage/ui/cases.js'
import { formulaBar } from '../bench/coverage/entry/cases.js'
import { formatKey, held } from '../bench/coverage/entry/compare.js'
import { numberFormatOf } from '../bench/coverage/stored.js'
import { createSession, LOCALES, type CellInput } from '../src/index.js'
import { UI_KNOWN, UI_REASONS } from './ui-known.js'
import { localeSets } from './stored-sets.js'
import './locales.js'

interface Run {
  excel: 'Mac' | 'Windows'
  tag: string
  set: string
  stored: StoredUi
}

const RUNS: Run[] = (['Mac', 'Windows'] as const).flatMap(
  (excel) =>
    localeSets<StoredUi>(`${excel === 'Windows' ? 'win/' : ''}ui`).map((r) => ({ excel, ...r })),
  // Point mode's sets are compared with `refAt` in `point-excel.test.ts`
).filter((r) => r.stored.cases.every((c) => c.act.do !== 'point'))

/** What a paste or a fill left, as compared: each read cell's value or formula and format, the width typing gave each column (null where it kept its own), the cell a window came up for. */
interface Outcome {
  cells: Record<string, string>
  widths: Record<string, number | null>
  windows: string[]
}

const cellKey = (c: CellInput | null | undefined, tag: string) => JSON.stringify(held(c ?? null, tag))

function excelOutcome(c: UiCase, a: UiAnswer, tag: string): Outcome {
  return {
    cells: Object.fromEntries(readCells(c).map((x) => [x, cellKey(a.cells[x].cell, tag)])),
    widths: Object.fromEntries(readColumns(c).map((k) => [k, a.columns[k].saved?.bestFit ? a.columns[k].saved.width : null])),
    // a window comes up for the entry of the active cell, the action's first; an edit's is read with its bar
    windows: (a.dialog ?? (c.act.do === 'edit' ? a.cells[c.act.at].bar?.dialog : undefined)) ? [entriesOf(c.act)[0]?.cell ?? readCells(c)[0]] : [],
  }
}

/** The action on the day of the run: a paste by `pasteText`, a fill by `fillCellText`, an edit by `setCellText` in place, a pick by its item's `input` from `getValidation`, the ribbon's buttons by `pressInEngine`; a stop alert's call taken back, as the host does. */
function engineOutcome(c: UiCase, tag: string, date: string): Outcome {
  const s = createSession(uiBook(c, tag), { clock: { now: `${date}T12:00:00Z`, seed: 1 } })
  try {
    const act = c.act
    if (act.do === 'pick') {
      // the item's input from getValidation, written by setCells: Excel checks no rule and widens no column
      const list = s.getValidation('S', act.at)?.list
      if (!Array.isArray(list)) throw new Error(`${c.id}: no list at ${act.at}`)
      s.setCells([{ sheet: 'S', cell: act.at, input: list[act.item - 1].input }])
      const cells = s.getBook().sheets[0].cells
      return { cells: Object.fromEntries(readCells(c).map((x) => [x, cellKey(cells[x], tag)])), widths: Object.fromEntries(readColumns(c).map((k) => [k, null])), windows: [] }
    }
    if (act.do === 'press') {
      pressInEngine(s, c, tag)
      const cells = s.getBook().sheets[0].cells
      return { cells: Object.fromEntries(readCells(c).map((x) => [x, cellKey(cells[x], tag)])), widths: Object.fromEntries(readColumns(c).map((k) => [k, null])), windows: [] }
    }
    if (act.do !== 'paste' && act.do !== 'fill' && act.do !== 'edit') throw new Error(`${c.id} is no paste, fill, edit or pick`)
    const r = act.do === 'paste' ? s.pasteText('S', act.at, clipboardText(act)) : act.do === 'fill' ? s.fillCellText('S', act.range, act.text) : s.setCellText('S', act.at, act.text, { inPlace: true })
    const alert = r.ok && 'validation' in r ? r.validation : undefined
    const undone = alert?.severity === 'stop'
    if (undone) s.undo()
    const width = r.ok && !undone && 'width' in r && r.width !== undefined ? [{ column: readColumns(c)[0], width: r.width }] : undefined
    const widths = r.ok && !undone && 'widths' in r ? r.widths : width
    const cells = s.getBook().sheets[0].cells
    return {
      cells: Object.fromEntries(readCells(c).map((x) => [x, cellKey(cells[x], tag)])),
      widths: Object.fromEntries(readColumns(c).map((k) => [k, widths?.find((w) => w.column === k)?.width ?? null])),
      windows: !r.ok || alert ? [entriesOf(act)[0].cell] : [],
    }
  } finally {
    s.close()
  }
}

/** The engine's window over a case's cells: each one's `editText`, and the anchor of a spilled cell. */
function engineCells(c: UiCase, tag: string, date: string): Record<string, { editText: string; spillOf?: string }> {
  const s = createSession(uiBook(c, tag), { clock: { now: `${date}T12:00:00Z`, seed: 1 } })
  try {
    const cells = readCells(c)
    const v = s.getViewport('S', windowOver(cells), { editText: true })
    return Object.fromEntries(cells.map((a) => [a, { editText: v.cells[a]?.editText ?? '', ...(v.cells[a]?.spillOf ? { spillOf: v.cells[a].spillOf } : {}) }]))
  } finally {
    s.close()
  }
}

/** What the ribbon showed in each cell of a case in Excel: the box; for an item applied, its name (the first of its two lines), the format the saved file holds and the box after. */
function excelRibbon(c: UiCase, a: UiAnswer): Record<string, RibbonCell> {
  return Object.fromEntries(
    readCells(c).map((x) => {
      const cell = a.cells[x]
      const { box, item } = cell.ribbon!
      return [x, c.act.do === 'choose' ? { name: item!.split('\n')[0], format: numberFormatOf(cell.cell) ?? 0, box } : box]
    }),
  )
}

/** A cell's ribbon as compared: an item's format by what it shows in the locale (`44` and the code Excel writes for it alike). */
const ribbonKey = (r: RibbonCell, tag: string) => JSON.stringify(typeof r === 'string' ? r : { ...r, format: formatKey(r.format, tag) })

/** The cells whose ribbon differs from Excel's, `numberFormatCategory` and `numberFormatChoices` against what Excel showed and applied: each cell with a line that says how. */
function ribbonOff(c: UiCase, a: UiAnswer, tag: string, got: Record<string, RibbonCell> = ribbonOf(c, tag)): [string, string][] {
  const want = excelRibbon(c, a)
  return Object.keys(want).flatMap((x) => (ribbonKey(got[x], tag) === ribbonKey(want[x], tag) ? [] : [[x, `${x} ${JSON.stringify(c.cells[x]?.numberFormat ?? 0)}: Excel ${JSON.stringify(want[x])}, the functions ${JSON.stringify(got[x])}`] as [string, string]]))
}

const ribbonCase = (c: UiCase) => c.act.do === 'ribbon' || c.act.do === 'choose'

/** The cells of a ribbon case where Excel for Mac in a locale other than en-US shows Custom and the function another category, all else the same (`macSyntax` in `test/ui-known.ts`). */
function macSyntax(run: Run, c: UiCase, a: UiAnswer): Set<string> {
  const custom = LOCALES.get(run.tag)?.regional?.formats?.categories.custom
  if (run.excel !== 'Mac' || !custom) return new Set()
  const got = ribbonOf(c, run.tag)
  const want = excelRibbon(c, a)
  const box = (r: RibbonCell) => (typeof r === 'string' ? r : r.box)
  const withBox = (r: RibbonCell, b: string): RibbonCell => (typeof r === 'string' ? b : { ...r, box: b })
  return new Set(Object.keys(want).filter((x) => box(want[x]) === custom && box(got[x]) !== custom && ribbonKey(withBox(want[x], box(got[x])), run.tag) === ribbonKey(got[x], run.tag)))
}

/** What Excel opened each cell of a case of `bars` with: on the Mac its formula bar once F2 began the editing, on Windows `Formula2Local`. */
const opened = (run: Run, a: UiAnswer, cell: string) => (run.excel === 'Mac' ? formulaBar(a.cells[cell].bar!.edit[0]) : a.cells[cell].formula2Local)

const generators = new Map<string, Promise<UiCase[]>>()
const current = (set: string) => {
  let g = generators.get(set)
  if (!g) generators.set(set, (g = (import(`../bench/coverage/ui/${set}.ts`) as Promise<UiSet>).then((m) => m.buildUiCases())))
  return g
}

for (const run of RUNS) {
  const { excel, tag, set, stored } = run
  const date = String(stored.meta.date)
  describe(`ui ${tag} ${set} (Excel for ${excel})`, () => {
    it('the stored cases are the generator\'s current ones, each answered', async () => {
      const cases = (await current(set)).filter((c) => excel === 'Mac' || !macOnly(c.act))
      expect(stored.cases).toEqual(cases)
      expect(cases.filter((c) => !stored.answers[c.id])).toEqual([])
    })

    for (const c of stored.cases)
      it(c.id, () => {
        const a = stored.answers[c.id]
        const reasons = UI_KNOWN.get(`${tag}:${set}:${c.id}`)
        if (ribbonCase(c)) {
          // a known difference names its cell: the others agree, and each known one still differs
          const off = ribbonOff(c, a, tag)
          const mac = macSyntax(run, c, a)
          const known = (x: string) => UI_KNOWN.get(`${tag}:${set}:${c.id}:${x}`) ?? (mac.has(x) ? (['macSyntax'] as const) : undefined)
          const fixed = readCells(c).filter((x) => known(x) && !off.some(([y]) => y === x)).map((x) => `${x}: ${known(x)!.map((r) => UI_REASONS[r]).join('; ')}`)
          expect({ unknown: off.filter(([x]) => !known(x)).map(([, line]) => line), fixed }).toEqual({ unknown: [], fixed: [] })
          return
        }
        if (c.act.do === 'bars') {
          const engine = engineCells(c, tag, date)
          const off = c.act.cells.filter((x) => engine[x].editText !== opened(run, a, x)).map((x) => `${x}: Excel ${JSON.stringify(opened(run, a, x))}, editText ${JSON.stringify(engine[x].editText)}`)
          if (reasons) expect(off.length, reasons.map((r) => UI_REASONS[r]).join('; ')).toBeGreaterThan(0)
          else expect(off).toEqual([])
          return
        }
        const [got, want] = [engineOutcome(c, tag, date), excelOutcome(c, a, tag)]
        if (reasons) expect(got, reasons.map((r) => UI_REASONS[r]).join('; ')).not.toEqual(want)
        else expect(got).toEqual(want)
      })

    // a rule without a message of its own: the alert's text is Excel's (fill-held's stop and information rules)
    const alerts = stored.cases.filter((c) => c.act.do === 'fill' && stored.answers[c.id].dialog && c.validations?.some((v) => v.error === undefined))
    if (excel === 'Mac' && alerts.length)
      it("a rule's alert without a message of its own says what Excel's says", () => {
        for (const c of alerts) {
          const act = c.act as Extract<UiCase['act'], { do: 'fill' }>
          const r = createSession(uiBook(c, tag), { clock: { now: `${date}T12:00:00Z`, seed: 1 } }).fillCellText('S', act.range, act.text)
          expect(stored.answers[c.id].dialog).toContain(r.ok && 'validation' in r ? r.validation!.message : 'no alert')
        }
      })

    const edits = stored.cases.filter((c) => c.act.do === 'edit')
    if (edits.length)
      it('an edit in place: F2 opens what editText gives, and the bar holds the case\'s text once its keys are typed', () => {
        const off = edits.flatMap((c) => {
          const act = c.act as Extract<UiCase['act'], { do: 'edit' }>
          const bar = stored.answers[c.id].cells[act.at].bar!
          const engine = engineCells(c, tag, date)[act.at].editText
          return [...(formulaBar(bar.edit[0]) === engine ? [] : [`${c.id} opened ${bar.edit[0]}, editText ${engine}`]), ...(formulaBar(bar.typed![0]) === act.text ? [] : [`${c.id} typed ${bar.typed![0]}, text ${act.text}`])]
        })
        expect(off).toEqual([])
      })

    const bars = stored.cases.filter((c) => c.act.do === 'bars')
    if (excel === 'Mac' && bars.length)
      it('the formula bar of a selected cell: what F2 opens, an old-kind array in braces, a spilled cell its anchor\'s formula; the object model\'s Formula2Local is what F2 opens', () => {
        const off = bars.flatMap((c) => {
          const a = stored.answers[c.id]
          const engine = engineCells(c, tag, date)
          return readCells(c).flatMap((x) => {
            const cell = a.cells[x]
            const edit = opened(run, a, x)
            const anchor = engine[x].spillOf
            const shown = cell.hasArray ? `{${edit}}` : anchor ? opened(run, a, anchor) : edit
            return [...(formulaBar(cell.bar!.selected[0]) === shown ? [] : [`${x} selected: ${cell.bar!.selected[0]}, expected ${shown}`]), ...(cell.formula2Local === edit ? [] : [`${x} Formula2Local ${cell.formula2Local}, F2 ${edit}`])]
          })
        })
        expect(off).toEqual([])
      })
  })
}

/** How many cases of each set's prediction Excel for Mac confirmed, by `<tag>:<set>`. */
const CONFIRMED: Record<string, number> = {
  'en-US:arrays-held': 3,
  'ru-RU:arrays-held': 3,
  'en-US:paste-held': 11,
  'ru-RU:paste-held': 12,
  'en-US:fill-held': 8,
  'ru-RU:fill-held': 9,
  'en-US:ribbon': 1,
  'ru-RU:ribbon': 0,
  'en-US:edit-held': 14,
  'en-US:pick-held': 5,
  'ru-RU:pick-held': 5,
  'ru-RU:edit-held': 13,
}

interface Predicted {
  editText?: Record<string, string>
  ribbon?: Record<string, RibbonCell>
  after?: Record<string, { input: CellInput | null }>
  set?: { refused?: unknown; widths?: { column: string; width: number }[]; validations?: { cell: string }[] }
}

describe('ui sets (Excel for Mac)', () => {
  const mac = RUNS.filter((r) => r.excel === 'Mac')
  it('every known difference is a stored case, or a cell read in one', () => {
    const stored = new Set(mac.flatMap((r) => r.stored.cases.flatMap((c) => [`${r.tag}:${r.set}:${c.id}`, ...readCells(c).map((x) => `${r.tag}:${r.set}:${c.id}:${x}`)])))
    expect([...UI_KNOWN.keys()].filter((k) => !stored.has(k))).toEqual([])
  })
  it.each(mac.map((r) => [`${r.tag}:${r.set}`, r] as const))('%s: the predictions written before Excel ran it', (key, run) => {
    const predicted = JSON.parse(readFileSync(new URL(`../bench/coverage/ui/${run.set}.predicted.json`, import.meta.url), 'utf8')).locales[run.tag] as Record<string, Predicted>
    const confirmed = run.stored.cases.filter((c) => {
      const p = predicted[c.id]
      const a = run.stored.answers[c.id]
      if (ribbonCase(c)) return ribbonOff(c, a, run.tag, p.ribbon).length === 0
      if (c.act.do === 'bars') return c.act.cells.every((x) => p.editText![x] === opened(run, a, x))
      const got: Outcome = {
        cells: Object.fromEntries(readCells(c).map((x) => [x, cellKey(p.after![x].input, run.tag)])),
        widths: Object.fromEntries(readColumns(c).map((k) => [k, p.set!.widths?.find((w) => w.column === k)?.width ?? null])),
        windows: p.set!.refused ? [entriesOf(c.act)[0].cell] : (p.set!.validations ?? []).map((v) => v.cell),
      }
      return JSON.stringify(got) === JSON.stringify(excelOutcome(c, a, run.tag))
    })
    expect(confirmed.length).toBe(CONFIRMED[key])
  })
})
