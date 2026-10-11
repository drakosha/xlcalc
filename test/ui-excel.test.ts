/**
 * The engine against what Excel did with a person's actions in its window, en-US and ru-RU (`test/excel/ui/` from Excel
 * for Mac, `test/excel/win/ui/` from Excel for Windows, the step `ui` of `bench/coverage` and `bench/windows`):
 * - a cell of an array, selected and in editing (`arrays-held`): `ViewportCell.editText` is what F2 opens it with on the
 *   Mac, and what the object model gives as `Formula2Local` on both; the formula bar of a selected cell shows the same
 *   text, but an old-kind array's formula in braces in every cell of its area and a spilled cell's anchor's formula;
 * - text pasted from the clipboard (`paste-held`, both Excels) and entered into a selection with Control+Return
 *   (`fill-held`, the Mac's keyboard): `setCellTexts` of the same texts, a host taking back a call with a stop alert as
 *   Excel takes no such value. What each cell holds after (its value or formula and its format), the width the columns
 *   took, and the cells whose entry brought a window (a refused formula, a value its rule does not take).
 * Known differences are in `test/ui-known.ts`; the predictions committed before Excel ran each set
 * (`bench/coverage/ui/<set>.predicted.json`) are kept with how many cases Excel confirmed.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredUi } from '../bench/coverage/ui.js'
import { entriesOf, readCells, readColumns, uiBook, windowOver, type UiAnswer, type UiCase, type UiSet } from '../bench/coverage/ui/cases.js'
import { formulaBar } from '../bench/coverage/entry/cases.js'
import { held } from '../bench/coverage/entry/compare.js'
import { localeDir, TAGS } from '../bench/coverage/locales.js'
import { createSession, type CellInput } from '../src/index.js'
import { UI_KNOWN, UI_REASONS } from './ui-known.js'
import './locales.js'

interface Run {
  excel: 'Mac' | 'Windows'
  tag: string
  set: string
  stored: StoredUi
}

const RUNS: Run[] = (['Mac', 'Windows'] as const).flatMap((excel) =>
  TAGS.flatMap((tag) => {
    const dir = new URL(`./excel/${excel === 'Windows' ? 'win/' : ''}ui/${localeDir(tag)}`, import.meta.url)
    if (!existsSync(dir)) return []
    return readdirSync(dir)
      .filter((f) => f.endsWith('.json'))
      .map((f) => ({ excel, tag, set: f.slice(0, -5), stored: JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredUi }))
  }),
)

/** What a paste or a fill left, as compared: each read cell's value or formula and format, the width typing gave each column (null where it kept its own), the cells whose entry brought a window. */
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
    // a window comes up for the entry of the active cell, the action's first
    windows: a.dialog ? [entriesOf(c.act)[0].cell] : [],
  }
}

/** `setCellTexts` of the action's texts on the day of the run; a stop alert's call taken back, as the host does. */
function engineOutcome(c: UiCase, tag: string, date: string): Outcome {
  const s = createSession(uiBook(c, tag), { clock: { now: `${date}T12:00:00Z`, seed: 1 } })
  try {
    const entries = entriesOf(c.act)
    const r = s.setCellTexts(entries)
    const undone = r.ok && r.validations?.some((v) => v.severity === 'stop') === true
    if (undone) s.undo()
    const cells = s.getBook().sheets[0].cells
    return {
      cells: Object.fromEntries(readCells(c).map((x) => [x, cellKey(cells[x], tag)])),
      widths: Object.fromEntries(readColumns(c).map((k) => [k, (r.ok && !undone ? r.widths?.find((w) => w.column === k)?.width : undefined) ?? null])),
      windows: r.ok ? (r.validations ?? []).map((v) => v.cell) : [entries[0].cell],
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
      const cases = (await current(set)).filter((c) => excel === 'Mac' || c.act.do !== 'fill')
      expect(stored.cases).toEqual(cases)
      expect(cases.filter((c) => !stored.answers[c.id])).toEqual([])
    })

    for (const c of stored.cases)
      it(c.id, () => {
        const a = stored.answers[c.id]
        const reasons = UI_KNOWN.get(`${tag}:${set}:${c.id}`)
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
}

interface Predicted {
  editText?: Record<string, string>
  after?: Record<string, { input: CellInput | null }>
  set?: { refused?: unknown; widths?: { column: string; width: number }[]; validations?: { cell: string }[] }
}

describe('ui sets (Excel for Mac)', () => {
  const mac = RUNS.filter((r) => r.excel === 'Mac')
  it('every known difference is a stored case', () => {
    const stored = new Set(mac.flatMap((r) => r.stored.cases.map((c) => `${r.tag}:${r.set}:${c.id}`)))
    expect([...UI_KNOWN.keys()].filter((k) => !stored.has(k))).toEqual([])
  })
  it.each(mac.map((r) => [`${r.tag}:${r.set}`, r] as const))('%s: the predictions written before Excel ran it', (key, run) => {
    const predicted = JSON.parse(readFileSync(new URL(`../bench/coverage/ui/${run.set}.predicted.json`, import.meta.url), 'utf8')).locales[run.tag] as Record<string, Predicted>
    const confirmed = run.stored.cases.filter((c) => {
      const p = predicted[c.id]
      const a = run.stored.answers[c.id]
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
