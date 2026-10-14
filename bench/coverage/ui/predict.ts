/**
 * The engine's answers to a ui set as its predictions, written before Excel runs the set: `npx tsx
 * bench/coverage/ui/predict.ts <set>` writes `ui/<set>.predicted.json`, each case's answer in each of the package's
 * locales, by the session's clock on the day of the prediction. A case of `bars`: each cell's `editText` in a window of
 * `getViewport` (empty for a cell the window does not list). A paste or a fill: `previewCellTexts` and `setCellTexts` of
 * the texts the action puts into cells (`entriesOf`), and after it the cells read as the book holds them, their values
 * and the widths of their columns. An edit in place (F2): `pasteText` of its text, which reads it with no percent
 * editor, and the cells after it. The ribbon's box: `numberFormatCategory` of each cell's format; an item of its list
 * applied: the item of `numberFormatChoices`, its name, the format it applies and the box's category after. Point mode:
 * the sheet component's model (`point-model.ts`), the formula bar after each step and the cells after.
 */
import { fileURLToPath } from 'node:url'
import { setModule, writeAnswers } from '../stored.js'
import { createSession, type CellInput, type CellValidationAlert, type CellValue } from '../../../src/index.js'
import { TAGS } from '../locales.js'
import { entriesOf, readCells, readColumns, ribbonOf, uiBook, windowOver, type UiCase, type UiSet } from './cases.js'
import { pointPrediction } from './point-model.js'

const set = process.argv[2]
const { buildUiCases, POINT_BY } = await setModule<UiSet & { POINT_BY?: 'refAt' }>('ui', set)
const now = new Date().toISOString().slice(0, 10)
const clock = { now: `${now}T12:00:00Z`, seed: 1 }

/** An alert as the prediction keeps it: the cell, the style and the message. */
const alert = (v: CellValidationAlert) => ({ cell: v.cell, severity: v.severity, message: v.message })

function predict(c: UiCase, tag: string) {
  if (c.act.do === 'ribbon' || c.act.do === 'choose') return { ribbon: ribbonOf(c, tag) }
  // Point mode: the sheet component's model, the formula bar after each step and the cells after
  if (c.act.do === 'point') return pointPrediction(uiBook(c, tag), c.act.at, c.act.steps, tag, readCells(c), POINT_BY)
  const s = createSession(uiBook(c, tag), { clock })
  try {
    const cells = readCells(c)
    if (c.act.do === 'bars') {
      const v = s.getViewport('S', windowOver(cells), { editText: true })
      return { editText: Object.fromEntries(cells.map((a) => [a, v.cells[a]?.editText ?? ''])) }
    }
    const after = () => {
      const book = s.getBook().sheets[0]
      return Object.fromEntries(cells.map((a) => [a, { input: book.cells[a] ?? null, value: s.getCell('S', a)?.value ?? null }])) as Record<string, { input: CellInput | null; value: CellValue }>
    }
    if (c.act.do === 'edit') {
      const r = s.pasteText('S', c.act.at, c.act.text)
      return { set: r.ok ? {} : { refused: { cell: r.cell, code: r.code } }, after: after() }
    }
    const entries = entriesOf(c.act)
    const p = s.previewCellTexts(entries)
    const preview = p.ok ? { cells: Object.fromEntries(p.cells.map((x) => [x.cell, { input: x.input, value: x.value }])), ...(p.widths ? { widths: p.widths } : {}), ...(p.validations ? { validations: p.validations.map(alert) } : {}) } : { refused: { cell: p.cell, code: p.code, at: p.at } }
    const r = s.setCellTexts(entries)
    const book = s.getBook().sheets[0]
    const widths = Object.fromEntries(readColumns(c).map((k) => [k, book.columns?.[k]?.width ?? null]))
    return { preview, set: r.ok ? { ...(r.widths ? { widths: r.widths } : {}), ...(r.validations ? { validations: r.validations.map(alert) } : {}) } : { refused: { cell: r.cell, code: r.code, at: r.at } }, after: after(), widths }
  } finally {
    s.close()
  }
}

const locales = Object.fromEntries(TAGS.map((tag) => [tag, Object.fromEntries(buildUiCases().map((c) => [c.id, predict(c, tag)]))]))
writeAnswers(fileURLToPath(new URL(`./${set}.predicted.json`, import.meta.url)), { now, locales })
