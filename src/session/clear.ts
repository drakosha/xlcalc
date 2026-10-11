/**
 * Cells cleared (`Session.clearCells`), as Excel's ClearContents, ClearFormats and Clear (the set `lines/clear`): the plan
 * of the cells' new inputs, the lines' looks and the sheets' rules. Contents: values, formulas and hyperlinks; notes and
 * meta stay. Formats: each cell's style and number format, so that it shows the book's plain look whatever its row or
 * column has (a cell whose lines show a look gets the plain style or General of its own), whole rows' and columns' own
 * looks, and the cells of conditional formats over them. All: both, the cells' notes and meta, and the cells of rules
 * of data validation over them (`docs/edits.md`).
 */
import type { ClearKind, RangeRef } from '../contract.js'
import type { EngineState } from '../engine.js'
import { checkOneOf, checkSheet, object, onlyKeys, array } from '../input.js'
import { clearConditional } from '../cond/structure.js'
import { sheetNames } from '../structure/rewrite.js'
import type { Rect } from '../structure/area.js'
import { sheetKey } from '../syntax.js'
import { clearValidations } from '../validation/structure.js'
import { checkRange, linesOf, type Look } from './patch.js'
import { Draft, lookAt, type Plan } from './styles.js'

const LOOKS: Record<ClearKind, Look> = {
  contents: { content: true, link: null },
  formats: { style: { reset: true }, format: { reset: true } },
  all: { content: true, style: { reset: true }, format: { reset: true }, meta: null, note: null, link: null },
}

/** The plan of clearing the rectangles `ranges` of `what`. */
export function clearPlan(s: EngineState, ranges: RangeRef[], what: ClearKind): Plan {
  checkOneOf(what, 'what', Object.keys(LOOKS))
  const d = new Draft(s)
  const holes = new Map<number, Rect[]>()
  array(ranges, 'ranges', 'an array of ranges such as [{ sheet: "Sheet1", range: "B2:D9" }]').forEach((v, k) => {
    const p = `ranges[${k}]`
    const o = object(v, p, 'a range such as { sheet: "Sheet1", range: "B2:D9" }')
    onlyKeys(o, p, ['sheet', 'range'])
    const i = checkSheet(o.sheet, (name) => s.sheetIdx.get(sheetKey(name)), `${p}.sheet`)
    const rect = checkRange(o.range, `${p}.range`)
    const lines = linesOf(rect)
    lookAt(d, { i, rect, ...(lines ? { lines } : {}) }, LOOKS[what])
    holes.set(i, [...(holes.get(i) ?? []), rect])
  })
  if (what !== 'contents') {
    const names = sheetNames(s.input)
    for (const [i, xs] of holes) {
      const sh = d.sheet(i)
      const conditionalFormats = clearConditional(sh.conditionalFormats, xs, sh.name, names)
      const validations = what === 'all' ? clearValidations(sh.validations, xs, sh.name, names) : sh.validations
      if (conditionalFormats === sh.conditionalFormats && validations === sh.validations) continue
      const next = { ...sh, conditionalFormats, validations }
      if (!conditionalFormats) delete next.conditionalFormats
      if (!validations) delete next.validations
      d.sheets.set(i, next)
    }
  }
  return d.plan()
}
