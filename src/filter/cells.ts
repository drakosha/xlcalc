/** The cells of a filter's or a sort's lines as a computed state holds them, for `keep.ts`: values, texts, formats, styles, runs and conditional looks. */
import type { CellValue, ConditionalAppearance } from '../contract.js'
import { windowLooks } from '../cond/looks.js'
import type { EngineState } from '../engine.js'
import { run } from '../exec/budget.js'
import { sheetLines } from '../layout/lines.js'
import { formatAddress, MAX_COLS } from '../model.js'
import { cellValue } from '../result.js'
import type { Rect } from '../structure/area.js'
import type { ColumnCells, KeepRules } from './keep.js'

/** What a filter compares by in the state's book. */
export const keepRules = (state: EngineState): KeepRules => ({ compareText: state.ctx.rules.compareText, epoch: state.ctx.rules.epoch, today: Math.floor(state.ctx.now) })

/**
 * The cells of sheet `s` inside `area`, by line: `at(i)` is the row and column of a line's `i`-th cell. The appearance
 * conditional formats give the area's cells is looked at once, when a color or an icon asks for it.
 */
export function areaCells(state: EngineState, s: number, area: Rect): (n: number, at: (i: number) => [number, number]) => ColumnCells {
  const ls = sheetLines(state.input.sheets[s], state)
  let looks: Map<number, ConditionalAppearance> | undefined
  return (n, at) => {
    const value = (i: number): CellValue => cellValue(state, s, ...at(i))
    const format = (i: number) => ls.format(...at(i))
    return {
      n,
      value,
      format,
      text: (i) => state.ctx.rules.display(value(i), format(i), Infinity).text,
      style: (i) => {
        const k = ls.look(...at(i)).style
        return k === undefined ? undefined : state.input.styles?.[k]
      },
      runs: (i) => state.input.sheets[s].cells[formatAddress(...at(i))]?.richText,
      look: (i) => {
        const [r, c] = at(i)
        return (looks ??= run(windowLooks(state, s, area))).get(r * MAX_COLS + c)
      },
    }
  }
}
