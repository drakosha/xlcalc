/**
 * A filter-acts set's prediction, written before Excel runs the set: `npx tsx bench/coverage/filter-acts/predict.ts <set>`
 * writes `filter-acts/<set>.predicted.json`. The range an AutoFilter acts on is the engine's (`getAutoFilterRange`, the
 * documented rule: nothing in it looks at what a row holds). The rest is the rule of the docs carried over to what the
 * engine has no step for: criteria and a sort act on that range and leave it saved as the AutoFilter's; criteria hide
 * every row of it that does not meet them and show every one that does, whatever hid it before; a row a sort takes along
 * moves with its cells, so a total on top of the amounts goes first and its reference falls off the sheet (#REF!); a label
 * alone sorts last, as an empty cell does; removing a filter with criteria shows every row of the range it acts on and
 * nothing else, one without criteria shows nothing; a new range is the one given, Data > Filter takes a selection down
 * over the data around it and a cell its current region; `Range.AutoFilter` with no arguments on a sheet with an
 * AutoFilter removes it, with a field it moves it to its range. A table's filter stays over its data, its total row out
 * of it, and it marks nothing as the sheet's filter mode.
 */
import { createSession } from '../../../src/index.js'
import { setModule, writePrediction } from '../stored.js'
import { actBook, ACT_ROWS, amountOf, TOTAL, type ActCase, type ActSet, type PredictedAct, type PredictedState } from './cases.js'

/** The range an AutoFilter on `own` acts on, as the engine reads it over the case's cells. */
function acted(c: ActCase, own: string): string {
  const book = actBook({ ...c, filter: own, table: undefined, hidden: undefined, zero: undefined, fileCriteria: undefined })
  const s = createSession(book)
  const got = s.getAutoFilterRange('S')!
  s.close()
  return got
}

const rowsOf = (range: string) => {
  const [, top, bottom] = /^[A-Z]+(\d+):[A-Z]+(\d+)$/.exec(range)!
  return { top: Number(top), bottom: Number(bottom) }
}

const textOf = (c: ActCase, row: number, col: string) => {
  const cell = c.cells[`${col}${row}`]
  return cell && 'value' in cell ? String(cell.value) : ''
}

function predict(c: ActCase): PredictedAct {
  let own = c.filter ?? c.table ?? null
  const hidden = new Set([...(c.hidden ?? []), ...(c.zero ?? [])])
  let criteria: Record<string, string> = c.fileCriteria ? { '0': 'Север' } : {}
  let filtered = !!c.fileCriteria
  const label = Object.keys(c.cells).find((k) => textOf(c, Number(k.slice(1)), k[0]) === TOTAL)
  let totalRow = label ? Number(label.slice(1)) : null
  let totalFormula = label ? (c.cells[`C${totalRow}`] as { formula?: string } | undefined)?.formula ?? '' : ''
  let broken = false
  let tableRef: string | null = null
  const range = () => (!own ? '' : c.table ? own : acted(c, own))
  if (c.keptOff) {
    const { top, bottom } = rowsOf(`A${c.keptOff.replace(':', ':A')}`)
    for (let r = top; r <= bottom; r++) hidden.delete(r)
    own = null
    criteria = {}
    filtered = false
  }
  const shown = (r: number) => !hidden.has(r)
  const totalText = () => {
    if (totalRow === null || !totalFormula) return ''
    if (broken) return '#REF!'
    if (totalFormula === 'SUM(C2:C21)') return '46030'
    let sum = 0
    for (let r = 2; r <= 21; r++) if (shown(r)) sum += amountOf(r)
    return String(sum)
  }
  const state = (): PredictedState => ({
    range: range(),
    filterMode: filtered && !c.table,
    hidden: [...hidden].filter((r) => r <= ACT_ROWS).sort((a, b) => a - b),
    heights: Object.fromEntries((c.watch ?? []).map((r) => [String(r), hidden.has(r) ? 0 : 'standard'])),
    total: totalRow === null ? null : { row: totalRow, text: totalText() },
  })
  const removed = () => {
    if (filtered) {
      const { top, bottom } = rowsOf(range())
      for (let r = top + 1; r <= bottom; r++) hidden.delete(r)
    }
    own = null
    criteria = {}
    filtered = false
  }
  const opened = state()
  const after = c.steps.map((s) => {
    if ('criteria' in s) {
      const acts = range()
      const { top, bottom } = rowsOf(acts)
      for (let r = top + 1; r <= bottom; r++) {
        if (textOf(c, r, s.column) === s.criteria) hidden.delete(r)
        else hidden.add(r)
      }
      if (!c.table) own = acts
      criteria = { [String(s.column.charCodeAt(0) - 65)]: s.criteria }
      filtered = true
    } else if ('sort' in s) {
      const acts = range()
      const { bottom } = rowsOf(acts)
      if (!c.table) own = acts
      if (totalRow !== null && totalRow <= bottom && totalFormula) {
        totalRow = 2
        broken = true
      }
    } else if ('off' in s) removed()
    else if ('over' in s) {
      if (s.via === 'plain' && own) removed()
      else {
        if (own) removed()
        const cell = !s.over.includes(':')
        // a cell takes its current region, a selection Data > Filter takes down over the data around it
        own = cell ? acted(c, 'A1:C2') : s.via === 'button' ? acted(c, s.over) : s.over
      }
    } else if ('totals' in s) {
      totalFormula = 'SUBTOTAL(109,Sales[Сумма])'
      tableRef = 'A1:C22'
    } else if (s.set === TOTAL) totalRow = 22
    return state()
  })
  const end = state()
  return { opens: true, opened, after, saved: { ref: own, tableRef, criteria, filterMode: end.filterMode, hidden: end.hidden } }
}

const set = process.argv[2]
const { buildActCases, predictAct } = await setModule<ActSet>('filter-acts', set)
writePrediction(new URL(`./${set}.ts`, import.meta.url).href, set, { cases: Object.fromEntries(buildActCases().map((c) => [c.id, predictAct ? predictAct(c, (own) => acted(c, own)) : predict(c)])) })
