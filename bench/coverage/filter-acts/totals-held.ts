/**
 * Held out after `totals`: where a total row stops being part of what an AutoFilter on A1:C21 filters and sorts. SUBTOTAL
 * in other columns and in several, over another range, inside another function and with a function inside it, AGGREGATE,
 * in one of two rows under it, with its label and without, each of its codes, a formula that refers to such a cell. Each
 * under criteria, and all but the codes on a copy sorted, through the object model. And Range(A1).AutoFilter against
 * Data > Filter on A1 over a row of SUBTOTAL, in both orders in one Excel and in one book, as they saved different ranges.
 */
import type { CellInput } from '../../../src/contract.js'
import { regionOf, sales, TOTAL, type ActCase, type ActStep, type PredictedAct } from './cases.js'

const SUB = 'SUBTOTAL(109,C2:C21)'
const COUNT = (c: string) => `SUBTOTAL(103,${c}2:${c}21)`
const f = (formula: string): CellInput => ({ formula })
const LABEL: CellInput = { value: TOTAL }
const away = () => Object.fromEntries([1, 2, 3, 4, 5].map((r) => [`E${r}`, { value: r }]))

/** The rows of a case: what they add to the book, and the rows the model leaves out of what the filter does. */
const ROWS: [string, Record<string, CellInput>, number[]][] = [
  ['no-label', { C22: f(SUB) }, [22]],
  ['first-column', { A22: f(COUNT('A')) }, [22]],
  ['first-and-last', { A22: f(COUNT('A')), C22: f(SUB) }, [22]],
  ['several', { A22: LABEL, B22: f(COUNT('B')), C22: f(SUB) }, [22]],
  ['with-counta', { A22: LABEL, B22: f('COUNTA(B2:B21)'), C22: f(SUB) }, [22]],
  ['with-sum', { A22: LABEL, B22: f(COUNT('B')), C22: f('SUM(C2:C21)') }, [22]],
  ['beside', { A22: LABEL, D22: f(SUB) }, []],
  ['other-range', { A22: LABEL, C22: f('SUBTOTAL(109,E1:E5)'), ...away() }, [22]],
  ['part-range', { A22: LABEL, C22: f('SUBTOTAL(109,C2:C10)') }, [22]],
  ['in-round', { A22: LABEL, C22: f(`ROUND(${SUB},0)`) }, []],
  ['in-sum', { A22: LABEL, C22: f(`1+${SUB}`) }, []],
  ['in-if', { A22: LABEL, C22: f(`IF(TRUE,${SUB})`) }, []],
  ['over-offset', { A22: LABEL, C22: f('SUBTOTAL(109,OFFSET(C2,0,0,20,1))') }, [22]],
  ['over-indirect', { A22: LABEL, C22: f('SUBTOTAL(109,INDIRECT("C2:C21"))') }, [22]],
  ['aggregate-5', { A22: LABEL, C22: f('AGGREGATE(9,5,C2:C21)') }, []],
  ['aggregate-7', { A22: LABEL, C22: f('AGGREGATE(9,7,C2:C21)') }, []],
  ['sum-then-sub', { C22: f('SUM(C2:C21)'), C23: f(SUB) }, [23]],
  ['sub-then-sum', { C22: f(SUB), C23: f('SUM(C2:C21)') }, []],
  ['sub-then-sub', { C22: f(SUB), C23: f(SUB) }, [22, 23]],
  ['sub-then-label', { C22: f(SUB), A23: LABEL }, []],
  ['refers-under', { C22: f(SUB), C23: f('C22') }, []],
  ['refers-beside', { A22: LABEL, C22: f('D22'), D22: f(SUB) }, []],
  ['refers-away', { A22: LABEL, C22: f('H1'), H1: f(SUB) }, []],
  ['as-text', { A22: LABEL, C22: { value: SUB } }, []],
]
const CODES = [...Array(11).keys()].flatMap((k) => [k + 1, k + 101])

const CRITERIA: ActStep = { criteria: 'Север', column: 'A', via: 'model' }
const SORT: ActStep = { sort: 'C', via: 'model' }

/** The rows a case's cells leave out, by id. */
const LEFT_OUT = new Map<string, number[]>()

export function buildActCases(): ActCase[] {
  const out: ActCase[] = []
  const filter = 'A1:C21'
  for (const [name, rows, left] of ROWS) {
    const cells = { ...sales(), ...rows }
    out.push({ id: `${name}-criteria`, cells, filter, steps: [CRITERIA] }, { id: `${name}-sort`, cells, filter, steps: [SORT] })
    LEFT_OUT.set(`${name}-criteria`, left).set(`${name}-sort`, left)
  }
  for (const code of CODES) {
    const id = `code-${code}`
    out.push({ id, cells: { ...sales(), A22: LABEL, C22: f(`SUBTOTAL(${code},C2:C21)`) }, filter, steps: [CRITERIA] })
    LEFT_OUT.set(id, [22])
  }
  const cells = { ...sales(), A22: LABEL, C22: f(SUB) }
  const button: ActStep = { over: 'A1', via: 'button' }
  const model: ActStep = { over: 'A1', via: 'plain' }
  // one book each, in both orders, then both in one book
  ;(['button', 'model', 'button', 'model', 'model', 'button', 'model', 'button'] as const).forEach((how, k) => out.push({ id: `order-${k + 1}-${how}`, cells, steps: [how === 'button' ? button : model] }))
  out.push({ id: 'one-book-model-button', cells, steps: [model, { off: 'model' }, button] }, { id: 'one-book-button-model', cells, steps: [button, { off: 'model' }, model] })
  return out
}

/**
 * The model, from `totals`: what the filter does stops above the rows at the bottom of the range it reads in which a cell
 * of its columns holds a formula that is a SUBTOTAL itself, whatever its code and its range; a row of anything else below
 * them, or a SUBTOTAL inside another function, takes them back in. A sort saves the AutoFilter's own range, criteria the
 * range they filtered. Range(A1).AutoFilter over such a row saves the AutoFilter above it, Data > Filter on A1 saves the range it reads.
 */
export function predictAct(c: ActCase, acted: (own: string) => string): PredictedAct {
  const reads = acted('A1:C21')
  const bottom = Number(/\d+$/.exec(reads)![0])
  if (c.id.startsWith('order') || c.id.startsWith('one-book')) {
    const last = c.steps[c.steps.length - 1] as { via: string }
    const shown = { range: 'A1:C22' }
    return { opens: true, opened: { range: '' }, after: c.steps.map((s) => ('off' in s ? { range: '' } : shown)), saved: { ref: last.via === 'button' ? 'A1:C22' : 'A1:C21' } } as unknown as PredictedAct
  }
  const left = LEFT_OUT.get(c.id)!
  const acts = bottom - left.filter((r) => r <= bottom).length
  const step = c.steps[0]
  if ('sort' in step) return { opens: true, opened: { range: reads }, after: [{ range: reads, hidden: [] }], saved: { ref: 'A1:C21', sortRef: `A2:C${acts}` } } as unknown as PredictedAct
  const hidden: number[] = []
  for (let r = 2; r <= acts; r++) if (r > 21 || regionOf(r) !== 'Север') hidden.push(r)
  return { opens: true, opened: { range: reads }, after: [{ range: reads, hidden, filterMode: true }], saved: { ref: `A1:C${acts}`, criteria: { '0': 'Север' }, hidden } } as unknown as PredictedAct
}
