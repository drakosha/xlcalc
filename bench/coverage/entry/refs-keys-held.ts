/**
 * Held out, typed on the keyboard (Excel for Mac): formulas of the set `refs-held` over its data, to check that typing
 * gives a formula the format of what it refers to as `Formula2Local` does, a cell of a format typed into too.
 */
import { type EntryCase, entryCases } from './cases.js'

export { data } from './refs-held.js'
export const keys = true

export const buildEntryCases = (): EntryCase[] =>
  entryCases({ key: ['=D12', '=D6+D12', '=D12-D6', '=D11-D7', '=D12*2', '=SUM(D6,D3)', '=SUM(E:E)', '=ROUNDUP(D3,0)', '=ABS(D3)', '=U!C1', '=Rate'] }, [['in-pct2', 10, ['=D3', '=D12']]])
