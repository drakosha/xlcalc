/**
 * An AutoFilter on A1:C21 with criteria on column A, its filter given another range: narrower, wider (column D has a
 * header and data), without the column of the criteria, and aside (E1:F5 has a header and data). Through `Range.AutoFilter`
 * with a field and with no arguments, and with the filter removed first: `AutoFilterMode = false` and `Range.AutoFilter`,
 * or Data > Filter twice, the range selected the second time.
 */
import type { CellInput } from '../../../src/contract.js'
import { sales, type ActCase, type ActStep } from './cases.js'

const CRITERIA: ActStep = { criteria: 'Север', column: 'A', via: 'model' }

function columnD(): Record<string, CellInput> {
  const out: Record<string, CellInput> = { D1: { value: 'Доля' } }
  for (let r = 2; r <= 21; r++) out[`D${r}`] = { value: r }
  return out
}

function aside(): Record<string, CellInput> {
  const out: Record<string, CellInput> = { E1: { value: 'Код' }, F1: { value: 'Вес' } }
  for (let r = 2; r <= 5; r++) Object.assign(out, { [`E${r}`]: { value: `к${r}` }, [`F${r}`]: { value: r } })
  return out
}

const RANGES: [string, string, () => Record<string, CellInput>][] = [
  ['a', 'A1:C10', sales],
  ['b', 'A1:D21', () => ({ ...sales(), ...columnD() })],
  ['c', 'B1:C21', sales],
  ['d', 'E1:F5', () => ({ ...sales(), ...aside() })],
]

export function buildActCases(): ActCase[] {
  const out: ActCase[] = []
  for (const [name, over, cells] of RANGES) {
    const base = { cells: cells(), filter: 'A1:C21' }
    out.push({ id: `${name}-field`, ...base, steps: [CRITERIA, { over, via: 'field' }] })
    out.push({ id: `${name}-plain`, ...base, steps: [CRITERIA, { over, via: 'plain' }] })
    out.push({ id: `${name}-off-model`, ...base, steps: [CRITERIA, { off: 'model' }, { over, via: 'plain' }] })
    out.push({ id: `${name}-off-button`, ...base, steps: [CRITERIA, { off: 'button' }, { over, via: 'button' }] })
  }
  return out
}
