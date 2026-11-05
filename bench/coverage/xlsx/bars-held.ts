/**
 * The lengths of an Excel 2010 data bar, which a file holds twice: on the 2007 `<dataBar>` and on `<x14:dataBar>`. [MS-XLSX] 2.6.30
 * says that with lengths other than 0 and 100 the two MUST be equal, and with 0 and 100 the 2007 part MUST be 10 and 90. Each case
 * is a file `writeXlsx` writes with the 2007 part changed by hand (`old`), opened and saved by Excel: `npx tsx bench/coverage/run.ts
 * xlsx bars-held` and `npx tsx bench/windows/run.ts xlsx bars-held` keep the saved files in `xlsx/bars-held/` of each Excel's answers.
 * `made` (Windows alone: Excel for Mac's AppleScript crashes making a rule) is a bar Excel makes itself, `PercentMin` 5 and
 * `PercentMax` 95, saved into `xlsx/bars-held-made/`. `bars-held.predicted.json` was committed before Excel ran the set: the lengths
 * of both parts in the saved file, an attribute left out as absent, and the lengths the reader gives.
 */
import type { ConditionalFormat } from '../../../src/contract.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { rawFile } from './sent.js'

export interface BarCase {
  id: string
  minLength: number
  maxLength: number
  /** The attributes the 2007 part is given by hand; without it the part stays as the writer leaves it. */
  old?: string
}

export const BAR_CASES: BarCase[] = [
  { id: 'x14-only', minLength: 5, maxLength: 95 },
  { id: 'both', minLength: 5, maxLength: 95, old: ' minLength="5" maxLength="95"' },
  { id: 'both-20-80', minLength: 20, maxLength: 80, old: ' minLength="20" maxLength="80"' },
  { id: 'min-only', minLength: 5, maxLength: 100, old: ' minLength="5" maxLength="100"' },
  { id: 'max-only', minLength: 0, maxLength: 95, old: ' minLength="0" maxLength="95"' },
  { id: 'default-0-100', minLength: 0, maxLength: 100 },
  { id: 'both-10-90', minLength: 10, maxLength: 90 },
]

export const barRule = (c: { minLength: number; maxLength: number }): ConditionalFormat => ({ type: 'dataBar', ranges: ['A1:A4'], min: { type: 'min' }, max: { type: 'max' }, color: '#638EC6', minLength: c.minLength, maxLength: c.maxLength })

const cells = { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, A4: { value: 4 } }

export async function barFile(c: BarCase): Promise<Uint8Array> {
  const { data } = await writeXlsx({ book: { sheets: [{ name: 'S', cells, conditionalFormats: [barRule(c)] }] } })
  return c.old ? rawFile(data, [{ part: 'xl/worksheets/sheet1.xml', from: '<dataBar>', to: `<dataBar${c.old}>` }]) : data
}

/** The book of `made`: the same cells, the bar Excel's to make. */
export async function barBook(): Promise<Uint8Array> {
  return (await writeXlsx({ book: { sheets: [{ name: 'S', cells }] } })).data
}
