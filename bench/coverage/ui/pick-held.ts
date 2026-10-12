/**
 * Held out, predicted by `Session.setCellTexts` (an item read as typed, its alert as typing gives it) before Excel ran it:
 * an item of a cell's dropdown list picked on the keyboard. The list's cells hold text that reads as something else when
 * typed (`001`, `1/2`, `ИСТИНА`, `TRUE`, `5%`), a number shown as a percent and a date, and plain text; the cell picked
 * into is in General, in the text format or a percent. A literal list (`"001,1/2,x"`) and a list of plain text as the
 * control of the keys. Both locales' items are in the set, which runs in each.
 */
import { collector, type UiCase } from './cases.js'
import type { CellInput, DataValidation } from '../../../src/contract.js'

const SOURCE: Record<string, CellInput> = {
  E1: { value: '001' }, E2: { value: '1/2' }, E3: { value: 'ИСТИНА' }, E4: { value: 0.5, numberFormat: '0%' },
  E5: { value: 'TRUE' }, E6: { value: 'abc' }, E7: { value: '5%' }, E8: { value: 46096, numberFormat: 14 },
}
const RANGE: DataValidation = { range: 'B1:B9', type: 'list', formula1: '$E$1:$E$8' }
const LITERAL: DataValidation = { range: 'C1:C9', type: 'list', formula1: '"001,1/2,x"' }
const PLAIN: DataValidation = { range: 'D1:D9', type: 'list', formula1: '"a,b,c"' }

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const pick = (group: string, at: string, item: number, text: string, cell?: CellInput) =>
    add(group, { cells: { ...SOURCE, ...(cell ? { [at]: cell } : {}) }, validations: [RANGE, LITERAL, PLAIN], act: { do: 'pick', at, item, text } })

  // the keys: the second item of a plain list
  pick('plain', 'D2', 2, 'b')
  // a list of cells into a cell in General: each item as the list shows it (the date, last, shows by the locale and is left out)
  const shown = ['001', '1/2', 'ИСТИНА', '50%', 'TRUE', 'abc', '5%']
  shown.forEach((t, i) => pick('cells', 'B2', i + 1, t))
  // into a cell of the text format and of a percent
  pick('text-cell', 'B3', 1, '001', { numberFormat: '@' })
  pick('text-cell', 'B3', 2, '1/2', { numberFormat: '@' })
  pick('pct-cell', 'B4', 4, '50%', { numberFormat: '0%' })
  pick('pct-cell', 'B4', 1, '001', { numberFormat: '0%' })
  // a literal list
  pick('literal', 'C2', 1, '001')
  pick('literal', 'C2', 2, '1/2')
  return out
}
