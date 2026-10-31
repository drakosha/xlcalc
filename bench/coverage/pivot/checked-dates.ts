/**
 * The set `checked-dates`, run in ru-RU: fields of dates checked in Excel's PivotTable Fields list, which Excel 365
 * groups by itself, the fields it makes named in the language of its interface (`Годы (Years)`, `Месяцы (OneYear)`).
 * Over the rows of `kinds`: dates over years, dates within one year, dates with a blank, two fields of dates.
 */
import type { CheckedCase } from './checked.js'

export function buildCheckedCases(): CheckedCase[] {
  return [
    { id: 'Years', fields: [{ name: 'Years' }] },
    { id: 'OneYear', fields: [{ name: 'OneYear' }] },
    { id: 'DateBlank', fields: [{ name: 'DateBlank' }] },
    { id: 'two-dates', fields: [{ name: 'Years' }, { name: 'OneYear' }] },
  ]
}
