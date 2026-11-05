/**
 * A host's removal of a filter the file holds, opened in Excel: `writeKept` over `setLayout` that marks the sheet not
 * filtered, shows rows 2:21 and removes the AutoFilter. The criteria in the form a file by another program holds them,
 * over 15 hidden rows with the total row shown; as Excel writes them, with its `_xlnm._FilterDatabase`; and with the
 * total row hidden too, under the rows shown.
 */
import { regionOf, sales, total, type ActCase } from './cases.js'

const failing = Array.from({ length: 20 }, (_, k) => k + 2).filter((r) => regionOf(r) !== 'Север')
const CUSTOM = '<filterColumn colId="0"><customFilters><customFilter val=""/><customFilter operator="notEqual" val=""/></customFilters></filterColumn>'
const VALUES = '<filterColumn colId="0"><filters><filter val="Север"/></filters></filterColumn>'

export function buildActCases(): ActCase[] {
  const cells = { ...sales(), ...total(22, 'SUM(C2:C21)') }
  const base = { cells, filter: 'A1:C21', keptOff: '2:21', steps: [] }
  return [
    { id: 'custom', ...base, hidden: failing, fileCriteria: CUSTOM },
    { id: 'values', ...base, hidden: failing, fileCriteria: VALUES, database: true },
    { id: 'total-hidden', ...base, hidden: [...failing, 22], fileCriteria: VALUES, watch: [22] },
  ]
}
