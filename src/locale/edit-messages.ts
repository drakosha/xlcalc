/**
 * en-US's text of the messages only an edit of a book gives (`EditKey`): a structure operation's refusals, typing into
 * a cell, the words of `explain`. A module of its own, out of `compute`'s bundle: `say` takes it, `english` does not.
 */
import { ERROR_TEXT } from '../syntax/errors.js'
import { english } from './english.js'
import type { EditMessages } from './messages.js'

export const EDIT_MESSAGES: EditMessages = {
  sheetNameTaken: (p) => `a sheet named ${JSON.stringify(p.name)} already exists (sheet names are case-insensitive)`,
  lastSheet: () => 'a workbook must contain at least one sheet',
  lastVisibleSheet: (p) => `a workbook must contain at least one visible sheet, and ${JSON.stringify(p.sheet)} is the last one`,
  insertOffSheet: (p) => `inserting ${p.count} ${p.lines} would push ${p.sheet}!${p.cell} off the sheet`,
  insertMergeOffSheet: (p) => `inserting ${p.count} ${p.lines} would push a merged area of ${p.sheet} off the sheet`,
  arrayPart: (p) => `the operation would change part of the array formula ${p.sheet}!${p.cell} (${p.area})`,
  moveOffSheet: (p) => `the range does not fit on ${p.sheet} at ${p.cell}`,
  copyOffSheet: (p) => `the copy does not fit on ${p.sheet} at ${p.cell}`,
  writesLimit: (p) => `one edit writes ${p.size} cells, over the limit of ${p.max}`,
  moveCutsMerge: (p) => `the moved cells would split the merged area ${p.sheet}!${p.area}`,
  pasteOnMerge: (p) => `the pasted cells would cover part of the merged area ${p.sheet}!${p.area}`,
  contentOnMerge: (p) => `values or formulas cannot be pasted onto the merged area ${p.sheet}!${p.area}`,
  filteredPaste: (p) => `the copy skips the rows a filter hid, so it cannot be pasted over the hidden rows of ${p.sheet}!${p.area}`,
  tableLands: (p) => `the table ${p.table} would overlap another table`,
  tableInsertTwo: (p) => `inserting ${p.lines} there would shift cells of the tables ${p.tables.join(' and ')} at once`,
  tableDeleteTwo: (p) => `deleting these ${p.lines} would shift cells of the tables ${p.tables.join(' and ')} at once`,
  tableHeader: (p) => `the header row of the table ${p.table} cannot be deleted without the rest of the table`,
  tableName: (p) => `${JSON.stringify(p.name)}: ${english({ key: 'nameRule', name: p.name, rule: p.rule })}`,
  tableNameTaken: (p) => `a table named ${p.table} already exists (table names are case-insensitive)`,
  tableNameDefined: (p) => `${JSON.stringify(p.name)} is a defined name (tables and names share one namespace)`,
  tableOverTable: (p) => `${p.range} overlaps the table ${p.table}`,
  tableOverMerge: (p) => `${p.range} overlaps the merged area ${p.area}`,
  tableOverFilter: (p) => `${p.range} overlaps the sheet's AutoFilter over ${p.filter}: a table takes its place over all of it alone`,
  mergeInTable: (p) => `${p.range} overlaps the table ${p.table}, whose cells cannot be merged`,
  arrayInTable: (p) => `the array formula over ${p.area} overlaps the table ${p.table} (Excel does not allow array formulas of several cells in tables)`,
  headerFormula: (p) => `${p.sheet}!${p.cell} is a header cell of the table ${p.table}: it holds its column's name, not a formula`,
  validationError: () => "This value doesn't match the data validation restrictions defined for this cell.",
  explainCells: (p) => `${p.count} cells`,
  explainSpill: (p) => `spilled from ${p.ref}`,
  // the value as JSON, an error as a cell shows it (`#DIV/0!`)
  explainValue: ({ value: v }) => (v === undefined ? '' : v === null ? 'empty' : typeof v === 'object' ? ERROR_TEXT[v.error] : JSON.stringify(v)),
  explainFormula: (p) => p.formula,
  explainErrorFrom: (p) => `(error from ${p.ref})`,
  explainErrorHere: () => '(the error begins here)',
  explainHeldFrom: (p) => `(held by a circular reference through ${p.ref})`,
  explainHeldHere: () => '(held by a circular reference)',
}
