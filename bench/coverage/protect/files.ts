/**
 * Protection, the first set: a sheet's protection and its cells' as `writeXlsx` writes them (Excel's default, what it
 * allows), read back by Excel; files changed by hand for how Excel reads them (`applyProtection="0"`, a cell style's
 * protection, `<sheetProtection>` without `sheet="1"` or with it alone, a password); where a cell's protection goes in
 * Excel (a copy, Paste Special Formats and Values, rows inserted below an unlocked cell, Clear Formats); Excel setting
 * locked and formula hidden and protecting a sheet; and files `writeKept` wrote over a protected one (a value changed,
 * the protection changed or removed, rows inserted, a sheet protected anew). Predicted before Excel ran it
 * (`files.predicted.json`).
 */
import type { BookInput } from '../../../src/contract.js'
import type { RawEdit } from '../xlsx/sent.js'
import type { ProtectCase, ProtectCases } from './cases.js'

/** S: A1 unlocked, A2 as by default, A3 a formula bold and hidden, B1:B3 values; T: values. */
const cells = (): BookInput['sheets'][number]['cells'] => ({ A1: { value: 1, style: 0 }, A2: { value: 2 }, A3: { formula: 'A1+A2', style: 1 }, B1: { value: 'b1' }, B2: { value: 'b2' }, B3: { value: 'b3' } })
const STYLES = [{ protection: { locked: false } }, { font: { bold: true }, protection: { hidden: true } }]

const book = (protection?: BookInput['sheets'][number]['protection']): BookInput => ({
  styles: STYLES,
  sheets: [
    { name: 'S', cells: cells(), ...(protection ? { protection } : {}) },
    { name: 'T', cells: { A1: { value: 't' }, B2: { value: 5 } } },
  ],
})

const S = 'xl/worksheets/sheet1.xml'
const STYLES_PART = 'xl/styles.xml'
const GUARD = '<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0" selectLockedCells="1"/>'
const sheetTag = (to: string): RawEdit => ({ part: S, from: GUARD, to })
const READ: ProtectCase['read'] = [['S', 'A1:B4']]

export function buildProtectCases(): ProtectCases {
  const cases: ProtectCase[] = [
    // written: the protection Excel reads, what it allows, the saved file read back
    { id: 'written', book: 'guarded', read: READ },
    { id: 'written-default', book: 'default', read: READ },
    { id: 'written-allow', book: 'allowing', read: READ },
    // read as Excel reads it
    { id: 'apply-off', book: 'guarded', raw: [{ part: STYLES_PART, from: 'applyProtection="1"><protection locked="0"/>', to: 'applyProtection="0"><protection locked="0"/>' }], read: READ },
    {
      id: 'cell-style',
      book: 'guarded',
      raw: [{ part: STYLES_PART, from: '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>', to: '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyProtection="1"><protection locked="0" hidden="1"/></xf></cellStyleXfs>' }],
      read: READ,
    },
    { id: 'sheet-missing', book: 'guarded', raw: [sheetTag('<sheetProtection objects="1" scenarios="1"/>')], read: READ },
    { id: 'sheet-zero', book: 'guarded', raw: [sheetTag('<sheetProtection sheet="0" objects="1" scenarios="1"/>')], read: READ },
    { id: 'sheet-alone', book: 'guarded', raw: [sheetTag('<sheetProtection sheet="1"/>')], read: READ },
    { id: 'sheet-password', book: 'guarded', raw: [sheetTag('<sheetProtection password="CC1A" sheet="1" objects="1" scenarios="1" insertRows="0"/>')], read: READ },
    // where a cell's protection goes
    { id: 'copy', book: 'plain', steps: [{ copy: { sheet: 'S', range: 'A1:A3' }, to: { sheet: 'S', range: 'C1' } }], read: [['S', 'C1:C3']] },
    { id: 'paste-formats', book: 'plain', steps: [{ copy: { sheet: 'S', range: 'A1:A3' }, to: { sheet: 'T', range: 'B1' }, paste: 'formats' }], read: [['T', 'B1:B4']] },
    { id: 'paste-values', book: 'plain', steps: [{ copy: { sheet: 'S', range: 'A1:A3' }, to: { sheet: 'S', range: 'C1' }, paste: 'values' }], read: [['S', 'C1:C3']] },
    { id: 'paste-all-over', book: 'plain', steps: [{ copy: { sheet: 'S', range: 'B1:B3' }, to: { sheet: 'S', range: 'A1' }, paste: 'all' }], read: READ },
    { id: 'insert-below', book: 'plain', steps: [{ insertRows: { sheet: 'S', range: '2:3' } }], read: [['S', 'A1:B6']] },
    { id: 'insert-above', book: 'plain', steps: [{ insertRows: { sheet: 'S', range: '1:1' } }], read: READ },
    { id: 'clear-formats', book: 'plain', steps: [{ clearFormats: { sheet: 'S', range: 'A1:A2' } }], read: READ },
    // Excel sets them
    { id: 'lock-cells', book: 'plain', steps: [{ lock: { sheet: 'T', range: 'A1:B2' }, locked: false, hidden: false }, { lock: { sheet: 'T', range: 'B2' }, locked: true, hidden: true }], read: [['T', 'A1:B2']] },
    { id: 'protect-dialog', book: 'plain', steps: [{ protectSheet: 'T', allow: [] }], read: [['T', 'A1:B2']] },
    { id: 'protect-allow', book: 'plain', steps: [{ protectSheet: 'T', allow: ['formatCells', 'insertRows', 'sort', 'autoFilter'] }], read: [['T', 'A1:B2']] },
    { id: 'unprotect', book: 'guarded', steps: [{ unprotectSheet: 'S' }], read: READ },
    // written over a protected file with a password
    { id: 'kept-value', book: 'guarded', raw: [sheetTag(PASSWORDED)], kept: { cells: [{ sheet: 'S', cell: 'A1', input: { value: 10, style: 0 } }, { sheet: 'S', cell: 'A2', input: { value: 20 } }] }, read: READ },
    { id: 'kept-protection', book: 'guarded', raw: [sheetTag(PASSWORDED)], kept: { layout: [{ sheet: 'S', protection: { insertRows: true, formatCells: true } }] }, read: READ },
    { id: 'kept-unprotected', book: 'guarded', raw: [sheetTag(PASSWORDED)], kept: { layout: [{ sheet: 'S', protection: null }] }, read: READ },
    { id: 'kept-insert', book: 'plain', kept: { op: { op: 'insertRows', sheet: 'S', row: 1 } }, read: READ },
    { id: 'kept-new', book: 'plain', kept: { layout: [{ sheet: 'T', protection: { deleteRows: true } }], cells: [{ sheet: 'T', cell: 'A1', input: { value: 'u', style: 0 } }] }, read: [['T', 'A1:B2']] },
  ]
  return { books: { guarded: book({ formatCells: true, selectLockedCells: false }), default: book({}), allowing: book(ALLOWING), plain: book() }, cases }
}

/** A protection with the legacy password's hash, which Excel asks for to unprotect the sheet. */
const PASSWORDED = '<sheetProtection password="CC1A" sheet="1" objects="1" scenarios="1" formatCells="0" selectLockedCells="1"/>'

const ALLOWING = { objects: true, scenarios: true, sort: true, autoFilter: true, pivotTables: true, insertHyperlinks: true, deleteRows: true, selectLockedCells: false, selectUnlockedCells: false }
