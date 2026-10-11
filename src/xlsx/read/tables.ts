/**
 * Table parts (`xl/tables/tableN.xml`) as a sheet's tables: the name formulas use (`displayName`), the range, the header
 * and totals rows, each column with its name, its calculated column formula and its totals function or label. A header
 * cell that does not hold its column's name as text is given it, and a table the model cannot hold (over another, over
 * a merged area, a name taken) is left out; both are reported, and so is what the engine cannot take of a calculated
 * column formula, as of a cell's. A column formula that differs from its cells' is kept as it is, as Excel keeps it.
 */
import type { CellInput, TableColumn, TableInput, TotalsFunction } from '../../contract.js'
import { InputError } from '../../input.js'
import { formatAddress } from '../../model.js'
import { parseRange } from '../../structure/area.js'
import { checkTables, TOTALS } from '../../tables/check.js'
import type { Issues } from '../issues.js'
import { parseFormula, reportFormula, type ReadFormula } from './formulas.js'
import { attributes, element, elements } from '../xml/scan.js'
import { readText } from '../xml/text.js'

/** A table part's table with its calculated column formulas as read (by column), or null when it has no name or range. */
function tableOf(xml: string): { t: TableInput; formulas: (ReadFormula | undefined)[] } | null {
  const root = element(xml, 'table')
  if (!root) return null
  const a = attributes(root.attrs)
  const name = readText(a.displayName ?? a.name ?? '')
  const range = a.ref?.toUpperCase()
  if (!name || !range) return null
  const formulas: (ReadFormula | undefined)[] = []
  const columns = elements(element(xml, 'tableColumns')?.body ?? '', 'tableColumn').map((e): TableColumn => {
    const c = attributes(e.attrs)
    const col: TableColumn = { name: readText(c.name ?? '') }
    const calc = element(e.body ?? '', 'calculatedColumnFormula')
    const f = calc?.body ? parseFormula(readText(calc.body)) : undefined
    formulas.push(f)
    if (f) col.formula = f.formula
    const fn = c.totalsRowFunction as TotalsFunction | undefined
    // `none`, the file's default, is no function
    if (fn !== undefined && TOTALS.includes(fn)) col.totals = { function: fn }
    else if (c.totalsRowLabel !== undefined) col.totals = { label: readText(c.totalsRowLabel) }
    return col
  })
  const t: TableInput = { name, range, columns }
  if (a.headerRowCount !== undefined && Number(a.headerRowCount) === 0) t.headers = false
  if (Number(a.totalsRowCount ?? 0) > 0) t.totals = true
  return { t, formulas }
}

/** The tables of a sheet's table parts; `cells` (the sheet's) get the header names they lack. */
export function readTables(parts: string[], sheet: string, cells: Record<string, CellInput>, merges: string[] | undefined, issues: Issues): TableInput[] {
  const out: TableInput[] = []
  for (const xml of parts) {
    const read = tableOf(xml)
    if (!read) {
      issues.add('table', { key: 'tablePartBad' }, { sheet })
      continue
    }
    const { t, formulas } = read
    // the header cells the table needs: tried on a copy, kept with the table
    const x = parseRange(t.range)
    const fixed: [string, CellInput][] = []
    if (x && t.headers !== false)
      t.columns.forEach((c, i) => {
        const addr = formatAddress(x.r1, x.c1 + i)
        const cell = cells[addr]
        if (cell?.formula === undefined && cell?.value === c.name) return
        const { formula: _, value: __, array: ___, ...look } = cell ?? {}
        fixed.push([addr, { ...look, value: c.name }])
      })
    try {
      checkTables({ sheets: [{ name: sheet, cells: { ...cells, ...Object.fromEntries(fixed) }, merges, tables: [...out, t] }] }, 'file')
    } catch (e) {
      if (!(e instanceof InputError)) throw e
      issues.add('table', { key: 'tableLeftOut', table: t.name, reason: e.message.replace(/^[^:]*: /, '') }, { sheet })
      continue
    }
    for (const [addr, cell] of fixed) {
      issues.add('table', { key: 'headerGiven', name: cell.value as string }, { sheet, cell: addr })
      cells[addr] = cell
    }
    // a calculated column formula reports what the engine cannot take of it as a cell's would, at the column's first cell
    if (x) formulas.forEach((f, i) => f && reportFormula(f, issues, { sheet, cell: formatAddress(x.r1, x.c1 + i) }))
    out.push(t)
  }
  return out
}
