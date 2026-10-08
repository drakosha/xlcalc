/**
 * Table parts (`xl/tables/tableN.xml`) as a sheet's tables: the name formulas use (`displayName`), the range, the header
 * and totals rows, each column with its name, its calculated column formula and its totals function or label. A header
 * cell that does not hold its column's name as text is given it, and a table the model cannot hold (over another, over
 * a merged area, a name taken) is left out; both are reported.
 */
import type { CellInput, TableColumn, TableInput, TotalsFunction } from '../../contract.js'
import { InputError } from '../../input.js'
import { formatAddress } from '../../model.js'
import { parseRange } from '../../structure/area.js'
import { checkTables } from '../../tables/check.js'
import { fromFile } from '../formulas/from-file.js'
import type { Issues } from '../issues.js'
import { attributes, element, elements } from '../xml/scan.js'
import { readText } from '../xml/text.js'

const FUNCTIONS = new Set<TotalsFunction>(['sum', 'average', 'count', 'countNums', 'max', 'min', 'stdDev', 'var', 'custom'])

/** A table part's table, or null when it has no name or range. */
function tableOf(xml: string): TableInput | null {
  const root = element(xml, 'table')
  if (!root) return null
  const a = attributes(root.attrs)
  const name = readText(a.displayName ?? a.name ?? '')
  const range = a.ref?.toUpperCase()
  if (!name || !range) return null
  const columns = elements(element(xml, 'tableColumns')?.body ?? '', 'tableColumn').map((e): TableColumn => {
    const c = attributes(e.attrs)
    const col: TableColumn = { name: readText(c.name ?? '') }
    const calc = element(e.body ?? '', 'calculatedColumnFormula')
    if (calc?.body) col.formula = fromFile(readText(calc.body)).formula
    const fn = c.totalsRowFunction as TotalsFunction | undefined
    if (fn && fn !== ('none' as TotalsFunction) && FUNCTIONS.has(fn)) col.totals = { function: fn }
    else if (c.totalsRowLabel !== undefined) col.totals = { label: readText(c.totalsRowLabel) }
    return col
  })
  const t: TableInput = { name, range, columns }
  if (a.headerRowCount !== undefined && Number(a.headerRowCount) === 0) t.headers = false
  if (Number(a.totalsRowCount ?? 0) > 0) t.totals = true
  return t
}

/** The tables of a sheet's table parts; `cells` (the sheet's) get the header names they lack. */
export function readTables(parts: string[], sheet: string, cells: Record<string, CellInput>, merges: string[] | undefined, issues: Issues): TableInput[] {
  const out: TableInput[] = []
  for (const xml of parts) {
    const t = tableOf(xml)
    if (!t) {
      issues.add('table', 'a table part without a name or a range is left out', { sheet })
      continue
    }
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
      issues.add('table', `the table ${t.name} is left out: ${e.message.replace(/^[^:]*: /, '')}`, { sheet })
      continue
    }
    for (const [addr, cell] of fixed) {
      issues.add('table', `a header cell that does not hold its column's name as text is read with the name ${JSON.stringify(cell.value)}`, { sheet, cell: addr })
      cells[addr] = cell
    }
    out.push(t)
  }
  return out
}
