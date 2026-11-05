/**
 * Table parts (`xl/tables/tableN.xml`) from a sheet's tables, as Excel writes them: the name twice (`name` and
 * `displayName`), the range, the header and totals row counts, an autofilter over the header and data rows, each
 * column with its name, its totals function or label and its calculated column formula, and its style (Excel's default
 * without one).
 * A formula in the totals row that no totals function of its column gives is the column's `custom` totals formula:
 * Excel does not open a file whose totals row holds a formula its table part does not record. The table that holds a
 * filtered sheet's filter mode gets criteria every row meets (`KEEP_ALL`).
 */
import type { SheetInput, TableInput, TableStyle, TotalsFunction } from '../../contract.js'
import { DEFAULT_TABLE_STYLE, STYLE_OPTIONS } from '../../tables/check.js'
import { formatAddress } from '../../model.js'
import { cornersText, parseRange } from '../../structure/area.js'
import type { Issues } from '../issues.js'
import { escapeAttr, escapeText } from '../xml/text.js'
import { fileSpelling, fileTree, type FileNames } from './formula.js'
import { KEEP_ALL } from './sheet.js'
import { counted, HEAD, MAIN } from './xml.js'

/** A table's `tableStyleInfo`: Excel's default style for a table without `style`, every option written. */
function styleInfo(s: TableStyle = DEFAULT_TABLE_STYLE): string {
  const name = s.name === undefined ? '' : ` name="${escapeAttr(s.name)}"`
  return `<tableStyleInfo${name}${STYLE_OPTIONS.map(([k, attr]) => ` ${attr}="${s[k] ? 1 : 0}"`).join('')}/>`
}

const SUBTOTALS: Record<Exclude<TotalsFunction, 'custom'>, number> = { average: 101, countNums: 102, count: 103, max: 104, min: 105, stdDev: 107, sum: 109, var: 110 }

/** The formula of a totals function as a file holds it. */
const subtotal = (fn: Exclude<TotalsFunction, 'custom'>, table: string, column: string) => escapeText(`SUBTOTAL(${SUBTOTALS[fn]},${table}[${column.replace(/['#@[\]]/g, "'$&")}])`)

/** The table's part; `id` is unique in the book; `filtered`: the table holds its sheet's filter mode. */
export function tableXml(t: TableInput, sheet: SheetInput, id: number, names: FileNames, issues: Issues, filtered = false): string {
  const x = parseRange(t.range)!
  const headers = t.headers !== false
  const totals = t.totals === true
  const file = (f: string | undefined, cell: string) => {
    const tree = f === undefined ? null : fileTree(f, names, t.name)
    const spelled = tree && fileSpelling(tree)
    if (f !== undefined && !spelled) issues.add('formula', { key: 'tableFormulaOut', table: t.name }, { sheet: sheet.name, cell })
    return spelled ? escapeText(spelled) : undefined
  }
  const cols = t.columns.map((c, i) => {
    let attrs = ''
    let inner = ''
    const calc = file(c.formula, formatAddress(x.r1 + (headers ? 1 : 0), x.c1 + i))
    if (calc) inner += `<calculatedColumnFormula>${calc}</calculatedColumnFormula>`
    // Excel opens a formula in the totals row only with its column's function, `custom` with the formula for any other
    const foot = totals ? sheet.cells[formatAddress(x.r2, x.c1 + i)] : undefined
    const own = foot?.formula === undefined ? undefined : file(foot.formula, formatAddress(x.r2, x.c1 + i))
    const fn = c.totals && 'function' in c.totals ? c.totals.function : undefined
    const label = c.totals && 'label' in c.totals ? c.totals.label : undefined
    // nor one whose cell holds another formula than its function's, a number with neither function nor label, or text other than its label (the set `links/file-refusals`)
    if (own && (fn === undefined || fn === 'custom' || own !== subtotal(fn, t.name, c.name))) {
      attrs += ' totalsRowFunction="custom"'
      inner += `<totalsRowFormula>${own}</totalsRowFormula>`
    } else if (fn) attrs += ` totalsRowFunction="${fn}"`
    else if (typeof foot?.value === 'string') attrs += label === foot.value ? ` totalsRowLabel="${escapeAttr(label)}"` : ''
    else if (foot?.value !== undefined && foot.value !== null) attrs += ' totalsRowFunction="custom"'
    else if (label !== undefined) attrs += ` totalsRowLabel="${escapeAttr(label)}"`
    return `<tableColumn id="${i + 1}" name="${escapeAttr(c.name)}"${attrs}${inner ? `>${inner}</tableColumn>` : '/>'}`
  })
  const filterRef = `${formatAddress(x.r1, x.c1)}:${formatAddress(x.r2 - (totals ? 1 : 0), x.c2)}`
  const filter = headers ? (filtered ? `<autoFilter ref="${filterRef}">${KEEP_ALL}</autoFilter>` : `<autoFilter ref="${filterRef}"/>`) : ''
  const counts = (headers ? '' : ' headerRowCount="0"') + (totals ? ' totalsRowCount="1"' : '')
  const name = escapeAttr(t.name)
  const ref = cornersText(x)
  return (
    `${HEAD}<table xmlns="${MAIN}" id="${id}" name="${name}" displayName="${name}" ref="${ref}"${counts}>` +
    `${filter}${counted('tableColumns', cols)}` +
    `${styleInfo(t.style)}</table>`
  )
}
