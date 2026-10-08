/**
 * Table parts (`xl/tables/tableN.xml`) from a sheet's tables, as Excel writes them: the name twice (`name` and
 * `displayName`), the range, the header and totals row counts, an autofilter over the header and data rows, each
 * column with its name, its totals function or label and its calculated column formula, and Excel's default style.
 * A formula in the totals row that no totals function of its column gives is the column's `custom` totals formula:
 * Excel does not open a file whose totals row holds a formula its table part does not record.
 */
import type { SheetInput, TableInput } from '../../contract.js'
import { formatAddress } from '../../model.js'
import { parseRange } from '../../structure/area.js'
import { toFile } from '../formulas/to-file.js'
import { escapeAttr, escapeText } from '../xml/text.js'
import { fileTree } from './sheet.js'

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

/** The table's part; `id` is unique in the book. */
export function tableXml(t: TableInput, sheet: SheetInput, id: number, sheets: Set<string>, tables: { name: string; columns: string[] }[]): string {
  const x = parseRange(t.range)!
  const headers = t.headers !== false
  const totals = t.totals === true
  const file = (f: string | undefined) => {
    const tree = f === undefined ? null : fileTree(f, sheets, tables, t.name)
    return tree ? escapeText(toFile(tree)) : undefined
  }
  const cols = t.columns.map((c, i) => {
    let attrs = ''
    let inner = ''
    const calc = file(c.formula)
    if (calc) inner += `<calculatedColumnFormula>${calc}</calculatedColumnFormula>`
    // Excel opens a formula in the totals row only with its column's function, `custom` with the formula for any other
    const foot = totals ? sheet.cells[formatAddress(x.r2, x.c1 + i)] : undefined
    const own = foot?.formula === undefined ? undefined : file(foot.formula)
    const fn = c.totals && 'function' in c.totals ? c.totals.function : undefined
    if (own && (fn === undefined || fn === 'custom')) {
      attrs += ' totalsRowFunction="custom"'
      inner += `<totalsRowFormula>${own}</totalsRowFormula>`
    } else if (fn) attrs += ` totalsRowFunction="${fn}"`
    else if (c.totals && 'label' in c.totals) attrs += ` totalsRowLabel="${escapeAttr(c.totals.label)}"`
    return `<tableColumn id="${i + 1}" name="${escapeAttr(c.name)}"${attrs}${inner ? `>${inner}</tableColumn>` : '/>'}`
  })
  const filter = headers ? `<autoFilter ref="${formatAddress(x.r1, x.c1)}:${formatAddress(x.r2 - (totals ? 1 : 0), x.c2)}"/>` : ''
  const counts = (headers ? '' : ' headerRowCount="0"') + (totals ? ' totalsRowCount="1"' : '')
  const name = escapeAttr(t.name)
  const ref = `${formatAddress(x.r1, x.c1)}:${formatAddress(x.r2, x.c2)}`
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="${MAIN}" id="${id}" name="${name}" displayName="${name}" ref="${ref}"${counts}>` +
    `${filter}<tableColumns count="${cols.length}">${cols.join('')}</tableColumns>` +
    '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'
  )
}
