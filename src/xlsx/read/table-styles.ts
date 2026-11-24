import type { TableStyleDefinition, WorkbookTableStyles } from '../../contract.js'
import type { Palette } from '../styles/color.js'
import { dxfCellStyle, type Dxf } from '../styles/dxf.js'
import { attributes, elements, flag, within } from '../xml/scan.js'
import { readText } from '../xml/text.js'

/** The theme palette and custom styles of styles.xml, whose regions share the workbook's differential formats. */
export function readTableStyles(xml: string | undefined, colors: Palette, dxfs: Dxf[]): WorkbookTableStyles | undefined {
  const custom = Object.create(null) as Record<string, TableStyleDefinition>
  for (const entry of within(xml ?? '', 'tableStyles', 'tableStyle')) {
    const a = attributes(entry.attrs)
    const name = readText(a.name ?? '')
    if (!name) continue
    const definition: TableStyleDefinition = { elements: [] }
    if (a.table !== undefined) definition.table = flag(a.table)
    if (a.pivot !== undefined) definition.pivot = flag(a.pivot)
    for (const region of elements(entry.body ?? '', 'tableStyleElement')) {
      const { type, size, dxfId } = attributes(region.attrs)
      const dxf = dxfs[Number(dxfId)]
      if (!type) continue
      const style = dxf ? dxfCellStyle(dxf) : {}
      if (!style.fill && dxf?.pattern) style.fill = { color: dxf.pattern }
      const width = Number(size)
      definition.elements.push({ type, ...(Number.isInteger(width) && width > 0 && width <= 0xffffffff ? { size: width } : {}), style })
    }
    custom[name] = definition
  }
  const theme = colors.theme.every((color) => /^[\dA-F]{6}$/.test(color)) ? colors.theme.map((color) => `#${color}`) : undefined
  const hasCustom = Object.keys(custom).length > 0
  return theme || hasCustom ? { ...(theme ? { theme } : {}), ...(hasCustom ? { custom } : {}) } : undefined
}
