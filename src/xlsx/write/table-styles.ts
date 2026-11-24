import type { CellStyle, WorkbookTableStyles } from '../../contract.js'
import { escapeAttr } from '../xml/text.js'
import { alignmentXml, cellProtectionXml, dxfXml } from './styles.js'
import { themeXml } from './theme.js'
import { counted } from './xml.js'

/** A table style's differential format can also name a font and carry alignment and protection. */
function tableDxf(style: CellStyle): string {
  const { font, alignment, protection } = style
  const size = font?.size === undefined ? '' : `<sz val="${font.size}"/>`
  const name = font?.name === undefined ? '' : `<name val="${escapeAttr(font.name)}"/>`
  const align = alignmentXml(alignment)
  const protect = cellProtectionXml(protection)
  const base = dxfXml(style, () => 0).replace('</font>', () => `${size}${name}</font>`)
  return base.replace(/(<border>|<\/dxf>)/, `${align}$1`).replace('</dxf>', `${protect}</dxf>`)
}

export function writeTableStyles(data: WorkbookTableStyles, dxf: (xml: string) => number) {
  const option = (name: string, value: boolean | undefined) => value === undefined ? '' : ` ${name}="${+value}"`
  const styles = Object.entries(data.custom ?? {}).map(([name, def]) => {
    const elements = def.elements.map(({ type, size, style }) =>
      `<tableStyleElement type="${escapeAttr(type)}"${size === undefined ? '' : ` size="${size}"`} dxfId="${dxf(tableDxf(style))}"/>`)
    return `<tableStyle name="${escapeAttr(name)}"${option('table', def.table)}${option('pivot', def.pivot)} count="${elements.length}">${elements.join('')}</tableStyle>`
  })
  return {
    ...(styles.length ? { styles: counted('tableStyles', styles) } : {}),
    ...(data.theme ? { theme: themeXml(data.theme) } : {}),
  }
}
