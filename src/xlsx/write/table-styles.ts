import type { NumberFormat, TableStyleElement, WorkbookTableStyles } from '../../contract.js'
import { escapeAttr } from '../xml/text.js'
import { alignmentXml, borderXml, cellProtectionXml, dxfXml } from './styles.js'
import { themeXml } from './theme.js'
import { counted } from './xml.js'

/**
 * A table style region's differential format, its parts in the order of the schema (font, number format, fill,
 * alignment, protection, border): beyond a rule's, a font's name and size, alignment, protection and the lines inside.
 */
function tableDxf({ style, inside, numberFormat }: TableStyleElement, numFmt: (code: NumberFormat) => number): string {
  const { font, alignment, protection, border } = style
  const sized = `${font?.size === undefined ? '' : `<sz val="${font.size}"/>`}${font?.name === undefined ? '' : `<name val="${escapeAttr(font.name)}"/>`}`
  // the font's flags and colour, number format and fill as a rule's dxf has them
  let head = dxfXml({ ...style, border: undefined, ...(numberFormat === undefined ? {} : { numberFormat }) }, numFmt).slice('<dxf>'.length, -'</dxf>'.length)
  if (sized) head = head.includes('</font>') ? head.replace('</font>', () => `${sized}</font>`) : `<font>${sized}</font>${head}`
  const lines = border || inside ? borderXml(border ?? {}, inside) : ''
  return `<dxf>${head}${alignmentXml(alignment)}${cellProtectionXml(protection)}${lines}</dxf>`
}

export function writeTableStyles(data: WorkbookTableStyles, dxf: (xml: string) => number, numFmt: (code: NumberFormat) => number) {
  const option = (name: string, value: boolean | undefined) => value === undefined ? '' : ` ${name}="${+value}"`
  const styles = Object.entries(data.custom ?? {}).map(([name, def]) => {
    const elements = def.elements.map((element) =>
      `<tableStyleElement type="${escapeAttr(element.type)}"${element.size === undefined ? '' : ` size="${element.size}"`} dxfId="${dxf(tableDxf(element, numFmt))}"/>`)
    return `<tableStyle name="${escapeAttr(name)}"${option('table', def.table)}${option('pivot', def.pivot)} count="${elements.length}">${elements.join('')}</tableStyle>`
  })
  return {
    ...(styles.length ? { styles: counted('tableStyles', styles) } : {}),
    ...(data.theme ? { theme: themeXml(data.theme) } : {}),
  }
}
