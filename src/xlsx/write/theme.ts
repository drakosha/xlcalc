import { THEME_ORDER } from '../styles/color.js'
import { HEAD } from './xml.js'

/** The workbook's palette in DrawingML order, with the required font and format schemes. */
export function themeXml(colors: string[]): string {
  const order = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11]
  const palette = order.map((i) => {
    const name = THEME_ORDER[i]
    return `<a:${name}><a:srgbClr val="${colors[i].slice(1)}"/></a:${name}>`
  }).join('')
  const font = (kind: string, name: string) => `<a:${kind}Font><a:latin typeface="${name}"/><a:ea typeface=""/><a:cs typeface=""/></a:${kind}Font>`
  const solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
  const lines = [6350, 12700, 19050].map((w) => `<a:ln w="${w}">${solid}</a:ln>`).join('')
  const effects = '<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3)
  return `${HEAD}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements>`
    + `<a:clrScheme name="Office">${palette}</a:clrScheme>`
    + `<a:fontScheme name="Office">${font('major', 'Calibri Light')}${font('minor', 'Calibri')}</a:fontScheme>`
    + `<a:fmtScheme name="Office"><a:fillStyleLst>${solid.repeat(3)}</a:fillStyleLst><a:lnStyleLst>${lines}</a:lnStyleLst>`
    + `<a:effectStyleLst>${effects}</a:effectStyleLst><a:bgFillStyleLst>${solid.repeat(3)}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`
}
