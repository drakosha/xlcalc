/**
 * Colours of styles.xml as `#RRGGBB`: `rgb` (ARGB, the alpha ignored as Excel ignores it), `indexed` (the legacy
 * palette, or the book's own `<indexedColors>`), `theme` with `tint` (the theme's colour lightened or darkened by
 * its luminance, as ECMA-376 defines `tint` and Excel rounds it). `auto` and the system colours are no colour.
 */
import { attributes, element, elements, type Element } from '../xml/scan.js'

/** The legacy palette (ECMA-376 18.8.27): 64 is the system foreground, 65 the background. */
export const INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF',
  'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366',
  'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080',
  '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00',
  'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
]

/** Theme colours as a cell's `theme` index names them: the first two pairs are swapped against the theme part's order. */
const THEME_ORDER = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']

export interface Palette {
  theme: string[]
  indexed: string[]
}

/** The theme's colour scheme (`theme1.xml`) by index, and the book's palette (`<indexedColors>` of styles.xml). */
export function palette(themeXml: string | undefined, stylesXml: string | undefined): Palette {
  const scheme = themeXml ? element(themeXml, 'clrScheme')?.body : undefined
  const theme = THEME_ORDER.map((name) => {
    const block = scheme ? element(scheme, name)?.body : undefined
    if (!block) return ''
    const a = attributes(element(block, 'srgbClr')?.attrs ?? element(block, 'sysClr')?.attrs ?? '')
    return (a.val && /^[0-9A-Fa-f]{6}$/.test(a.val) ? a.val : (a.lastClr ?? '')).toUpperCase()
  })
  const own = stylesXml ? element(stylesXml, 'indexedColors')?.body : undefined
  const indexed = own ? elements(own, 'rgbColor').map((e) => (attributes(e.attrs).rgb ?? '').slice(-6).toUpperCase()) : INDEXED
  return { theme, indexed }
}

/** RGB to HSL and back on [0, 1]. */
function toHsl([r, g, b]: number[]): number[] {
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  if (max === min) return [0, 0, l]
  const d = max - min
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4
  return [h / 6, s, l]
}

function fromHsl([h, s, l]: number[]): number[] {
  if (s === 0) return [l, l, l]
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  const hue = (t: number) => {
    t = t < 0 ? t + 1 : t > 1 ? t - 1 : t
    return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p
  }
  return [hue(h + 1 / 3), hue(h), hue(h - 1 / 3)]
}

/** Windows' HLS scale: Excel rounds hue, luminance and saturation to whole steps of 1/240 before applying a tint. */
const HLS = 240

/** A colour lightened (tint > 0) or darkened (tint < 0) by its luminance, as Excel computes it. */
export function tinted(hex: string, tint: number): string {
  if (!tint) return hex
  const [h, s, l] = toHsl([0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255)).map((x) => Math.round(x * HLS))
  const lum = Math.round(tint < 0 ? l * (1 + tint) : l * (1 - tint) + (HLS - HLS * (1 - tint)))
  return fromHsl([h / HLS, s / HLS, lum / HLS])
    .map((c) => Math.round(c * 255).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase()
}

/** The colour of a `<color>`-like element (`<fgColor>`, `<color>` of a font or a border), or undefined. */
export function colorOf(e: Element | undefined, p: Palette): string | undefined {
  if (!e) return undefined
  const a = attributes(e.attrs)
  if (a.auto === '1' || a.auto === 'true') return undefined
  let hex: string | undefined
  if (a.rgb && /^[0-9A-Fa-f]{6,8}$/.test(a.rgb)) hex = a.rgb.slice(-6).toUpperCase()
  else if (a.theme !== undefined) hex = p.theme[Number(a.theme)] || undefined
  else if (a.indexed !== undefined) hex = p.indexed[Number(a.indexed)]
  if (!hex) return undefined
  return `#${tinted(hex, Number(a.tint ?? 0))}`
}
