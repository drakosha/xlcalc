/**
 * Colours of styles.xml as `#RRGGBB`: `rgb` (ARGB, the alpha ignored as Excel ignores it), `indexed` (the legacy
 * palette, or the book's own `<indexedColors>`), `theme` with `tint` (the theme's colour lightened or darkened by
 * its luminance, as ECMA-376 defines `tint` and Excel rounds it). `auto` and the system colours are no colour.
 */
import { attributes, element, elements, flag, type Element } from '../xml/scan.js'

/** The legacy palette (ECMA-376 18.8.27): 64 is the system foreground, 65 the background. */
const INDEXED = [
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

/** Windows' HLS scale (hue, luminance, saturation from 0 to 240), in which Excel applies a tint. */
const HLS = 240
const div = (a: number, b: number) => Math.trunc(a / b)

/** RGB (0 to 255) to HLS in whole steps, as Windows' ColorRGBToHLS rounds them. */
function toHls(r: number, g: number, b: number): [h: number, l: number, s: number] {
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = div((max + min) * HLS + 255, 510)
  if (max === min) return [0, l, 0]
  const s = l <= HLS / 2 ? div((max - min) * HLS + div(max + min, 2), max + min) : div((max - min) * HLS + div(510 - max - min, 2), 510 - max - min)
  const d = (c: number) => div((max - c) * (HLS / 6) + div(max - min, 2), max - min)
  let h = r === max ? d(b) - d(g) : g === max ? HLS / 3 + d(r) - d(b) : (2 * HLS) / 3 + d(g) - d(r)
  if (h < 0) h += HLS
  if (h > HLS) h -= HLS
  return [h, l, s]
}

/** HLS in whole steps back to RGB (0 to 255), rounded. */
function toRgb(h: number, l: number, s: number): number[] {
  if (s === 0) return [l, l, l].map((x) => div(x * 255 + HLS / 2, HLS))
  const m2 = l <= HLS / 2 ? div(l * (HLS + s) + HLS / 2, HLS) : l + s - div(l * s + HLS / 2, HLS)
  const m1 = 2 * l - m2
  const hue = (t: number) => {
    t = t < 0 ? t + HLS : t > HLS ? t - HLS : t
    if (t < HLS / 6) return m1 + div((m2 - m1) * t + HLS / 12, HLS / 6)
    if (t < HLS / 2) return m2
    if (t < (2 * HLS) / 3) return m1 + div((m2 - m1) * ((2 * HLS) / 3 - t) + HLS / 12, HLS / 6)
    return m1
  }
  return [h + HLS / 3, h, h - HLS / 3].map((t) => Math.max(0, Math.min(255, div(hue(t) * 255 + HLS / 2, HLS))))
}

/**
 * A colour lightened (tint > 0) or darkened (tint < 0) by its luminance in Windows' HLS steps. Measured against
 * Excel for Mac, this gives its colour exactly for most theme colours and tints, and within 3 of 255 per channel
 * for the rest.
 */
export function tinted(hex: string, tint: number): string {
  if (!tint) return hex
  const [h, l, s] = toHls(...([0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16)) as [number, number, number]))
  const lum = Math.floor(tint < 0 ? l * (1 + tint) : l * (1 - tint) + HLS * tint)
  return toRgb(h, lum, s)
    .map((c) => c.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase()
}

/** The colour of a `<color>`-like element (`<fgColor>`, `<color>` of a font or a border), or undefined. */
export function colorOf(e: Element | undefined, p: Palette): string | undefined {
  if (!e) return undefined
  const a = attributes(e.attrs)
  if (flag(a.auto)) return undefined
  let hex: string | undefined
  if (a.rgb && /^[0-9A-Fa-f]{6,8}$/.test(a.rgb)) hex = a.rgb.slice(-6).toUpperCase()
  else if (a.theme !== undefined) hex = p.theme[Number(a.theme)] || undefined
  else if (a.indexed !== undefined) hex = p.indexed[Number(a.indexed)]
  if (!hex) return undefined
  return `#${tinted(hex, Number(a.tint ?? 0))}`
}
