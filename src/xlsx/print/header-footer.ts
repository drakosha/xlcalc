/** Excel's header/footer codes resolved to three independently formatted sections. */
import type { FontStyle } from '../../contract.js'

/** What header and footer codes print: page number and count, sheet and file labels, date and time. */
export interface PrintHeaderFooterContext {
  pageNumber: number
  pageCount: number
  sheetName: string
  fileName?: string
  /** The path as it should appear before the file name, including any separator. */
  filePath?: string
  /** Date and time as the host wants them printed. */
  date?: string
  time?: string
}

/** The font of a header or footer run. */
export interface PrintHeaderFooterStyle extends Omit<FontStyle, 'color'> {
  /** A font face name whose meaning depends on the font or application locale. */
  face?: string
  themeFont?: 'heading' | 'body'
  color?: string | { theme: number; tint: number }
  script?: 'superscript' | 'subscript'
  outline?: boolean
  shadow?: boolean
}

/** A header or footer run: formatted text, or the place of a picture. */
export type PrintHeaderFooterRun = { text: string; style: PrintHeaderFooterStyle } | { picture: true }
/** A header or footer's left, center and right sections. */
export type PrintHeaderFooterSections = Record<'left' | 'center' | 'right', PrintHeaderFooterRun[]>

type Section = keyof PrintHeaderFooterSections
const ALIGN: Record<string, Section> = { L: 'left', C: 'center', R: 'right' }
const FIELDS = { A: 'sheetName', F: 'fileName', Z: 'filePath', D: 'date', T: 'time' } as const
const TOGGLES = { B: 'bold', I: 'italic', S: 'strike', O: 'outline', H: 'shadow' } as const

function withFont(style: PrintHeaderFooterStyle, quoted: string): PrintHeaderFooterStyle {
  const delimiter = quoted.lastIndexOf(',')
  const name = (delimiter < 0 ? quoted : quoted.slice(0, delimiter)).trim()
  const face = delimiter < 0 ? '' : quoted.slice(delimiter + 1).trim()
  const next = { ...style }
  if (name === '+' || (name === '-' && !face)) {
    next.themeFont = name === '+' ? 'heading' : 'body'
    delete next.name
  } else if (name && name !== '-') {
    next.name = name
    delete next.themeFont
  }
  if (face) {
    const known = face.toLowerCase()
    if (known === 'regular' || known === 'bold' || known === 'italic' || known === 'bold italic' || known === 'italic bold') {
      next.bold = known.includes('bold')
      next.italic = known.includes('italic')
      delete next.face
    } else {
      next.face = face
      delete next.bold
      delete next.italic
    }
  }
  return next
}

function color(text: string, start: number): { value: NonNullable<PrintHeaderFooterStyle['color']>; end: number } | undefined {
  const tail = text.slice(start)
  const rgb = /^[\da-f]{6}/i.exec(tail)
  if (rgb) return { value: '#' + rgb[0].toUpperCase(), end: start + rgb[0].length }
  const theme = /^(\d\d)\.?([+-])(\d{2,3})/.exec(tail)
  if (!theme) return undefined
  const amount = Math.min(100, Number(theme[3]))
  return { value: { theme: Number(theme[1]), tint: theme[2] === '-' ? -amount : amount }, end: start + theme[0].length }
}

/**
 * Resolves a header or footer text into its left, center and right sections of formatted runs, with the page
 * counters, sheet and file labels, date and time of the context. A picture marker keeps its place for the renderer.
 */
export function formatPrintHeaderFooter(text: string, context: PrintHeaderFooterContext): PrintHeaderFooterSections {
  const result: PrintHeaderFooterSections = { left: [], center: [], right: [] }
  const styles: Record<Section, PrintHeaderFooterStyle> = { left: {}, center: {}, right: {} }
  let section: Section = 'center'
  const append = (plain: string) => {
    if (!plain) return
    const runs = result[section]
    const style = styles[section]
    const previous = runs.at(-1)
    if (previous && 'text' in previous && previous.style === style) previous.text += plain
    else runs.push({ text: plain, style })
  }

  let at = 0
  while (at < text.length) {
    const amp = text.indexOf('&', at)
    if (amp < 0) {
      append(text.slice(at))
      break
    }
    append(text.slice(at, amp))
    const code = text[amp + 1]?.toUpperCase() ?? ''
    at = amp + 2
    const style = styles[section]
    if (Object.hasOwn(ALIGN, code)) {
      section = ALIGN[code]
    } else if (Object.hasOwn(FIELDS, code)) {
      append(context[FIELDS[code as keyof typeof FIELDS]] ?? '')
    } else if (Object.hasOwn(TOGGLES, code)) {
      const key = TOGGLES[code as keyof typeof TOGGLES]
      styles[section] = { ...style, [key]: !style[key] }
    } else if (code === 'P' || code === 'N') {
      const offset = /^[+-]\d*/.exec(text.slice(at))?.[0]
      // Excel leaves the last offset digit literal when it ends the entire field.
      const consumed = offset && offset.length > 1 && at + offset.length === text.length ? offset.slice(0, -1) : offset
      const base = code === 'P' ? context.pageNumber : context.pageCount
      append(String(base + (consumed && consumed.length > 1 ? Number(consumed) : 0)))
      if (consumed) at += consumed.length
    } else if (code === 'K') {
      const selected = color(text, at)
      if (selected) {
        styles[section] = { ...style, color: selected.value }
        at = selected.end
      }
    } else if (code === '"') {
      const end = text.indexOf('"', at)
      if (end >= 0) {
        styles[section] = withFont(style, text.slice(at, end))
        at = end + 1
      }
    } else if (code >= '0' && code <= '9') {
      const size = /^\d{1,3}/.exec(text.slice(amp + 1))![0]
      styles[section] = { ...style, size: Number(size) }
      at = amp + 1 + size.length
    } else if (code === 'U' || code === 'E') {
      const underline = code === 'U' ? 'single' : 'double'
      const next = { ...style }
      if (style.underline === underline) delete next.underline
      else next.underline = underline
      styles[section] = next
    } else if (code === 'X' || code === 'Y') {
      const script = code === 'X' ? 'superscript' : 'subscript'
      if (!style.script) styles[section] = { ...style, script }
      else if (style.script === script) {
        const next = { ...style }
        delete next.script
        styles[section] = next
      }
    } else if (code === '&') {
      append('&')
    } else if (code === 'G') {
      result[section].push({ picture: true })
    }
  }
  return result
}
