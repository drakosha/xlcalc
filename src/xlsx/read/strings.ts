/**
 * Text of cells: the shared strings table and inline strings. A string is the text of its `<t>` elements (runs of
 * rich text joined); phonetic runs (`<rPh>`, readings of East Asian text) are not part of it. Rich text carries the
 * effective fonts of its runs; the font properties the model has no place for are reported.
 */
import type { FontStyle, TextRun } from '../../contract.js'
import type { Issues } from '../issues.js'
import { palette, type Palette } from '../styles/color.js'
import { font } from '../styles/read.js'
import { attributes, element, elements, outside, placed } from '../xml/scan.js'
import { readText } from '../xml/text.js'

export interface Strings {
  text: string[]
  /** The runs of the strings with formatted runs, by index. */
  runs?: Map<number, TextRun[]>
  /** The indexes of the strings whose runs have a property the model has no place for. */
  lost?: Set<number>
}

const PHONETIC_PR = /<(?:[\w.-]+:)?phoneticPr\b[^<>]*>/g
const RUN = /<(?:[\w.-]+:)?r[\s>]/
const OPEN_SI = /<(?:[\w.-]+:)?si[\s/>]/g

/** The palette of a file without a theme or colors of its own, made once. */
let plainPalette: Palette | undefined

/** The text of `<t>` elements of `body`, joined. */
function textOf(body: string): string {
  let text = ''
  for (const t of elements(body, 't')) if (t.body) text += readText(t.body)
  return text
}

/**
 * The plain text of a string item (`<si>`, `<is>`), and with formatted runs their fonts: a run's own properties over the
 * workbook's font (`base`); the first run without them none, as it takes its cell's; a later one the workbook's font.
 */
export function stringItem(body: string, colors?: Palette, base: FontStyle = {}): { text: string; rich: boolean; runs?: TextRun[]; lost?: boolean } {
  const rich = RUN.test(body)
  const plain = body.includes('rPh') ? outside(body, placed(body, 'rPh')).replace(PHONETIC_PR, '') : body
  if (!rich) return { text: textOf(plain), rich: false }
  const p = colors ?? (plainPalette ??= palette(undefined, undefined))
  const whole = { name: base.name ?? 'Calibri', size: base.size ?? 11 }
  let text = ''
  let lost = false
  const runs: TextRun[] = []
  for (const e of elements(plain, 'r')) {
    const value = textOf(e.body ?? '')
    if (!value) continue
    const pr = element(e.body ?? '', 'rPr')
    let own: FontStyle | undefined
    if (pr && /<(?![!?/])/.test(pr.body ?? '')) {
      const read = font(pr, p)
      lost ||= read.lost.length > 0
      own = { ...whole, ...read.style }
      const named = element(pr.body ?? '', 'rFont')
      if (named) own.name = readText(attributes(named.attrs).val ?? '')
    } else if (!pr && runs.length) own = { ...whole, ...base }
    runs.push({ start: text.length, ...(own ? { font: own } : {}) })
    text += value
  }
  return { text, rich: true, runs, ...(lost ? { lost: true } : {}) }
}

/** The shared strings; items left open (no `</si>`), which are none or run into the next one, are reported. */
export function readStrings(xml: string | undefined, issues: Issues, colors?: Palette, base?: FontStyle): Strings {
  const out: Strings = { text: [] }
  if (!xml) return out
  for (const si of elements(xml, 'si')) {
    const s = stringItem(si.body ?? '', colors, base)
    if (s.runs) (out.runs ??= new Map()).set(out.text.length, s.runs)
    if (s.lost) (out.lost ??= new Set()).add(out.text.length)
    out.text.push(s.text)
  }
  let opened = 0
  OPEN_SI.lastIndex = 0
  while (OPEN_SI.exec(xml)) opened++
  if (opened > out.text.length) {
    issues.add('part', { key: 'stringsOpen' })
    // the items left open count: a cell that names one reads empty text, as the issue says
    while (out.text.length < opened) out.text.push('')
  }
  return out
}
