/**
 * Text of cells: the shared strings table and inline strings. A string is the text of its `<t>` elements (runs of
 * rich text joined); phonetic runs (`<rPh>`, readings of East Asian text) are not part of it. Rich text is marked:
 * its formatting is lost.
 */
import { elements, outside, placed } from '../xml/scan.js'
import { readText } from '../xml/text.js'

export interface Strings {
  text: string[]
  /** Indexes of the strings with formatted runs. */
  rich: Set<number>
}

const PHONETIC_PR = /<(?:[\w.-]+:)?phoneticPr\b[^<>]*>/g
const RUN = /<(?:[\w.-]+:)?r[\s>]/

/** The plain text of a string item (`<si>`, `<is>`), and whether it has formatted runs. */
export function stringItem(body: string): { text: string; rich: boolean } {
  const rich = RUN.test(body)
  const base = rich || body.includes('rPh') ? outside(body, placed(body, 'rPh')).replace(PHONETIC_PR, '') : body
  let text = ''
  for (const t of elements(base, 't')) if (t.body) text += t.body
  return { text: readText(text), rich }
}

export function readStrings(xml: string | undefined): Strings {
  const out: Strings = { text: [], rich: new Set() }
  if (!xml) return out
  for (const si of elements(xml, 'si')) {
    const s = stringItem(si.body ?? '')
    if (s.rich) out.rich.add(out.text.length)
    out.text.push(s.text)
  }
  return out
}
