/**
 * Text of cells: the shared strings table and inline strings. A string is the text of its `<t>` elements (runs of
 * rich text joined); phonetic runs (`<rPh>`, readings of East Asian text) are not part of it. Rich text is marked:
 * its formatting is lost.
 */
import { readText } from '../xml/text.js'

export interface Strings {
  text: string[]
  /** Indexes of the strings with formatted runs. */
  rich: Set<number>
}

const SI = /<(?:[\w.-]+:)?si(?:\s[^>]*)?(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?si>)/g
const T = /<(?:[\w.-]+:)?t(?:\s[^>]*)?(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?t>)/g
const PHONETIC = /<(?:[\w.-]+:)?rPh\b[\s\S]*?<\/(?:[\w.-]+:)?rPh>|<(?:[\w.-]+:)?phoneticPr\b[^>]*>/g
const RUN = /<(?:[\w.-]+:)?r[\s>]/

/** The plain text of a string item (`<si>`, `<is>`), and whether it has formatted runs. */
export function stringItem(body: string): { text: string; rich: boolean } {
  const rich = RUN.test(body)
  const base = rich || body.includes('rPh') ? body.replace(PHONETIC, '') : body
  let text = ''
  for (const m of base.matchAll(T)) if (m[1]) text += m[1]
  return { text: readText(text), rich }
}

export function readStrings(xml: string | undefined): Strings {
  const out: Strings = { text: [], rich: new Set() }
  if (!xml) return out
  for (const m of xml.matchAll(SI)) {
    const s = stringItem(m[1] ?? '')
    if (s.rich) out.rich.add(out.text.length)
    out.text.push(s.text)
  }
  return out
}
