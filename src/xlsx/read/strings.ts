/**
 * Text of cells: the shared strings table and inline strings. A string is the text of its `<t>` elements (runs of
 * rich text joined); phonetic runs (`<rPh>`, readings of East Asian text) are not part of it. Rich text carries the effective fonts of its runs; unsupported font properties are reported.
 */
import type { FontStyle, TextRun } from '../../contract.js'
import { font } from '../styles/read.js'
import { palette, type Palette } from '../styles/color.js'
import { attributes, element } from '../xml/scan.js'
import type { Issues } from '../issues.js'
import { elements, outside, placed } from '../xml/scan.js'
import { readText } from '../xml/text.js'

export interface Strings {
  text: string[]
  runs?: Map<number, TextRun[]>
  lost?: Set<number>
  /** Indexes of the strings with formatted runs. */
  rich: Set<number>
}

const PHONETIC_PR = /<(?:[\w.-]+:)?phoneticPr\b[^<>]*>/g
const RUN = /<(?:[\w.-]+:)?r[\s>]/
const OPEN_SI = /<(?:[\w.-]+:)?si[\s/>]/g

/** The plain text of a string item (`<si>`, `<is>`), and whether it has formatted runs. */
export function stringItem(body: string, colors:Palette=palette(undefined,undefined), base:FontStyle={}): { text: string; rich: boolean; runs?:TextRun[]; lost?:boolean } {
  const rich=RUN.test(body)
  const plain=body.includes('rPh')?outside(body,placed(body,'rPh')).replace(PHONETIC_PR,''):body
  if(!rich){let text='';for(const t of elements(plain,'t'))if(t.body)text+=readText(t.body);return {text,rich:false}}
  let text='';let lost=false
  const runs:TextRun[]=[]
  for(const e of elements(plain,'r')){
    let value='';for(const t of elements(e.body??'','t'))if(t.body)value+=readText(t.body)
    if(!value)continue
    const pr=element(e.body??'','rPr')
    let own:FontStyle|undefined
    if(pr && /<(?![!?/])/.test(pr.body??'')){
      const read=font(pr,colors);lost ||= read.lost.length>0
      own={name:base.name??'Calibri',size:base.size??11,...read.style}
      const named=element(pr.body??'','rFont')
      if(named)own.name=readText(attributes(named.attrs).val??'')
    }else if(!pr && runs.length)own={name:base.name??'Calibri',size:base.size??11,...base}
    runs.push({start:text.length,...(own?{font:own}:{})});text+=value
  }
  return {text,rich:true,runs,...(lost?{lost:true}:{})}
}

/** The shared strings; items left open (no `</si>`), which are none or run into the next one, are reported. */
export function readStrings(xml: string | undefined, issues: Issues, colors?:Palette, base?:FontStyle): Strings {
  const out: Strings = { text: [], rich: new Set() }
  if (!xml) return out
  for (const si of elements(xml, 'si')) {
    const s = stringItem(si.body ?? '',colors,base)
    if(s.runs)(out.runs??=new Map()).set(out.text.length,s.runs)
    if(s.lost)(out.lost??=new Set()).add(out.text.length)
    if (s.rich) out.rich.add(out.text.length)
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
