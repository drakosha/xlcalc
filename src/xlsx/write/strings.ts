/** A rich string stores each run's effective font, including explicit resets of Boolean and automatic properties. */
import type { FontStyle, TextRun } from '../../contract.js'
import type { Issues } from '../issues.js'
import type { CellFormats } from './styles.js'
import { fontXml } from './styles.js'
import { textElement } from '../xml/text.js'

export function stringXml(text:string,runs:TextRun[]|undefined,styles:CellFormats,style?:number,issues?:Issues):string {
  if(!runs?.length)return textElement(text)
  return runs.filter(r=>r.start<text.length).map((run,i)=>{
    const end=Math.min(runs[i+1]?.start??text.length,text.length)
    const font:FontStyle=styles.fontOf(style,run.font,issues)
    let body=fontXml(font).slice(6,-7).replace(/<name /g,'<rFont ').replace(/<\/name>/g,'</rFont>')
    if(!font.bold)body='<b val="0"/>'+body
    if(!font.italic)body=body.replace(/^(<b[^>]*\/>)/,'$1<i val="0"/>')
    if(!font.strike)body=body.replace(/^(<b[^>]*\/><i[^>]*\/>)/,'$1<strike val="0"/>')
    if(!font.underline)body=body.replace(/(<sz\b)/,'<u val="none"/>$1')
    if(!font.color)body=body.replace(/(<rFont\b)/,'<color auto="1"/>$1')
    return `<r><rPr>${body}</rPr>${textElement(text.slice(run.start,end))}</r>`
  }).join('')
}
