/** Text runs are formatting bound to a constant string; offsets count JavaScript UTF-16 units. */
import { InputError } from '../input-error.js'
import { array, checkNumber, object, onlyKeys } from './basic.js'
import { checkFont } from './styles.js'

export function checkRichText(cell:Record<string,unknown>,path:string,partial=false):void {
  if(cell.richText===undefined)return
  const text=cell.value
  if((!partial||text!==undefined||cell.formula!==undefined)&&(typeof text!=='string'||cell.formula!==undefined))throw new InputError(`${path}.richText`,'text runs need a text constant without a formula')
  const runs=array(cell.richText,`${path}.richText`,'an array of text runs')
  let previous=-1
  runs.forEach((v,i)=>{
    const p=`${path}.richText[${i}]`;const run=object(v,p,'a text run such as { start: 0, font: { bold: true } }')
    onlyKeys(run,p,['start','font'])
    checkNumber(run.start,`${p}.start`,0,Number.MAX_SAFE_INTEGER,true)
    const start=run.start as number
    if((i===0&&start!==0)||start<=previous)throw new InputError(`${p}.start`,'runs start at 0 and increase strictly')
    if(typeof text==='string'&&(start>=text.length||start>0&&/[\uD800-\uDBFF]/.test(text[start-1])&&/[\uDC00-\uDFFF]/.test(text[start])))throw new InputError(`${p}.start`,'a run starts inside the text and does not split a surrogate pair')
    if(run.font!==undefined)checkFont(run.font,`${p}.font`,(o,k)=>o[k]!==undefined,true)
    previous=start
  })
}
