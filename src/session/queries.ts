/** Bounded queries executed in the session: no whole-sheet value copy or additional retained index. */
import type { CellReadField, Clock, DependentCell, DependentsOptions, DependentsResult, FindCellsOptions, PageOptions, QueryPage, RangeRef, RangeSummary, ReadCell, ReadCellsOptions, ReadCellsResult } from '../contract.js'
import { toCellValue, valueAt, type EngineState } from '../engine.js'
import { CalcError, paced, type Steps } from '../exec/budget.js'
import { referencesOf } from '../graph.js'
import { array, checkBoolean, checkNumber, checkSheet, checkString, InputError, object, onlyKeys } from '../input.js'
import { issueDetail } from '../issues.js'
import { formatAddress, parseAddress } from '../model.js'
import { cellResult, cellsIn, outValue } from '../result.js'
import { rectText, hasCell, meets, type Rect } from '../structure/area.js'
import { sheetKey } from '../syntax.js'
import { CompensatedSum } from '../functions/statistical/sum.js'
import { checkBase } from './answer.js'
import { checkRange } from './viewport.js'

const PAGE=['maxCells','cursor','count'] as const
const FIELDS: CellReadField[]=['value','text','formula','style','numberFormat','note','link','issue']
interface Ref extends Rect { s: number }
const refsOf=(s: EngineState, refs: RangeRef[]) => array(refs,'ranges','an array of sheet/range references').map((v,i): Ref => {
 const p=`ranges[${i}]`,o=object(v,p,'a sheet and range');onlyKeys(o,p,['sheet','range'])
 return {...checkRange(o.range,`${p}.range`),s:checkSheet(o.sheet,n=>s.sheetIdx.get(sheetKey(n)),`${p}.sheet`)}
})

class Page<T> {
 readonly items: T[]=[]
 readonly max: number
 readonly offset: number
 readonly count: boolean
 matched=0
 more=false
 constructor(readonly revision:number, readonly clock:Clock, readonly signature:string, o:PageOptions) {
  this.max=o.maxCells??1000;checkNumber(this.max,'opts.maxCells',1,10000,true)
  if(o.count!==undefined)checkBoolean(o.count,'opts.count')
  this.count=o.count===true;let offset=0
  if(o.cursor!==undefined){
   let c: unknown
   try {c=JSON.parse(checkString(o.cursor,'opts.cursor'))}catch{throw new InputError('opts.cursor','expected a cursor returned by this query')}
   if(!Array.isArray(c)||c.length!==3||!Number.isSafeInteger(c[1])||c[1]<0)throw new InputError('opts.cursor','expected a cursor returned by this query')
   if(c[0]!==revision)throw new CalcError('stale_revision','the cursor belongs to another revision')
   if(c[2]!==signature)throw new InputError('opts.cursor','the cursor belongs to another query')
   offset=c[1]
  }
  this.offset=offset
 }
 wanted():boolean {
  const n=this.matched++
  if(n<this.offset)return false
  if(this.items.length<this.max)return true
  this.more=true;return false
 }
 get done():boolean{return this.more&&!this.count}
 result():QueryPage {
  return {revision:this.revision,clock:{...this.clock},truncated:this.more,...(this.more?{cursor:JSON.stringify([this.revision,this.offset+this.items.length,this.signature])}:{}),...(this.count?{total:this.matched}:{})}
 }
}

/** Native input order followed by spilled cells, once across overlapping ranges. Stable while the revision is unchanged. */
function* requested(s:EngineState,refs:Ref[]):Generator<[number,number,number]|undefined> {
 const earlier=(n:number,i:number,r:number,c:number)=>{for(let j=0;j<n;j++)if(refs[j].s===i&&hasCell(refs[j],r,c))return true;return false}
 for(const [n,x] of refs.entries()){
  for(const addr in s.input.sheets[x.s].cells){
   if(paced(1))yield undefined
   const {row,col}=parseAddress(addr)
   if(hasCell(x,row,col)&&!earlier(n,x.s,row,col))yield [x.s,row,col]
  }
  for(const spill of s.spills.on.values()){
   if(spill.s!==x.s||!meets(spill,x))continue
   for(let r=Math.max(spill.r1,x.r1);r<=Math.min(spill.r2,x.r2);r++)for(let c=Math.max(spill.c1,x.c1);c<=Math.min(spill.c2,x.c2);c++){
    if(paced(1))yield undefined
    if(!(formatAddress(r,c) in s.input.sheets[x.s].cells)&&!earlier(n,x.s,r,c))yield [x.s,r,c]
   }
  }
 }
}

function* projected(s:EngineState,i:number,r:number,c:number,fields:CellReadField[]):Steps<ReadCell>{
 const addr=formatAddress(r,c),out:ReadCell={sheet:s.sheetNames[i],cell:addr}
 const view=fields.includes('text')?(yield* cellsIn(s,i,[{r1:r,c1:c,r2:r,c2:c}],true))[addr]:cellResult(s,i,addr,null,undefined)
 const input=s.input.sheets[i].cells[addr]
 for(const f of fields){
  if(f==='issue')out.issue=issueDetail(s,i,r,c)
  else if(f==='style'&&input?.style!==undefined)out.style=input.style
  else if(f==='note'&&input?.note)out.note={...input.note}
  else if(f==='link'&&input?.link)out.link={...input.link}
  else if(view[f as keyof typeof view]!==undefined)(out as unknown as Record<string,unknown>)[f]=view[f as keyof typeof view]
 }
 return out
}

export function* readQuery(s:EngineState,revision:number,clock:Clock,given:RangeRef[],eo:ReadCellsOptions|FindCellsOptions|undefined,find?:string):Steps<ReadCellsResult>{
 const o=checkBase<FindCellsOptions>(eo,revision,[...PAGE,'fields',...(find===undefined?[]:['in','caseSensitive'])]),refs=refsOf(s,given)
 const fields=o.fields===undefined?['value','text','formula'] as CellReadField[]:array(o.fields,'opts.fields','an array of cell field names') as CellReadField[]
 if(fields.some(f=>!FIELDS.includes(f)))throw new InputError('opts.fields','unknown cell field')
 const where=o.in??'text'
 if(find!==undefined&&!['text','formula','value'].includes(where))throw new InputError('opts.in','expected text, formula or value')
 if(o.caseSensitive!==undefined)checkBoolean(o.caseSensitive,'opts.caseSensitive')
 const needle=find===undefined?undefined:checkString(find,'query'),lower=o.caseSensitive!==true
 const page=new Page<ReadCell>(revision,clock,JSON.stringify([refs,fields,needle,where,lower]),o)
 for(const at of requested(s,refs)){
  if(!at){yield;continue}
  if(needle!==undefined){
   const v=yield* projected(s,...at,[where])
   const value=v[where],text=typeof value==='object'&&value!==null?JSON.stringify(value):String(value??'')
   if(!needle||!(lower?text.toLowerCase().includes(needle.toLowerCase()):text.includes(needle)))continue
  }
  if(page.wanted())page.items.push(yield* projected(s,...at,fields))
  if(page.done)break
 }
 return {...page.result(),cells:page.items}
}

export function* dependentsQuery(s:EngineState,revision:number,clock:Clock,given:RangeRef[],eo?:DependentsOptions):Steps<DependentsResult>{
 const o=checkBase<DependentsOptions>(eo,revision,[...PAGE,'maxReferences']),refs=refsOf(s,given)
 const max=o.maxReferences??64;checkNumber(max,'opts.maxReferences',1,1000,true)
 const page=new Page<DependentCell>(revision,clock,JSON.stringify(['dependents',refs,max]),o)
 for(const f of s.formulas){
  if(paced(1))yield
  if(!f)continue
  const matches=referencesOf(s,f).filter(g=>refs.some(x=>x.s===g.s&&meets(x,g)))
  if(!matches.length)continue
  if(page.wanted())page.items.push({sheet:s.sheetNames[f.s],cell:formatAddress(f.r,f.c),references:matches.slice(0,max).map(g=>({sheet:s.sheetNames[g.s],range:rectText(g)})),...(matches.length>max?{referencesTruncated:true}: {})})
  if(page.done)break
 }
 return {...page.result(),dependents:page.items}
}

export function* summaryQuery(s:EngineState,revision:number,clock:Clock,given:RangeRef[],eo?:PageOptions):Steps<RangeSummary>{
 checkBase(eo,revision)
 const out:RangeSummary={revision,clock:{...clock},cells:0,numbers:0,errors:0,sum:0,min:null,max:null},sum=new CompensatedSum()
 for(const at of requested(s,refsOf(s,given))){
  if(!at){yield;continue}
  const v=outValue(toCellValue(valueAt(s,...at)))
  if(v===null||v==='')continue
  out.cells++
  if(typeof v==='number'){out.numbers++;sum.add(v);out.min=out.min===null?v:Math.min(out.min,v);out.max=out.max===null?v:Math.max(out.max,v)}
  else if(typeof v==='object')out.errors++
 }
 out.sum=Number.isFinite(sum.value)?sum.value:{error:'NUM'}
 return out
}
