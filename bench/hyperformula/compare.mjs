// Five alternating process trials; compare candidate code generation against the frozen release and prior candidate.
import { mkdirSync,writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { runChild } from '../run-child.mjs'
import { median } from '../stats.mjs'
const args=process.argv.slice(2),opt=(n,d)=>args.includes(n)?args[args.indexOf(n)+1]:d
const roots=opt('--roots','').split(',').map(x=>resolve(x))
if(roots.length<2)throw new Error('Pass --roots with compiled checkouts')
const cases=opt('--case','fanout').split(','),out=resolve(opt('--out','bench/hyperformula/results-row-sharing'))
const report={node:process.version,roots,protocol:{trials:5,rotation:'root order rotates each trial',child:'same child.mjs, 5 opens and 20 edits, full verification, 1024 MiB cap and phase deadlines'},samples:[]}
mkdirSync(out,{recursive:true})
const save=()=>writeFileSync(resolve(out,'comparison.json'),JSON.stringify(report,null,2)+'\n')
const run=(root,id,engine)=>runChild(new URL('./child.mjs',import.meta.url),[id,engine],{execArgv:['--expose-gc','--max-old-space-size=1024'],env:{...process.env,XLCALC_BENCH_ROOT:root},phaseLimitMs:30000,stderrTail:2000}).then(f=>{
 if(f.error)throw f.error
 if(f.received?.status==='ok')return f.received
 throw new Error(`${id}/${engine}: ${f.code}: ${JSON.stringify(f.received??f.stderr)}`)})
for(const id of cases)for(const engine of ['xl-closures','xl-eval']){
 for(let trial=0;trial<5;trial++)for(let k=0;k<roots.length;k++){
  const root=roots[(trial+k)%roots.length],result=await run(root,id,engine)
  report.samples.push({root,id,engine,trial,...result});save()
 }
 console.log(JSON.stringify({id,engine,medians:roots.map(root=>{
  const s=report.samples.filter(x=>x.root===root&&x.id===id&&x.engine===engine)
  return {root,open:median(s.map(x=>x.coldMedianMs)),edits:Object.fromEntries(Object.keys(s[0].edits).map(key=>[key,median(s.map(x=>x.edits[key].medianMs))]))}
 })}))
}
report.complete=true;save()
