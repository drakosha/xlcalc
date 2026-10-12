// Five alternating process trials; compare candidate code generation against the frozen release and prior candidate.
import { fork } from 'node:child_process'
import { mkdirSync,writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
const args=process.argv.slice(2),opt=(n,d)=>args.includes(n)?args[args.indexOf(n)+1]:d
const roots=opt('--roots','').split(',').map(x=>resolve(x))
if(roots.length<2)throw new Error('Pass --roots with compiled checkouts')
const cases=opt('--case','fanout').split(','),out=resolve(opt('--out','bench/hyperformula/results-row-sharing'))
const report={node:process.version,roots,protocol:{trials:5,rotation:'root order rotates each trial',child:'same child.mjs, 5 opens and 20 edits, full verification, 1024 MiB cap and phase deadlines'},samples:[]}
mkdirSync(out,{recursive:true})
const save=()=>writeFileSync(resolve(out,'comparison.json'),JSON.stringify(report,null,2)+'\n')
function run(root,id,engine){return new Promise((done,reject)=>{
 const p=fork(new URL('./child.mjs',import.meta.url),[id,engine],{silent:true,execArgv:['--expose-gc','--max-old-space-size=1024'],env:{...process.env,XLCALC_BENCH_ROOT:root}})
 let result,timer,stderr='';const arm=()=>{clearTimeout(timer);timer=setTimeout(()=>p.kill('SIGKILL'),30000)};arm()
 p.stdout.resume();p.stderr.on('data',x=>stderr=(stderr+x).slice(-2000))
 p.on('message',m=>{if(m.type==='phase')arm();if(m.type==='result'){result=m.result;clearTimeout(timer)}})
 p.on('error',reject);p.on('exit',(code)=>{clearTimeout(timer);if(result?.status==='ok')done(result);else reject(new Error(`${id}/${engine}: ${code}: ${JSON.stringify(result??stderr)}`))})
})}
const median=xs=>xs.sort((a,b)=>a-b)[Math.floor(xs.length/2)]
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
