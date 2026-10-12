import { fork, execFileSync } from 'node:child_process'
import { readFileSync, mkdirSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { cpus, totalmem } from 'node:os'

const args=process.argv.slice(2)
const option=(name,fallback)=>args.includes(name)?args[args.indexOf(name)+1]:fallback
const before=resolve(option('--before','')),after=resolve(option('--after','.'))
if(!args.includes('--before'))throw new Error('Pass --before with a compiled baseline checkout')
const out=resolve(option('--out','bench/hyperformula/results'))
const cases=['unchanged','churn','cache-only','fanout','prefix-statistics','business-34000','lookup']
const selected=option('--case',cases.join(',')).split(',')
if(selected.some(name=>!cases.includes(name)))throw new Error('Unknown memory workload')
const revision=root=>{try{return execFileSync('git',['rev-parse','HEAD'],{cwd:root,encoding:'utf8',stdio:['ignore','pipe','pipe']}).trim()}catch{return readFileSync(resolve(root,'revision.txt'),'utf8').trim()}}
const report={before:revision(before),after:revision(after),node:process.version,machine:{cpu:cpus()[0].model,memoryGiB:totalmem()/2**30},protocol:{pairs:5,heapLimitMiB:512,deadlineMs:30000,history:'disabled',retainedHeap:'shared input excluded; no same-book warmup before baseline; held session pinned, two event-loop turns and two GCs after measured work',rss:'whole-process maximum before full-compute verification, not peak JS heap'},samples:[]}
mkdirSync(out,{recursive:true})
const save=()=>writeFileSync(resolve(out,'low-memory-comparison.json'),JSON.stringify(report,null,2)+'\n')
function run(root,mode,name){
 return new Promise((done,reject)=>{
  const child=fork(new URL('./memory-child.mjs',import.meta.url),[root,mode,name],{silent:true,execArgv:['--expose-gc','--max-old-space-size=512']})
  let stdout='',stderr=''
  const timer=setTimeout(()=>child.kill('SIGKILL'),30000)
  child.stdout.on('data',x=>stdout+=x)
  child.stderr.on('data',x=>stderr=(stderr+x).slice(-3000))
  child.on('error',error=>{clearTimeout(timer);reject(error)})
  child.on('exit',(code,signal)=>{clearTimeout(timer);if(code!==0)reject(new Error(`${name}/${mode}: ${code??signal}: ${stderr}`));else{try{done(JSON.parse(stdout))}catch(error){reject(error)}}})
 })
}
const median=xs=>xs.sort((a,b)=>a-b)[Math.floor(xs.length/2)]
for(const name of selected)for(const mode of(name==='cache-only'?['closures']:['closures','eval'])){
 for(let sample=0;sample<5;sample++){
  const roots=sample%2?[after,before]:[before,after]
  for(const root of roots){const measured=await run(root,mode,name);report.samples.push({...measured,sample,variant:root===before?'before':'after'});save()}
 }
 const samples=report.samples.filter(x=>x.name===name&&x.mode===mode)
 console.log(JSON.stringify({name,mode,medians:Object.fromEntries(['before','after'].map(v=>[v,Object.fromEntries(['heapMiB','maxRssMiB',...(name==='cache-only'?[]:['openMs','editMs'])].map(k=>[k,median(samples.filter(x=>x.variant===v).map(x=>x[k]))]))]))}))
}
report.complete=true
save()
