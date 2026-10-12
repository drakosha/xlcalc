import { pathToFileURL } from 'node:url'
import { setImmediate as tick } from 'node:timers/promises'
const [root, mode, name] = process.argv.slice(2)
if (!root || !['closures','eval'].includes(mode)) throw new Error('Expected a compiled checkout path and closures or eval')
if (typeof global.gc !== 'function') throw new Error('Run with node --expose-gc')
const {createSession,compute} = await import(pathToFileURL(`${root}/dist/index.js`))
const {workload}=await import('./workloads.mjs')
const CLOCK={now:'2026-10-04T12:00:00Z',seed:7}
if (name === 'cache-only') {
 const { RangeCache } = await import(pathToFileURL(`${root}/dist/range-cache.js`))
 await tick(); global.gc(); const before=process.memoryUsage().heapUsed
 const cache=new RangeCache()
 globalThis.heldCache=cache
 for(let i=0;i<50000;i++)cache.set(`0:${i}:0:${i}:0`,i)
 await tick(); await tick(); global.gc(); global.gc()
 const after=process.memoryUsage()
 if(cache.size!==50000 || cache.get('0:49999:0:49999:0')!==49999)throw new Error('Cache values differ')
 console.log(JSON.stringify({root,mode,name,entries:cache.size,heapMiB:(after.heapUsed-before)/2**20,rssMiB:after.rss/2**20,maxRssMiB:process.resourceUsage().maxRSS/1024}))
} else {
let book,runWork,edits=0
if(name==='churn'){
 book={sheets:[{name:'S',cells:{A1:{value:1},B1:{formula:'A1+1'},C1:{formula:'B1*2'}}}]}
 runWork=s=>{for(let i=0;i<5000;i++)s.setCells([{sheet:'S',cell:'B1',input:{formula:`A1+${i%2+1}`}}])};edits=5000
}else if(name==='unchanged'){
 const cells={A1:{value:1.1},A2:{formula:'INT(A1)'}}
 for(let r=3;r<=10000;r++)cells[`A${r}`]={formula:`A${r-1}+1`}
 book={sheets:[{name:'S',cells}]}
 runWork=s=>{for(let i=0;i<50;i++)s.setCells([{sheet:'S',cell:'A1',input:{value:i%2?1.1:1.9}}])};edits=50
}else{book=workload(name).book;runWork=()=>{}}
await tick();global.gc();const before=process.memoryUsage().heapUsed
const start=performance.now();let s=createSession(book,{clock:CLOCK,execution:mode,display:false,undo:{depth:0}});const openMs=performance.now()-start
globalThis.heldSession=s;const t=performance.now();runWork(s);const editMs=performance.now()-t
await tick();await tick();global.gc();global.gc();const after=process.memoryUsage()
const out={root,mode,name,openMs,editMs,edits,heapMiB:(after.heapUsed-before)/2**20,rssMiB:after.rss/2**20,maxRssMiB:process.resourceUsage().maxRSS/1024}
const full=compute(s.getBook(),{clock:CLOCK,execution:mode,display:false})
out.hash=s.getHash();out.fullHash=full.hash
if(out.hash!==out.fullHash)throw new Error(JSON.stringify(out))
s.close();globalThis.heldSession=null;s=null;await tick();global.gc()
console.log(JSON.stringify(out))

}
