// Check the component with this checkout without editing its source or installed xlcalc dependency.
import { mkdtempSync, copyFileSync, cpSync, mkdirSync, readdirSync, symlinkSync, existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { resolve, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { spawnSync } from 'node:child_process'
const args=process.argv.slice(2)
const source=resolve(args.includes('--component')?args[args.indexOf('--component')+1]:'../xlcalc-sheet')
const target=args.includes('--out')?resolve(args[args.indexOf('--out')+1]):mkdtempSync(join(tmpdir(),'xlcalc-component-'))
if(existsSync(join(target,'package.json')))throw new Error('Use an empty output directory')
const root=fileURLToPath(new URL('../../',import.meta.url))
cpSync(source,target,{recursive:true,filter:(path)=>!/(?:^|\/)(?:\.git|node_modules|dist|release|test-results|\.vite|\.cache)(?:\/|$)/.test(path)})
const modules=join(target,'node_modules');mkdirSync(modules,{recursive:true})
for(const entry of readdirSync(join(source,'node_modules'),{withFileTypes:true})){
 if(entry.name==='xlcalc'||entry.name.startsWith('.')&&entry.name!=='.bin')continue
 symlinkSync(join(source,'node_modules',entry.name),join(modules,entry.name))
}
const library=join(modules,'xlcalc');mkdirSync(library)
copyFileSync(join(root,'package.json'),join(library,'package.json'))
cpSync(join(root,'dist'),join(library,'dist'),{recursive:true})
copyFileSync(new URL('./component-adapter.test.ts.template',import.meta.url),join(target,'test/live-api-adapter.test.ts'))
for(const script of ['typecheck','build','test']){
 const result=spawnSync('npm',['run',script],{cwd:target,stdio:'inherit'})
 if(result.status!==0)process.exit(result.status??1)
}
console.log(`Component checks passed in ${target}`)
