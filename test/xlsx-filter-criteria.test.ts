/** Real Excel round trips of stored filter metadata; prospective expectations remain separate from recorded misses. */
import { existsSync, readFileSync } from 'node:fs'
import { expect, it } from 'vitest'
import { isDeepStrictEqual } from 'node:util'
import { readXlsx } from '../src/xlsx.js'
import { filterState } from '../bench/coverage/links/filter-criteria.js'
import type { StoredLinks } from '../bench/coverage/links.js'

/** Empty default columns were removed by Excel; table sorts nested inside autoFilter were discarded by the first writer. */
const MISSED:Record<string,string[]> = {
  'filter-criteria': [
    'write-sheet-long-value','keep-sheet-long-value','write-sheet-table-over','write-table-long-value','keep-table-long-value',
    'write-table-moved','write-table-pasted','write-table-row-above','write-table-row-deleted','write-table-row-inside','write-table-sheet-copied',
    'write-table-sort-case','write-table-sort-column-before','write-table-sort-desc','write-table-sort-fill','write-table-sort-filter',
    'write-table-sort-font','write-table-sort-icon','write-table-sort-list','write-table-sort-one-of-two-deleted','write-table-sort-pinyin',
    'write-table-sort-stars','write-table-sort-stroke','write-table-sort-two','write-table-sort-value',
  ],
  'filter-criteria-held': [],
}

for(const name of Object.keys(MISSED)){
  const url=new URL(`./excel/win/links/${name}.json`,import.meta.url)
  const set=existsSync(url)?JSON.parse(readFileSync(url,'utf8')) as StoredLinks:null
  it.skipIf(!set)(`${name}: Excel retains the predicted criteria, sort and hidden rows`,async()=>{
    const expected=JSON.parse(readFileSync(new URL(`../bench/coverage/links/${name}.predicted.json`,import.meta.url),'utf8'))
    const missed:string[]=[]
    for(const c of set!.cases){
      if(!!c.refused!==!!set!.rejected?.[c.id])missed.push(c.id)
      if(!set!.answers[c.id])continue
      const file=new URL(`./excel/win/links/${name}/${c.id}.xlsx`,import.meta.url)
      const got=filterState((await readXlsx(readFileSync(file))).book)
      if(!isDeepStrictEqual(got,expected[c.id.replace(/^(write|keep)-/, '')]))missed.push(c.id)
    }
    expect(missed).toEqual(MISSED[name])
  })
}
