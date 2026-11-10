/** Real Excel round trips of stored filter metadata; prospective expectations remain separate from recorded misses. */
import { existsSync, readFileSync } from 'node:fs'
import { expect, it } from 'vitest'
import { isDeepStrictEqual } from 'node:util'
import { readXlsx } from '../src/xlsx.js'
import { filterState } from '../bench/coverage/links/filter-criteria.js'
import type { StoredLinks } from '../bench/coverage/links.js'

/**
 * The AutoFilter's sort read a second time as the sheet's, by the reader the predictions were made with: the files
 * Excel kept unchanged hold it once, inside autoFilter, as the reader now reads it.
 */
const SORT_READ_TWICE = [
  'keep-sheet-moved','keep-sheet-row-above','keep-sheet-row-deleted','keep-sheet-row-inside','keep-sheet-sheet-copied','keep-sheet-sort-case',
  'keep-sheet-sort-column-before','keep-sheet-sort-desc','keep-sheet-sort-fill','keep-sheet-sort-filter','keep-sheet-sort-font','keep-sheet-sort-icon',
  'keep-sheet-sort-list','keep-sheet-sort-one-of-two-deleted','keep-sheet-sort-pinyin','keep-sheet-sort-stars','keep-sheet-sort-stroke','keep-sheet-sort-two',
  'keep-sheet-sort-value',
]

/** Empty default columns were removed by Excel; table sorts nested inside autoFilter were discarded by the first writer. */
const MISSED:Record<string,string[]> = {
  'filter-criteria': [
    'write-sheet-long-value','keep-sheet-long-value',...SORT_READ_TWICE,'write-sheet-table-over','write-table-long-value','keep-table-long-value',
    'write-table-moved','write-table-pasted','write-table-row-above','write-table-row-deleted','write-table-row-inside','write-table-sheet-copied',
    'write-table-sort-case','write-table-sort-column-before','write-table-sort-desc','write-table-sort-fill','write-table-sort-filter',
    'write-table-sort-font','write-table-sort-icon','write-table-sort-list','write-table-sort-one-of-two-deleted','write-table-sort-pinyin',
    'write-table-sort-stars','write-table-sort-stroke','write-table-sort-two','write-table-sort-value',
  ],
  'filter-criteria-held': SORT_READ_TWICE,
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
