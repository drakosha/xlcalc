/** Stored criteria of Excel's synthetic books through fresh and preserving writes, predicted before Excel opens them. */
import { readFileSync, readdirSync } from 'node:fs'
import type { BookInput } from '../../../src/contract.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'
import { openKept, writeKept } from '../../../src/xlsx-keep.js'
import { lines } from '../../../src/layout/spans.js'
import type { LinksCase, LinksCases } from './cases.js'

const ROOT = new URL('../../../test/excel/win/filter-criteria-made/', import.meta.url)
export const SOURCES = readdirSync(ROOT).filter(n => n.endsWith('.xlsx')).sort()

/** Only the state this set preserves: stored criteria, sorts and row visibility, without reapplying a filter. */
export function filterState(book: BookInput): unknown {
  return book.sheets.filter(s => !/^(Probe|WT|WN|WX)$/.test(s.name)).map(s => ({
    name: s.name,
    ...(s.autoFilter ? { autoFilter: s.autoFilter } : {}),
    ...(s.sort ? { sort: s.sort } : {}),
    filtered: !!s.filtered,
    hidden: lines(s.rows,'r').filter(([, ,v]) => v.hidden).flatMap(([a,b]) => Array.from({length:b-a+1},(_,i)=>a+i+1)),
    tables: (s.tables ?? []).map(t => ({name:t.name,range:t.range,...(t.filter===undefined?{}:{filter:t.filter})})),
  }))
}

const file = (name:string) => new Uint8Array(readFileSync(new URL(name,ROOT)))

export function buildLinksCases(): LinksCases {
  const cases: LinksCase[] = []
  for (const name of SOURCES) for (const mode of ['write','keep'] as const) {
    cases.push({
      id: `${mode}-${name.slice(0,-5)}`,
      about: `${mode==='write'?'written anew':'written over its original'}: ${name.slice(0,-5)}`,
      sheets:[{name:'Probe',cells:{A1:{value:1}}}],linked:{},links:[],
      read:[{sheet:'Probe',range:'A1'}],predicted:{'Probe!A1':1},written:true,
      file: async mirror => {
        const original=file(name)
        const read=mode==='keep'?await openKept(original):await readXlsx(original)
        const book={...read.book,sheets:[...read.book.sheets,...mirror.sheets]}
        return (mode==='keep'?await writeKept(original,{book}):await writeXlsx({book},{values:'omit'})).data
      },
    })
  }
  return { cases }
}
