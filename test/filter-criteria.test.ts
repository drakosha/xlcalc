import { readFile } from 'node:fs/promises'
import { readdirSync } from 'node:fs'
import { isDeepStrictEqual } from 'node:util'
import { expect, it } from 'vitest'
import { compute, editStructure, type BookInput, type FilterCriteria } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { pack, parts, partText } from './xlsx-fixtures.js'

const criteria: FilterCriteria = {
  columns: [
    { column: 0, hiddenButton: true, filter: { type: 'values', values: ['Apple', 'a*'], blank: true, dates: [{ level: 'month', year: 2026, month: 3 }], calendar: 'hijri' } },
    { column: 1, filter: { type: 'custom', and: true, conditions: [{ operator: 'greaterThan', value: '10' }, { operator: 'lessThan', value: '20' }] } },
    { column: 2, filter: { type: 'top', value: 10, percent: true, bottom: true, filterValue: 5 } },
    { column: 3, filter: { type: 'dynamic', kind: 'thisWeek', value: 46000, maxValue: 46007 } },
    { column: 4, filter: { type: 'color', font: true, style: { font: { color: '#FF0000' } } } },
    { column: 5, filter: { type: 'icon', set: '3Arrows', icon: 2 } },
  ],
  sort: { range: 'A2:F3', caseSensitive: true, conditions: [{ range: 'B2:B3', descending: true, customList: 'B,A' }] },
}
const book = (): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 'Head' }, A2: { value: 4 }, A3: { value: 9 } }, rows: { 3: { hidden: true } }, filtered: true, autoFilter: { range: 'A1:F3', criteria } }] })

it('holds every criterion and sort through writeXlsx without applying them', async () => {
  const source=book()
  const {data,issues}=await writeXlsx({book:source},{values:'omit'})
  expect(issues).toEqual([])
  const result=await readXlsx(data)
  expect(result.book.sheets[0].autoFilter).toEqual(source.sheets[0].autoFilter)
  expect(result.book.sheets[0].rows).toEqual(source.sheets[0].rows)
  expect(result.issues.filter(x=>x.code==='autoFilter')).toEqual([])
})

it('reads criteria already written by Excel',async()=>{
  const result=await readXlsx(await readFile(new URL('./excel/win/filter-criteria-made/sheet-values.xlsx',import.meta.url)))
  expect(result.book.sheets[0].autoFilter?.criteria?.columns?.length).toBeGreaterThan(0)
  expect(result.issues.filter(x=>x.code==='autoFilter')).toEqual([])
})

it('keeps criteria and sort outside the calculation hash',()=>{
  const a=book();const b=book();delete b.sheets[0].autoFilter!.criteria
  b.sheets[0].sort={range:'A2:F3',conditions:[{range:'A2:A3'}]}
  const x=compute(a);const y=compute(b)
  expect(x.inputHash).toBe(y.inputHash)
  expect(x.hash).toBe(y.hash)
})

it('represents a table with its filter buttons disabled',async()=>{
  const source:BookInput={sheets:[{name:'S',cells:{A1:{value:'Item'},A2:{value:1}},tables:[{name:'T',range:'A1:A2',columns:[{name:'Item'}],filter:false}]}]}
  const {data}=await writeXlsx({book:source})
  expect(await partText(data,'xl/tables/table1.xml')).not.toContain('<autoFilter')
  expect((await readXlsx(data)).book.sheets[0].tables?.[0].filter).toBe(false)
})


it('writes edited criteria through writeKept and keeps an untouched file byte for byte',async()=>{
  const {data}=await writeXlsx({book:book()},{values:'omit'})
  const read=await openKept(data)
  expect((await writeKept(data,{book:read.book})).data).toEqual(data)
  const changed=structuredClone(read.book)
  changed.sheets[0].autoFilter!.criteria!.columns![0].filter={type:'values',values:['Changed']}
  const result=await readXlsx((await writeKept(data,{book:changed})).data)
  expect(result.book.sheets[0].autoFilter?.criteria).toEqual(changed.sheets[0].autoFilter?.criteria)
})

it('moves criterion columns and sort references with inserted columns',()=>{
  const result=editStructure(book(),{op:'insertColumns',sheet:'S',column:'B',count:2})
  expect(result.ok).toBe(true)
  if(!result.ok)return
  const f=result.book.sheets[0].autoFilter!
  expect(f.range).toBe('A1:H3')
  expect(f.criteria?.columns?.map(c=>c.column)).toEqual([0,3,4,5,6,7])
  expect(f.criteria?.sort?.conditions?.[0].range).toBe('D2:D3')
})


it('uses the extended customFilter form for text longer than 255 characters',async()=>{
  const b=book();b.sheets[0].autoFilter!.criteria={columns:[{column:0,filter:{type:'custom',conditions:[{value:'X'.repeat(300)}]}}]}
  const {data}=await writeXlsx({book:b},{values:'omit'})
  expect(await partText(data,'xl/worksheets/sheet1.xml')).toContain('<x14:customFilters>')
  expect((await readXlsx(data)).book.sheets[0].autoFilter).toEqual(b.sheets[0].autoFilter)
})

const made=new URL('./excel/win/filter-criteria-made/',import.meta.url)
const state=(b:BookInput)=>b.sheets.map(s=>({name:s.name,filter:s.autoFilter,sort:s.sort,filtered:!!s.filtered,rows:s.rows,tables:s.tables?.map(t=>({name:t.name,filter:t.filter}))}))
it.each(readdirSync(made).filter(n=>n.endsWith('.xlsx')))('keeps the filters and hidden rows of Excel book %s',async name=>{
  const original=await readFile(new URL(name,made));const read=await readXlsx(original)
  const written=await writeXlsx({book:read.book},{values:'omit'})
  const result=await readXlsx(written.data)
  expect(read.issues.filter(i=>i.code==='autoFilter')).toEqual([])
  expect(isDeepStrictEqual(state(read.book),state(result.book))).toBe(true)
  const kept=await openKept(original)
  expect(Buffer.compare(Buffer.from((await writeKept(original,{book:kept.book})).data),original)).toBe(0)
})


it('uses x14 list values for long selections, not only for custom conditions',async()=>{
  const b=book();b.sheets[0].autoFilter!.criteria={columns:[{column:0,filter:{type:'values',values:['X'.repeat(300),'short']}}]}
  const {data}=await writeXlsx({book:b},{values:'omit'})
  expect(await partText(data,'xl/worksheets/sheet1.xml')).toContain('<x14:filter ')
  expect((await readXlsx(data)).book.sheets[0].autoFilter).toEqual(b.sheets[0].autoFilter)
})

it('resolves a supported Choice by its namespace rather than a fixed prefix',async()=>{
  const f=parts({sheets:[{name:'S',body:'<sheetData/><autoFilter ref="A1:A2"><filterColumn colId="0"><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:f="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><mc:Choice Requires="f"><filters><f:filter val="chosen"/></filters></mc:Choice><mc:Fallback><filters><filter val="fallback"/></filters></mc:Fallback></mc:AlternateContent></filterColumn></autoFilter>'}]})
  const result=await readXlsx(await pack(f))
  expect(result.book.sheets[0].autoFilter?.criteria?.columns?.[0].filter).toEqual({type:'values',values:['chosen']})
})


it('writes the table sort after its autoFilter, where Excel retains it',async()=>{
  const b=book();delete b.sheets[0].autoFilter
  b.sheets[0].tables=[{name:'T',range:'A1:F3',columns:Array.from({length:6},(_,i)=>({name:'Column'+i})),filter:structuredClone(criteria)}]
  for(let i=0;i<6;i++)b.sheets[0].cells[String.fromCharCode(65+i)+'1']={value:'Column'+i}
  const {data}=await writeXlsx({book:b},{values:'omit'})
  const xml=await partText(data,'xl/tables/table1.xml')
  expect(xml.indexOf('<sortState')).toBeGreaterThan(xml.indexOf('</autoFilter>'))
  expect((await readXlsx(data)).book.sheets[0].tables?.[0].filter).toEqual(criteria)
})
