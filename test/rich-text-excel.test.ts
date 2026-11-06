/** Character fonts recorded by Excel for fresh and preserving rich-text writes. */
import {readFileSync} from 'node:fs'
import {expect,it} from 'vitest'
import {createSession,type BookInput,type FontStyle} from '../src/index.js'
import {readXlsx,writeXlsx} from '../src/xlsx.js'
import {writeKept} from '../src/xlsx-keep.js'

const root=new URL('./excel/win/rich-text/',import.meta.url)
const fixture=JSON.parse(readFileSync(new URL('fonts.json',root),'utf8'))
const sources=fixture.source as {book:BookInput;changed:BookInput}
const data=(name:string)=>readFileSync(new URL(name,root))

function characterFonts(book:BookInput,addr:string){
  const sheet=book.sheets[0],cell=sheet.cells[addr]
  const row=addr.replace(/[^0-9]/g,''),col=addr.replace(/[^A-Z]/g,'')
  const style=cell.style??sheet.rows?.[row]?.style??sheet.columns?.[col]?.style
  const base=style===undefined?{}:book.styles?.[style]?.font??{}
  const norm=(font:FontStyle)=>({name:font.name??'Calibri',size:font.size??11,bold:!!font.bold,italic:!!font.italic,strike:!!font.strike,underline:font.underline==='double'?-4119:font.underline==='single'?2:-4142,color:font.color??'#000000',auto:font.color===undefined})
  return Array.from({length:String(cell.value).length},(_,i)=>{
    const run=cell.richText?.findLast(r=>r.start<=i)
    return norm(run?.font??base)
  })
}

function check(book:BookInput,expected:Record<string,{text:string;chars:unknown[]}>) {
  for(const [addr,cell] of Object.entries(expected)){
    expect(book.sheets[0].cells[addr].value,addr).toBe(cell.text)
    expect(characterFonts(book,addr),addr).toEqual(cell.chars)
  }
}

it.each(['written','kept'] as const)('reads %s character fonts before and after Excel edits',async mode=>{
  check((await readXlsx(data(`${mode}.xlsx`))).book,fixture.answers[mode].before)
  check((await readXlsx(data(`saved-${mode}.xlsx`))).book,fixture.answers[mode].after)
})

it('writes the character fonts Excel accepted through both writers',async()=>{
  const fresh=await writeXlsx({book:sources.book},{values:'omit'})
  check((await readXlsx(fresh.data)).book,fixture.answers.written.before)
  const kept=await writeKept(data('written.xlsx'),{book:sources.changed})
  check((await readXlsx(kept.data)).book,fixture.answers.kept.before)
})

it('applies whole-cell font and fill changes like Excel without losing other run properties',()=>{
  const session=createSession(sources.book)
  session.patchCells([{sheet:'Text',cell:'A3',set:{style:{font:{bold:true}}}},{sheet:'Text',cell:'A4',set:{style:{font:{color:'#0000FF'}}}},{sheet:'Text',cell:'A5',set:{style:{fill:{color:'#FFFF00'}}}}])
  const book=session.getBook()
  for(const addr of ['A3','A4','A5'])check(book,{[addr]:fixture.answers.written.after[addr]})
  session.close()
})
