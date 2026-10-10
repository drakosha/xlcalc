/**
 * Threads of comments put into a written file by hand, as Excel 365 writes them: neither Excel's AppleScript nor Excel
 * 2024's COM makes them (the latter answers that its object model has none). A sheet's threads go to a threaded comments
 * part, their authors to the book's persons part, and each thread's cell gets Excel's placeholder note for readers
 * without threads (author `tc={id}`), which the writer writes with its VML drawing and which is tied to its thread by its
 * uid.
 */
import type { BookInput } from '../../../src/contract.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { escapeAttr, escapeText } from '../../../src/xlsx/xml/text.js'
import { openArchive } from '../../../src/xlsx/zip/read.js'
import { zip } from '../../../src/xlsx/zip/write.js'
import type { ThreadSpec } from './cases.js'

const NS = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments'
const REL = 'http://schemas.microsoft.com/office/2017/10/relationships'
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

/** A GUID of fixed digits for the n-th thing of a kind, so the file is the same at every run. */
const guid = (kind: number, n: number) => `{${kind.toString(16).padStart(8, '0').toUpperCase()}-0000-4000-8000-${n.toString(16).padStart(12, '0').toUpperCase()}}`

/** The text of Excel's placeholder note for a thread. */
function placeholder(comments: ThreadSpec['comments']): string {
  const lines = comments.map((c, i) => `${i ? 'Reply' : 'Comment'}:\n    ${c.text}`)
  return `[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\n${lines.join('\n')}`
}

/** The book written by `writeXlsx` with the threads in its file. */
export async function threadedFile(book: BookInput, threads: ThreadSpec[]): Promise<Uint8Array> {
  const people = [...new Set(threads.flatMap((t) => t.comments.map((c) => c.author)))]
  const sheets = book.sheets.map((s) => ({ ...s, cells: { ...s.cells } }))
  threads.forEach((t, k) => {
    const s = sheets.find((x) => x.name === t.sheet)!
    s.cells[t.cell] = { ...s.cells[t.cell], note: { author: `tc=${guid(2, k + 1)}`, text: placeholder(t.comments) } }
  })
  const archive = openArchive((await writeXlsx({ ...book, sheets })).data)
  const files = new Map<string, string | Uint8Array>()
  for (const name of archive.names) files.set(name, (await archive.bytes(name))!)
  const text = async (name: string) => (await archive.text(name))!
  // a part of threads per sheet that has some, its relationship after the sheet's own
  let n = 0
  for (const [i, s] of book.sheets.entries()) {
    const own = threads.map((t, k) => ({ t, k })).filter(({ t }) => t.sheet === s.name)
    if (!own.length) continue
    const part = `threadedComment${++n}.xml`
    const items = own.flatMap(({ t, k }) =>
      t.comments.map((c, j) => `<threadedComment ref="${t.cell}" dT="2026-10-10T10:${String(j).padStart(2, '0')}:00.00" personId="${guid(1, people.indexOf(c.author) + 1)}" id="${j ? guid(3, (k + 1) * 100 + j) : guid(2, k + 1)}"${j ? ` parentId="${guid(2, k + 1)}"` : ''}><text>${escapeText(c.text)}</text></threadedComment>`),
    )
    files.set(`xl/threadedComments/${part}`, `${HEAD}<ThreadedComments xmlns="${NS}" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${items.join('')}</ThreadedComments>`)
    // Excel ties a placeholder note to its thread by the note's uid, the thread's id
    const comments = `xl/comments${i + 1}.xml`
    const authors = [...(await text(comments)).matchAll(/<author>([^<]*)<\/author>|<author\/>/g)].map((m) => m[1] ?? '')
    const tied = (await text(comments))
      .replace('<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', '$& xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"')
      .replace(/<comment ref="([A-Z]+\d+)" authorId="(\d+)"/g, (m, _ref: string, id: string) => (authors[Number(id)].startsWith('tc=') ? `${m} xr:uid="${authors[Number(id)].slice(3)}"` : m))
    files.set(comments, tied)
    const rels = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`
    files.set(rels, (await text(rels)).replace('</Relationships>', `<Relationship Id="rIdT1" Type="${REL}/threadedComment" Target="../threadedComments/${part}"/></Relationships>`))
  }
  const persons = people.map((p, i) => `<person displayName="${escapeAttr(p)}" id="${guid(1, i + 1)}" userId="${escapeAttr(p)}" providerId="None"/>`)
  files.set('xl/persons/person.xml', `${HEAD}<personList xmlns="${NS}" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${persons.join('')}</personList>`)
  files.set('xl/_rels/workbook.xml.rels', (await text('xl/_rels/workbook.xml.rels')).replace('</Relationships>', `<Relationship Id="rIdP1" Type="${REL}/person" Target="persons/person.xml"/></Relationships>`))
  const types = [...Array(n)].map((_, k) => `<Override PartName="/xl/threadedComments/threadedComment${k + 1}.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>`).join('')
  files.set('[Content_Types].xml', (await text('[Content_Types].xml')).replace('</Types>', `${types}<Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/></Types>`))
  const enc = new TextEncoder()
  return zip([...files].map(([name, data]) => ({ name, data: typeof data === 'string' ? enc.encode(data) : data })))
}
