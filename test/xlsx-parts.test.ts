/**
 * Notes, threads of comments, hyperlinks and frozen panes in xlsx files: what the writer writes reads back the same,
 * what the writer cannot hold is cut and reported, and the reader takes each form a file holds them in, reporting what
 * the book has no place for. Files Excel itself saved are checked in `sheet-parts-excel.test.ts`.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput, CellInput } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { codes, pack, parts, read } from './xlsx-fixtures.js'

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`

/** The text of a part of a written book. */
async function part(book: BookInput, name: string): Promise<string> {
  return (await openArchive((await writeXlsx({ book })).data).text(name))!
}

const roundTrip = async (book: BookInput) => {
  const w = await writeXlsx({ book })
  const r = await readXlsx(w.data)
  return { book: r.book, issues: [...w.issues, ...r.issues] }
}

describe('notes, hyperlinks and frozen panes written and read back', () => {
  it('every kind comes back as written, on sheets with tables and merged areas beside them', async () => {
    const cells: Record<string, CellInput> = {
      A1: { value: 'Qty', note: { author: 'Ann', text: 'Line 1\nLine 2 & <b>"x"</b>\t_x0041_' } },
      A2: { value: 1, note: { text: ' spaces at the ends ' }, link: { url: 'https://example.com/a?b=1&c=2', location: 'part', tooltip: 'a "tip" & <more>' } },
      A3: { value: 2, link: { location: "'T T'!B2" } },
      C5: { note: { author: 'Борис', text: 'Примечание' } },
      D6: { link: { url: 'mailto:ann@example.com' } },
      D7: { value: 'same', link: { url: 'mailto:ann@example.com', tooltip: 'again' } },
    }
    const book: BookInput = {
      sheets: [
        { name: 'S', cells, frozen: { rows: 1 }, merges: ['F1:G1'], tables: [{ name: 'Sales', range: 'A1:A3', columns: [{ name: 'Qty' }] }] },
        { name: 'T T', cells: { B2: { value: 'x', note: { author: 'Ann', text: 'on T T' } } }, frozen: { rows: 3, columns: 2 } },
        { name: 'U', cells: {}, frozen: { columns: 1 } },
      ],
    }
    const r = await roundTrip(book)
    expect(r.issues).toEqual([])
    expect(r.book.sheets.map((s) => s.frozen)).toEqual([{ rows: 1 }, { rows: 3, columns: 2 }, { columns: 1 }])
    expect(r.book.sheets[0].cells).toEqual(cells)
    expect(r.book.sheets[1].cells.B2).toEqual(book.sheets[1].cells.B2)
    // one relationship for an address used twice, after the table's
    const sheetRels = await part(book, 'xl/worksheets/_rels/sheet1.xml.rels')
    expect([...sheetRels.matchAll(/Id="(rId\d+)" Type="[^"]*\/(\w+)"/g)].map((m) => `${m[1]} ${m[2]}`)).toEqual(['rId1 table', 'rId2 comments', 'rId3 vmlDrawing', 'rId4 hyperlink', 'rId5 hyperlink'])
    const sheet = await part(book, 'xl/worksheets/sheet1.xml')
    // the schema's order: merged areas, hyperlinks, the legacy drawing, table parts
    expect(sheet.indexOf('<mergeCells')).toBeLessThan(sheet.indexOf('<hyperlinks>'))
    expect(sheet.indexOf('</hyperlinks>')).toBeLessThan(sheet.indexOf('<legacyDrawing'))
    expect(sheet.indexOf('<legacyDrawing')).toBeLessThan(sheet.indexOf('<tableParts'))
    expect(sheet).toContain('<sheetView workbookViewId="0" tabSelected="1"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>')
  })

  it("VML shape ids are unique through the book: a sheet's notes take blocks of 1,024 ids, the next sheet's start after them", async () => {
    const many: Record<string, CellInput> = {}
    for (let i = 0; i < 1100; i++) many[`A${i + 1}`] = { note: { text: String(i) } }
    const book: BookInput = { sheets: [{ name: 'S', cells: many }, { name: 'T', cells: { A1: { note: { text: 't' } } } }] }
    const [s, t] = [await part(book, 'xl/drawings/vmlDrawing1.vml'), await part(book, 'xl/drawings/vmlDrawing2.vml')]
    expect(s).toContain('<o:idmap v:ext="edit" data="1,2"/>')
    expect(s).toContain('id="_x0000_s1025"')
    expect(s).toContain('id="_x0000_s2124"')
    expect(t).toContain('<o:idmap v:ext="edit" data="3"/>')
    expect(t).toContain('id="_x0000_s3073"')
    expect(Object.keys((await roundTrip(book)).book.sheets[0].cells)).toHaveLength(1100)
  })

  it("what Excel does not take is cut and reported: a note over 32,767 characters, a tooltip over 255", async () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { note: { text: 'n'.repeat(40_000) } }, B1: { link: { url: 'https://example.com/', tooltip: 't'.repeat(300) } } } }] }
    const r = await roundTrip(book)
    expect(r.book.sheets[0].cells.A1.note!.text).toHaveLength(32_767)
    expect(r.book.sheets[0].cells.B1.link!.tooltip).toHaveLength(255)
    expect(codes(r.issues)).toEqual(['textLength B1', 'textLength A1'])
  })
})

/** A one-sheet book (`S`, `x` in A1 and 1 in B1 with `DATA`) with the sheet's relationships and more parts. */
const sheetWith = async (body: string, sheetRels: string, more: Record<string, string> = {}) => {
  const f = parts({ sheets: [{ name: 'S', body }], strings: ['x'] })
  f['xl/worksheets/_rels/sheet1.xml.rels'] = rels(sheetRels)
  return readXlsx(await pack({ ...f, ...more }))
}

const DATA = '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1</v></c></row></sheetData>'

describe('frozen and split panes read', () => {
  const view = (pane: string, attrs = '') => read({ sheets: [{ name: 'S', body: `<sheetViews><sheetView workbookViewId="0"${attrs}>${pane}</sheetView></sheetViews>${DATA}` }] })

  it('frozen panes by their rows and columns, frozen after a split too; split panes alone reported', async () => {
    expect((await view('<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>')).book.sheets[0].frozen).toEqual({ rows: 1, columns: 2 })
    expect((await view('<pane xSplit="1" ySplit="3" topLeftCell="B4" activePane="bottomRight" state="frozenSplit"/>')).book.sheets[0].frozen).toEqual({ rows: 3, columns: 1 })
    // the pane after the split scrolled down: the split stays where it is
    expect((await view('<pane ySplit="2" topLeftCell="A40" activePane="bottomLeft" state="frozen"/>')).book.sheets[0].frozen).toEqual({ rows: 2 })
    const split = await view('<pane xSplit="2000" ySplit="1500" topLeftCell="C5"/>')
    expect(split.book.sheets[0].frozen).toBeUndefined()
    expect(codes(split.issues)).toEqual(['frozenPanes'])
  })

  it('panes frozen over a scrolled sheet are read from its first row and column, and reported', async () => {
    const r = await view('<pane ySplit="2" topLeftCell="A7" activePane="bottomLeft" state="frozen"/>', ' topLeftCell="A5"')
    expect(r.book.sheets[0].frozen).toEqual({ rows: 6 })
    expect(codes(r.issues)).toEqual(['frozenPanes'])
  })
})

describe('hyperlinks read', () => {
  it('an address from the relationships, a place, a tooltip; a range of cells each; a missing relationship reported', async () => {
    const body = `${DATA}<hyperlinks><hyperlink ref="A1" r:id="rId1" tooltip="t &amp; u"/><hyperlink ref="B1" location="S!A1" display="ignored"/><hyperlink ref="C2:D3" r:id="rId2" location="x"/><hyperlink ref="E1" r:id="rId9"/><hyperlink ref="E2" r:id="rId9" location="kept"/></hyperlinks>`
    const r = await sheetWith(body, `<Relationship Id="rId1" Type="${REL}/hyperlink" Target="https://example.com/a%20b?x=1&amp;y=2" TargetMode="External"/><Relationship Id="rId2" Type="${REL}/hyperlink" Target="file.xlsx" TargetMode="External"/>`)
    const c = r.book.sheets[0].cells
    expect(c.A1).toEqual({ value: 'x', link: { url: 'https://example.com/a%20b?x=1&y=2', tooltip: 't & u' } })
    expect(c.B1).toEqual({ value: 1, link: { location: 'S!A1' } })
    for (const a of ['C2', 'D2', 'C3', 'D3']) expect(c[a]).toEqual({ link: { url: 'file.xlsx', location: 'x' } })
    expect(c.E1).toBeUndefined()
    expect(c.E2).toEqual({ link: { location: 'kept' } })
    expect(codes(r.issues)).toEqual(['hyperlink E1 E2'])
  })

  it('a hyperlink over more than 4,096 cells goes to the cells of its range that hold something, reported', async () => {
    const r = await sheetWith(`${DATA}<hyperlinks><hyperlink ref="A1:A1048576" location="S!B1"/></hyperlinks>`, '')
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 'x', link: { location: 'S!B1' } }, B1: { value: 1 } })
    expect(codes(r.issues)).toEqual(['hyperlink A1:A1048576'])
  })
})

const COMMENTS = (list: string, authors = '<author>Ann</author><author>tc={T1}</author>') => `<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>${authors}</authors><commentList>${list}</commentList></comments>`
const VML = (shapes: string) => `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">${shapes}</xml>`
const note = (row: number, col: number, more = '', kind = 'Note') => `<v:shape><x:ClientData ObjectType="${kind}">${more}<x:Row>${row}</x:Row><x:Column>${col}</x:Column></x:ClientData></v:shape>`

describe('notes read', () => {
  const withNotes = (list: string, vml: string, extra: Record<string, string> = {}, sheetRels = '') =>
    sheetWith(`${DATA}<legacyDrawing r:id="rIdV"/>`, `<Relationship Id="rIdC" Type="${REL}/comments" Target="../comments1.xml"/><Relationship Id="rIdV" Type="${REL}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>${sheetRels}`, { 'xl/comments1.xml': COMMENTS(list), 'xl/drawings/vmlDrawing1.vml': VML(vml), ...extra })

  it('the author and the plain text; runs of several looks, a note always shown and other shapes reported', async () => {
    const list =
      '<comment ref="A1" authorId="0"><text><r><rPr><b/><sz val="9"/></rPr><t>Ann:</t></r><r><rPr><sz val="9"/></rPr><t xml:space="preserve">\nchecked</t></r></text></comment>' +
      '<comment ref="B1" authorId="7"><text><r><rPr><b/></rPr><t>bold alone</t></r></text></comment><comment ref="C3" authorId="0"><text><t>plain</t></text></comment>'
    const r = await withNotes(list, note(0, 0) + note(2, 2, '<x:Visible/>') + note(5, 5, '', 'Checkbox'))
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 'x', note: { author: 'Ann', text: 'Ann:\nchecked' } }, B1: { value: 1, note: { text: 'bold alone' } }, C3: { note: { author: 'Ann', text: 'plain' } } })
    expect(codes(r.issues)).toEqual(['comment A1', 'comment C3', 'drawing'])
  })

  it('threads of comments as notes: the first author, every text with the replies authored, at the cell of the placeholder; placeholders left out', async () => {
    const TNS = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments'
    const threads =
      `<ThreadedComments xmlns="${TNS}"><threadedComment ref="B1" personId="{P1}" id="{T1}"><text>Question?</text></threadedComment>` +
      '<threadedComment ref="B1" personId="{P2}" id="{T2}" parentId="{T1}"><text>Answer\non two lines</text></threadedComment>' +
      '<threadedComment ref="D4" personId="{P9}" id="{T3}" done="1"><text>No person</text></threadedComment></ThreadedComments>'
    const persons = `<personList xmlns="${TNS}"><person displayName="Ann" id="{P1}"/><person displayName="Борис" id="{P2}"/></personList>`
    const f = parts({ sheets: [{ name: 'S', body: `${DATA}<legacyDrawing r:id="rIdV"/>` }], strings: ['x'] })
    const read2 = async (placeholderAt: string) => {
      const list = `<comment ref="${placeholderAt}" authorId="1"><text><t>[Threaded comment] …</t></text></comment><comment ref="A1" authorId="0"><text><t>a note</t></text></comment>`
      return readXlsx(
        await pack({
          ...f,
          'xl/_rels/workbook.xml.rels': f['xl/_rels/workbook.xml.rels']!.replace('</Relationships>', '<Relationship Id="rIdP" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/></Relationships>'),
          'xl/persons/person.xml': persons,
          'xl/worksheets/_rels/sheet1.xml.rels': rels(`<Relationship Id="rIdC" Type="${REL}/comments" Target="../comments1.xml"/><Relationship Id="rIdV" Type="${REL}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/><Relationship Id="rIdT" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>`),
          'xl/comments1.xml': COMMENTS(list),
          'xl/drawings/vmlDrawing1.vml': VML(note(0, 0) + note(0, 1, '<x:Visible/>')),
          'xl/threadedComments/threadedComment1.xml': threads,
        }),
      )
    }
    const r = await read2('B1')
    const thread = { author: 'Ann', text: 'Question?\n\nБорис: Answer\non two lines' }
    expect(r.book.sheets[0].cells).toEqual({ A1: { value: 'x', note: { author: 'Ann', text: 'a note' } }, B1: { value: 1, note: thread }, D4: { note: { text: 'No person' } } })
    // the visible placeholder's box is not reported: its thread takes its place
    expect(codes(r.issues)).toEqual(['comment B1 D4'])
    // an Excel without threads moved the placeholder with its row: the thread goes with it
    const moved = await read2('B2')
    expect(moved.book.sheets[0].cells.B2).toEqual({ note: thread })
    expect(moved.book.sheets[0].cells.B1).toEqual({ value: 1 })
  })
})
