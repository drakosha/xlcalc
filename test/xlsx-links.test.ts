/**
 * Links to other workbooks in xlsx: the reader takes each link in the order the workbook lists them, with its target,
 * sheets, names and stored values of every kind; a link of another kind keeps its number; formulas and names over the
 * links come as written. The writer writes the links as Excel does, a book without links byte for byte as before, and
 * the book comes back from its file.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput, ExternalBook } from '../src/contract.js'
import { compute } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { codes, pack, parts, type Fixture } from './xlsx-fixtures.js'
import { ruRU } from './locales.js'

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const rels = (items: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`
const linkPath = (id: string, target: string, kind = `${REL}/externalLinkPath`) => `<Relationship Id="${id}" Type="${kind}" Target="${target}" TargetMode="External"/>`

/** A file whose workbook lists `links` (each an externalLink part's content and its relationships) in the order given by `order`. */
function linked(links: { xml: string; rels?: string }[], order = links.map((_, i) => i), more: Partial<Fixture> = {}): Promise<Uint8Array> {
  const f = parts({ sheets: [{ name: 'S', body: '<sheetData/>' }], ...more })
  f['xl/workbook.xml'] = f['xl/workbook.xml']!.replace('</sheets>', `</sheets><externalReferences>${order.map((i) => `<externalReference r:id="rIdL${i + 1}"/>`).join('')}</externalReferences>`)
  f['xl/_rels/workbook.xml.rels'] = f['xl/_rels/workbook.xml.rels']!.replace('</Relationships>', `${links.map((_, i) => `<Relationship Id="rIdL${i + 1}" Type="${REL}/externalLink" Target="externalLinks/externalLink${i + 1}.xml"/>`).join('')}</Relationships>`)
  links.forEach((l, i) => {
    f[`xl/externalLinks/externalLink${i + 1}.xml`] = `<externalLink xmlns="${MAIN}" xmlns:r="${REL}">${l.xml}</externalLink>`
    if (l.rels) f[`xl/externalLinks/_rels/externalLink${i + 1}.xml.rels`] = rels(l.rels)
  })
  return pack(f)
}

/** An `externalBook` over `Book2.xlsx` with the given content. */
const book2 = (body: string, target = 'Book2.xlsx') => ({ xml: `<externalBook r:id="rId1">${body}</externalBook>`, rels: linkPath('rId1', target) })

describe('xlsx links: reading', () => {
  it('links in the order of the workbook, each with its target, sheets and stored values', async () => {
    const one = book2('<sheetNames><sheetName val="Sheet1"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>1</v></cell></row></sheetData></sheetDataSet>', 'One.xlsx')
    const two = book2('<sheetNames><sheetName val="Data"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>2</v></cell></row></sheetData></sheetDataSet>', 'Two.xlsx')
    // the second part is the first link
    const r = await readXlsx(await linked([one, two], [1, 0]))
    expect(r.book.externalBooks).toEqual([
      { target: 'Two.xlsx', sheets: [{ name: 'Data', cells: { A1: 2 } }] },
      { target: 'One.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 1 } }] },
    ])
    expect(r.issues).toEqual([])
  })

  it('values of every kind; a cell stored empty is null; text stays text when empty; an unknown error is reported', async () => {
    const cells =
      '<row r="1"><cell r="A1"><v>1.5</v></cell><cell r="B1" t="n"><v>-2</v></cell><cell r="C1" t="str"><v>a &amp; b</v></cell><cell r="D1" t="s"><v>s</v></cell><cell r="E1" t="inlineStr"><is><t>i</t></is></cell></row>' +
      '<row r="2"><cell r="A2" t="b"><v>1</v></cell><cell r="B2" t="b"><v>0</v></cell><cell r="C2" t="e"><v>#N/A</v></cell><cell r="D2" t="e"><v>#DIV/0!</v></cell><cell r="E2" t="e"><v>#GETTING_DATA</v></cell></row>' +
      '<row r="4"><cell r="A4"/><cell r="B4"><v/></cell><cell r="C4" t="str"><v/></cell><cell r="d4"><v>4</v></cell><cell r="ZZZZ9"><v>9</v></cell></row>'
    const r = await readXlsx(await linked([book2(`<sheetNames><sheetName val="Sheet1"/></sheetNames><sheetDataSet><sheetData sheetId="0">${cells}</sheetData></sheetDataSet>`)]))
    expect(r.book.externalBooks![0].sheets[0].cells).toEqual({
      A1: 1.5,
      B1: -2,
      C1: 'a & b',
      D1: 's',
      E1: 'i',
      A2: true,
      B2: false,
      C2: { error: 'NA' },
      D2: { error: 'DIV0' },
      E2: { error: 'ERROR' },
      A4: null,
      B4: null,
      C4: '',
      D4: 4,
    })
    expect(r.issues.map((i) => `${i.code} ${i.cells}`)).toEqual(['errorValue [1]Sheet1!E2'])
  })

  it("the mark of a failed update; a sheet without data has no cells, one with empty data no values", async () => {
    const body =
      '<sheetNames><sheetName val="Marked"/><sheetName val="None"/><sheetName val="Empty"/><sheetName val="My Sheet"/></sheetNames>' +
      '<sheetDataSet><sheetData sheetId="0" refreshError="1"><row r="1"><cell r="A1"><v>5</v></cell></row></sheetData><sheetData sheetId="2"/><sheetData sheetId="3" refreshError="1"/><sheetData sheetId="9"/></sheetDataSet>'
    const r = await readXlsx(await linked([book2(body)]))
    expect(r.book.externalBooks![0].sheets).toEqual([{ name: 'Marked', cells: { A1: 5 }, refreshError: true }, { name: 'None' }, { name: 'Empty', cells: {} }, { name: 'My Sheet', cells: {}, refreshError: true }])
  })

  it("the other book's names: of the book and of a sheet, what they refer to without `=`; one of a sheet it does not have is left out, reported", async () => {
    const body =
      '<sheetNames><sheetName val="Sheet1"/></sheetNames><definedNames><definedName name="Rate" refersTo="=\'Sheet1\'!$C$1"/><definedName name="Local" refersTo="=Sheet1!$C$2" sheetId="0"/><definedName name="Lost" refersTo="#REF!"/><definedName name="Far" refersTo="=Sheet1!$A$1" sheetId="3"/></definedNames>'
    const r = await readXlsx(await linked([book2(body)]))
    expect(r.book.externalBooks![0]).toEqual({
      target: 'Book2.xlsx',
      sheets: [{ name: 'Sheet1' }],
      names: [
        { name: 'Rate', ref: "'Sheet1'!$C$1" },
        { name: 'Local', ref: 'Sheet1!$C$2', sheet: 0 },
        { name: 'Lost', ref: '#REF!' },
      ],
    })
    expect(r.issues.map((i) => `${i.code} ${i.name}`)).toEqual(['name [1]Far'])
  })

  it("the target is the link's own relationship, of any kind; Excel's alternate addresses do not change it", async () => {
    const sheet = '<sheetNames><sheetName val="Sheet1"/></sheetNames>'
    const missing = { xml: `<externalBook r:id="rId1">${sheet}</externalBook>`, rels: linkPath('rId1', 'Book2', 'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing') }
    const alternate = {
      xml: `<externalBook xmlns:xxl21="http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021" r:id="rId2"><xxl21:alternateUrls><xxl21:relativeUrl r:id="rId1"/></xxl21:alternateUrls>${sheet}</externalBook>`,
      rels: linkPath('rId1', '../Book3.xlsx') + linkPath('rId2', '/xlcalc/work/Book3.xlsx'),
    }
    const escaped = book2(sheet, 'file:///C:/data/Sub%20Folder/A&amp;B.xlsx')
    const r = await readXlsx(await linked([missing, alternate, escaped]))
    expect(r.book.externalBooks!.map((b) => b.target)).toEqual(['Book2', '/xlcalc/work/Book3.xlsx', 'file:///C:/data/Sub%20Folder/A&B.xlsx'])
    expect(r.issues).toEqual([])
  })

  it('a DDE or OLE link, or one whose part is missing, keeps its number as a book without sheets, reported in both languages', async () => {
    const dde = { xml: '<ddeLink ddeService="Excel" ddeTopic="Book9"><ddeItems><ddeItem name="R1C1"/></ddeItems></ddeLink>' }
    const ole = { xml: '<oleLink r:id="rId1" progId="Word.Document.12"/>', rels: linkPath('rId1', 'doc.docx', `${REL}/oleObject`) }
    const real = book2('<sheetNames><sheetName val="Sheet1"/></sheetNames>')
    const data = await linked([dde, real, ole])
    const f = parts({ sheets: [{ name: 'S', body: '<sheetData/>' }] })
    const r = await readXlsx(data)
    expect(r.book.externalBooks).toEqual([{ target: '', sheets: [] }, { target: 'Book2.xlsx', sheets: [{ name: 'Sheet1' }] }, { target: '', sheets: [] }])
    expect(r.issues.map((i) => i.message)).toEqual([
      'a link to the data of another program (DDE, OLE), not to a workbook, or a link whose part is missing, is not in the model: its number is kept by a book without sheets, and formulas that refer to it compute to an error',
    ])
    expect((await readXlsx(data, { locale: ruRU })).issues.map((i) => i.message)).toEqual([
      'Связь с данными другой программы (DDE, OLE), а не с книгой, или связь, чьей части нет в файле, не поддерживается: её номер занят книгой без листов, формулы с ней дают ошибку',
    ])
    // the workbook lists a link whose part the file does not have
    f['xl/workbook.xml'] = f['xl/workbook.xml']!.replace('</sheets>', '</sheets><externalReferences><externalReference r:id="rIdX"/></externalReferences>')
    const lost = await readXlsx(await pack(f))
    expect(lost.book.externalBooks).toEqual([{ target: '', sheets: [] }])
    expect(codes(lost.issues)).toEqual(['externalLink'])
  })

  it("this book's names over a link come as written", async () => {
    const r = await readXlsx(await linked([book2('<sheetNames><sheetName val="Sheet1"/></sheetNames>')], [0], { workbook: { after: '<definedNames><definedName name="Ext">[1]Sheet1!$A$2</definedName><definedName name="ExtR">[1]Sheet1!$A$1:$A$3</definedName></definedNames>' } }))
    expect(r.book.names).toEqual([
      { name: 'Ext', ref: '[1]Sheet1!$A$2' },
      { name: 'ExtR', ref: '[1]Sheet1!$A$1:$A$3' },
    ])
  })

  it('a formula over a number the links do not hold, or over a link that is no workbook, is reported', async () => {
    const dde = { xml: '<ddeLink ddeService="Excel" ddeTopic="Book9"/>' }
    const real = book2('<sheetNames><sheetName val="Sheet1"/></sheetNames>')
    const body = '<sheetData><row r="1"><c r="A1"><f>[3]Sheet1!A1</f></c><c r="B1"><f>[2]Sheet1!A1+1</f></c><c r="C1"><f>\'[9]My Sheet\'!A1</f></c></row></sheetData>'
    const r = await readXlsx(await linked([real, dde], [0, 1], { sheets: [{ name: 'S', body }] }))
    expect(r.book.sheets[0].cells).toEqual({ A1: { formula: '[3]Sheet1!A1' }, B1: { formula: '[2]Sheet1!A1+1' }, C1: { formula: "'[9]My Sheet'!A1" } })
    expect(codes(r.issues)).toEqual(['externalLink', 'externalLink A1 B1 C1'])
  })

  // fails until a copy moves references to other workbooks (src/structure/rewrite.ts, the lead's): the copies take the master's text
  it("a shared formula over a link: each copy moves its relative reference into the link's sheet, as a paste does", async () => {
    const real = book2('<sheetNames><sheetName val="Sheet1"/></sheetNames>')
    const body = '<sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:B2" si="0">[1]Sheet1!A1+[1]Sheet1!$B$1</f></c><c r="B1"><f t="shared" si="0"/></c></row><row r="2"><c r="A2"><f t="shared" si="0"/></c><c r="B2"><f t="shared" si="0"/></c></row></sheetData>'
    const r = await readXlsx(await linked([real], [0], { sheets: [{ name: 'S', body }] }))
    expect(r.book.sheets[0].cells).toEqual({
      A1: { formula: '[1]Sheet1!A1+[1]Sheet1!$B$1' },
      B1: { formula: "'[1]Sheet1'!B1+'[1]Sheet1'!$B$1" },
      A2: { formula: "'[1]Sheet1'!A2+'[1]Sheet1'!$B$1" },
      B2: { formula: "'[1]Sheet1'!B2+'[1]Sheet1'!$B$1" },
    })
    expect(r.issues).toEqual([])
  })

  it('formulas and names over the links the file holds read without issues', async () => {
    const real = book2('<sheetNames><sheetName val="Sheet1"/><sheetName val="My Sheet"/></sheetNames><definedNames><definedName name="Rate" refersTo="=Sheet1!$C$1"/></definedNames>')
    const body = "<sheetData><row r=\"1\"><c r=\"A1\"><f>[1]Sheet1!A1+1</f></c><c r=\"B1\"><f>SUM('[1]My Sheet'!A1:B9)</f></c><c r=\"C1\"><f>[1]!Rate*2</f></c></row></sheetData>"
    const r = await readXlsx(await linked([real], [0], { sheets: [{ name: 'S', body }], workbook: { after: '<definedNames><definedName name="Ext">[1]Sheet1!$A$2</definedName></definedNames>' } }))
    expect(r.book.sheets[0].cells).toEqual({ A1: { formula: '[1]Sheet1!A1+1' }, B1: { formula: "SUM('[1]My Sheet'!A1:B9)" }, C1: { formula: '[1]!Rate*2' } })
    expect(r.issues).toEqual([])
  })
})

const BOOK2: ExternalBook = {
  target: 'Book2.xlsx',
  sheets: [
    { name: 'Sheet1', cells: { A1: 1, A2: 2.5, B1: 'abc', C1: '', B2: true, B3: { error: 'NA' }, B4: { error: 'DIV0' }, A4: null }, refreshError: true },
    { name: 'My Sheet' },
    { name: 'Лист 1', cells: {} },
  ],
  names: [
    { name: 'Rate', ref: 'Sheet1!$C$1' },
    { name: 'Local', ref: 'Sheet1!$C$2', sheet: 0 },
    { name: 'Bad', ref: '#REF!' },
  ],
}

const part = async (data: Uint8Array, name: string) => (await openArchive(data).text(name))!

describe('xlsx links: writing', () => {
  it("each link as Excel writes its part, its relationship and the workbook's reference", async () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 } } }], externalBooks: [BOOK2, { target: 'Sub Folder/A&B.xlsx', sheets: [{ name: 'Data', cells: { A1: 7 } }] }] }
    const { data, issues } = await writeXlsx({ book })
    expect(issues).toEqual([])
    expect(await part(data, 'xl/externalLinks/externalLink1.xml')).toBe(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<externalLink xmlns="${MAIN}"><externalBook xmlns:r="${REL}" r:id="rId1">` +
        '<sheetNames><sheetName val="Sheet1"/><sheetName val="My Sheet"/><sheetName val="Лист 1"/></sheetNames>' +
        '<definedNames><definedName name="Rate" refersTo="=Sheet1!$C$1"/><definedName name="Local" refersTo="=Sheet1!$C$2" sheetId="0"/><definedName name="Bad" refersTo="=#REF!"/></definedNames>' +
        '<sheetDataSet><sheetData sheetId="0" refreshError="1"><row r="1"><cell r="A1"><v>1</v></cell><cell r="B1" t="str"><v>abc</v></cell><cell r="C1" t="str"><v></v></cell></row>' +
        '<row r="2"><cell r="A2"><v>2.5</v></cell><cell r="B2" t="b"><v>1</v></cell></row><row r="3"><cell r="B3" t="e"><v>#N/A</v></cell></row><row r="4"><cell r="A4"/><cell r="B4" t="e"><v>#DIV/0!</v></cell></row></sheetData>' +
        '<sheetData sheetId="2"/></sheetDataSet></externalBook></externalLink>',
    )
    expect(await part(data, 'xl/externalLinks/_rels/externalLink2.xml.rels')).toContain(`<Relationship Id="rId1" Type="${REL}/externalLinkPath" Target="Sub Folder/A&amp;B.xlsx" TargetMode="External"/>`)
    const wb = await part(data, 'xl/workbook.xml')
    expect(wb).toContain('</sheets><externalReferences><externalReference r:id="rId3"/><externalReference r:id="rId4"/></externalReferences><calcPr/>')
    expect(await part(data, 'xl/_rels/workbook.xml.rels')).toContain(`<Relationship Id="rId3" Type="${REL}/externalLink" Target="externalLinks/externalLink1.xml"/><Relationship Id="rId4" Type="${REL}/externalLink" Target="externalLinks/externalLink2.xml"/>`)
    expect(await part(data, '[Content_Types].xml')).toContain('<Override PartName="/xl/externalLinks/externalLink2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>')
  })

  it('a book without links writes no part of them; a link without sheets an empty book', async () => {
    const { data } = await writeXlsx({ book: { sheets: [{ name: 'S', cells: {} }] } })
    expect(openArchive(data).names.some((n) => n.includes('externalLink'))).toBe(false)
    expect(await part(data, 'xl/workbook.xml')).not.toContain('externalReferences')
    const empty = await writeXlsx({ book: { sheets: [{ name: 'S', cells: {} }], externalBooks: [{ target: '', sheets: [] }] } })
    expect(await part(empty.data, 'xl/externalLinks/externalLink1.xml')).toContain(`<externalBook xmlns:r="${REL}" r:id="rId1"/>`)
  })

  it("formulas and names over the links in the file's spelling; one over a link or a sheet the links do not have is written as its value, one named by the file a link has by the link's number", async () => {
    const book: BookInput = {
      sheets: [
        {
          name: 'S',
          cells: {
            A1: { formula: "'[1]Sheet1'!A1+1" },
            A2: { formula: "SUM('[1]My Sheet'!A1:B9)" },
            A3: { formula: '[1]!Rate*2' },
            A4: { formula: "'[1]Sheet1'!Local" },
            A5: { formula: '[Book2.xlsx]Sheet1!A2' },
            A6: { formula: '[3]Sheet1!A1' },
            A7: { formula: "'[1]Nope'!A1" },
          },
        },
      ],
      names: [
        { name: 'Ext', ref: "'[1]Sheet1'!$A$2" },
        { name: 'Far', ref: "'[3]Sheet1'!$A$2" },
      ],
      externalBooks: [BOOK2],
    }
    const { data, issues } = await writeXlsx({ book })
    const sheet = await part(data, 'xl/worksheets/sheet1.xml')
    expect([...sheet.matchAll(/<f[^>]*>([^<]*)<\/f>/g)].map((m) => m[1])).toEqual(['[1]Sheet1!A1+1', "SUM('[1]My Sheet'!A1:B9)", '[1]!Rate*2', '[1]Sheet1!Local', '[1]Sheet1!A2'])
    expect(await part(data, 'xl/workbook.xml')).toContain('<definedNames><definedName name="Ext">[1]Sheet1!$A$2</definedName></definedNames>')
    expect(issues.map((i) => `${i.code} ${i.cells ?? i.name}`)).toEqual(['formula A6,A7', 'noCachedValues undefined', 'name Far'])
  })

  it('the links, their formulas and names come back from the file as the book has them, with the same input hash', async () => {
    const book: BookInput = {
      sheets: [{ name: 'S', cells: { A1: { formula: '[1]Sheet1!A1+[3]Data!XFD1048576' }, A2: { formula: "SUM('[1]Лист 1'!A:A)" }, A3: { formula: '[1]!Rate' }, A4: { formula: 'Ext*2' } } }],
      names: [{ name: 'Ext', ref: '[1]Sheet1!$A$1:$A$3' }],
      externalBooks: [BOOK2, { target: '', sheets: [] }, { target: 'file:///C:/data/Book4.xlsx', sheets: [{ name: 'Data', cells: { XFD1048576: 'end' } }] }],
    }
    const r = await readXlsx((await writeXlsx({ book }, { values: 'omit' })).data)
    expect(r.book).toEqual(book)
    expect(r.issues).toEqual([])
    expect(compute(r.book).inputHash).toBe(compute(book).inputHash)
  })
})
