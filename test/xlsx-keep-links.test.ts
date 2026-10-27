/**
 * `writeKept` and the links to other workbooks: a link the host adds after the file's, with the values it read from the
 * other workbook, is written as a part of its own with its relationship and `<externalReference>`; the values stored of a
 * file's link that the host extended, or that a formula typed over it marked, are written into its part, through the
 * relationship the part has; a formula over either reads back as entered. A link removed, moved or pointed at another
 * file is refused (`links`), as formulas read links by their numbers; an edit elsewhere leaves the links' parts as the
 * file has them.
 */
import { describe, expect, it } from 'vitest'
import { createSession, type BookInput, type ExternalBook } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { partsOf } from './keep-files.js'

const clock = { now: '2026-10-27T12:00:00Z', seed: 1 }

/** A file with one link, to Other.xlsx, that S!A1 reads. */
async function linked(): Promise<Uint8Array> {
  const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: "'[1]Data'!B2" } } }], externalBooks: [{ target: 'Other.xlsx', sheets: [{ name: 'Data', cells: { B2: 5 } }] }] }
  return (await writeXlsx({ book })).data
}

/** The file opened, its links replaced by `links` of the book read (`f`), a session over it. */
async function session(data: Uint8Array, f: (links: ExternalBook[]) => ExternalBook[]) {
  const { book } = await openKept(data)
  return createSession({ ...book, externalBooks: f(book.externalBooks ?? []) }, { clock })
}

describe('writeKept writes the links a host gives', () => {
  it("a new link after the file's, and the values of the file's extended, with a formula over each", async () => {
    const data = await linked()
    const s = await session(data, ([other]) => [
      { ...other, sheets: [{ name: 'Data', cells: { B2: 5, C3: 7 } }] },
      { target: 'Отчёт.xlsx', sheets: [{ name: 'Лист', cells: { I8: 42 } }] },
    ])
    expect(s.setCellText('S', 'A2', "='[Отчёт.xlsx]Лист'!$I$8+[Other.xlsx]Data!C3")).toMatchObject({ ok: true })
    expect(s.getCell('S', 'A2')).toMatchObject({ formula: "'[2]Лист'!$I$8+'[1]Data'!C3", value: 49 })
    const w = await writeKept(data, s.getSnapshot())
    const parts = await partsOf(w.data)
    expect(parts.get('xl/workbook.xml')).toMatch(/<externalReferences><externalReference r:id="(rId\d+)"\/><externalReference r:id="(rId\d+)"\/><\/externalReferences>/)
    expect(parts.get('xl/externalLinks/_rels/externalLink2.xml.rels')).toContain('Target="Отчёт.xlsx" TargetMode="External"')
    expect(parts.get('[Content_Types].xml')).toContain('<Override PartName="/xl/externalLinks/externalLink2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>')
    const back = await readXlsx(w.data)
    expect(back.issues).toEqual([])
    expect(back.book.externalBooks).toEqual(s.getBook().externalBooks)
    // as a file stores a link's reference, without quotes
    expect(back.book.sheets[0].cells.A2).toEqual({ formula: '[2]Лист!$I$8+[1]Data!C3' })
  })

  it("the file's link keeps the relationship its part points at the other workbook by", async () => {
    const data = await linked()
    const parts = await partsOf(data)
    const id = /<externalBook[^>]* r:id="([^"]+)"/.exec(parts.get('xl/externalLinks/externalLink1.xml')!)![1]
    const s = await session(data, ([other]) => [{ ...other, sheets: [{ name: 'Data', cells: { B2: 6 } }] }])
    const written = await partsOf((await writeKept(data, s.getSnapshot())).data)
    expect(written.get('xl/externalLinks/externalLink1.xml')).toContain(`r:id="${id}"`)
    expect(written.get('xl/externalLinks/_rels/externalLink1.xml.rels')).toBe(parts.get('xl/externalLinks/_rels/externalLink1.xml.rels'))
    expect((await readXlsx((await writeKept(data, s.getSnapshot())).data)).book.externalBooks).toEqual([{ target: 'Other.xlsx', sheets: [{ name: 'Data', cells: { B2: 6 } }] }])
  })

  it('a link removed, moved or pointed at another file is refused with links; an edit elsewhere keeps the parts of the links', async () => {
    const data = await linked()
    const extra: ExternalBook = { target: 'Third.xlsx', sheets: [{ name: 'A', cells: { A1: 1 } }] }
    for (const f of [() => [], ([o]: ExternalBook[]) => [extra, o], ([o]: ExternalBook[]) => [{ ...o, target: 'Elsewhere.xlsx' }]]) {
      const s = await session(data, f)
      await expect(writeKept(data, s.getSnapshot())).rejects.toMatchObject({ name: 'KeptRefusal', code: 'links' })
    }
    const s = await session(data, (links) => links)
    s.setCells([{ sheet: 'S', cell: 'B1', input: { value: 1 } }])
    const [was, now] = [await partsOf(data), await partsOf((await writeKept(data, s.getSnapshot())).data)]
    for (const p of ['xl/externalLinks/externalLink1.xml', 'xl/externalLinks/_rels/externalLink1.xml.rels']) expect(now.get(p)).toBe(was.get(p))
  })
})
