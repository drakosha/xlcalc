import { expect, it } from 'vitest'
import { compute, createSession, editStructure, InputError, type BookInput, type TextRun } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { pack, parts, partText } from './xlsx-fixtures.js'

const runs: TextRun[] = [{ start: 0, font: { bold: true } }, { start: 1, font: { color: '#FF0000' } }, { start: 3 }]
const rich = (): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 'a😀b', richText: structuredClone(runs) }, A2: { value: 'a😀b' } } }] })
const spans = (b: BookInput) => b.sheets[0].cells.A1.richText

it('holds rich text without changing a value or calculation hash', () => {
  const b = rich()
  const plain = structuredClone(b)
  delete (plain.sheets[0].cells.A1 as any).richText
  const a = compute(b)
  const p = compute(plain)
  expect(a.sheets[0].cells.A1.value).toBe('a😀b')
  expect(a.hash).toBe(p.hash)
  expect(a.inputHash).toBe(p.inputHash)
})

it('keeps differently formatted copies of the same string distinct in shared strings', async () => {
  const { data } = await writeXlsx({ book: rich() }, { values: 'omit' })
  expect(await partText(data, 'xl/sharedStrings.xml')).toContain('<rPr>')
  const next = await readXlsx(data)
  expect(spans(next.book)?.map((r) => r.start)).toEqual([0, 1, 3])
  expect(spans(next.book)?.[0].font).toMatchObject({ bold: true })
  expect(spans(next.book)?.[1].font).toMatchObject({ color: '#FF0000' })
  expect((next.book.sheets[0].cells.A2 as any).richText).toBeUndefined()
  expect(next.issues.filter((i) => i.code === 'richText')).toEqual([])
})

it('keeps unedited rich strings byte for byte and writes explicit edits inline', async () => {
  const { data } = await writeXlsx({ book: rich() }, { values: 'omit' })
  const read = await openKept(data)
  expect((await writeKept(data, { book: read.book })).data).toEqual(data)
  const b = structuredClone(read.book)
  ;(b.sheets[0].cells.A1 as any).richText[1].font.color = '#0000FF'
  const written = await writeKept(data, { book: b })
  expect(await partText(written.data, 'xl/worksheets/sheet1.xml')).toContain('t="inlineStr"')
  expect(spans((await readXlsx(written.data)).book)?.[1].font).toMatchObject({ color: '#0000FF' })
})

it('clears runs when the text changes, and undo restores them', () => {
  const s = createSession(rich())
  expect((s.getCell('S', 'A1') as any).richText).toEqual(runs)
  s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 'replacement' } }])
  expect(spans(s.getBook())).toBeUndefined()
  s.undo()
  expect(spans(s.getBook())).toEqual(runs)
  s.redo()
  expect(spans(s.getBook())).toBeUndefined()
  s.close()
})

it('a header renamed by a name typed beside it loses its runs, the others keep theirs', () => {
  const header = [{ start: 0, font: { bold: true } }]
  const s = createSession({
    sheets: [
      {
        name: 'S',
        cells: { A1: { value: 'X' }, B1: { value: 'Y', richText: structuredClone(header) }, C1: { value: 'Z', richText: structuredClone(header) } },
        tables: [{ name: 'T', range: 'A1:C2', columns: [{ name: 'X' }, { name: 'Y' }, { name: 'Z' }] }],
      },
    ],
  })
  s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 'Y' } }])
  const cells = s.getBook().sheets[0].cells
  expect(cells.B1).toEqual({ value: 'Y2' })
  expect(cells.C1).toEqual({ value: 'Z', richText: header })
  s.close()
})

it('copies runs with their text', () => {
  const done = editStructure(rich(), { op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', range: 'B1' } })
  expect(done.ok).toBe(true)
  if (done.ok) expect((done.book.sheets[0].cells.B1 as any).richText).toEqual(runs)
})

it.each([[{ start: 1 }], [{ start: 0 }, { start: 0 }], [{ start: 0 }, { start: 2 }], [{ start: 0 }, { start: 4 }]].map((bad) => [bad]))(
  'rejects a first offset, duplicate, split surrogate or out-of-text offset: %j',
  (bad) => {
    const b = rich()
    ;(b.sheets[0].cells.A1 as any).richText = bad
    try {
      compute(b)
      throw new Error('no InputError')
    } catch (e) {
      expect(e).toBeInstanceOf(InputError)
      expect((e as InputError).path).toContain('.richText')
    }
  },
)

it('requires a text constant, not a formula or number', () => {
  for (const value of [1, true, null]) {
    const b = rich()
    b.sheets[0].cells.A1.value = value
    expect(() => compute(b)).toThrow(InputError)
  }
  const b = rich()
  b.sheets[0].cells.A1.formula = '"a😀b"'
  expect(() => compute(b)).toThrow(InputError)
})

it('reports a formatting-only edit and allows run updates against existing text', () => {
  const s = createSession(rich())
  const next = structuredClone(runs)
  next[1].font!.color = '#0000FF'
  const changed = s.setCells([{ sheet: 'S', cell: 'A1', input: { richText: next } }])
  expect(changed.changes[0]).toMatchObject({ cell: 'A1', value: 'a😀b', richText: next })
  expect((s.getCell('S', 'A1') as any).richText).toEqual(next)
  s.close()
})

it('clears rich formatting without clearing text', () => {
  const s = createSession(rich())
  s.clearCells([{ sheet: 'S', range: 'A1' }], 'formats')
  expect(s.getCell('S', 'A1')!.value).toBe('a😀b')
  expect(spans(s.getBook())).toBeUndefined()
  s.close()
})

it('a paste of values does not copy source text formatting', () => {
  const done = editStructure(rich(), { op: 'copyRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', range: 'B1' }, paste: 'values' })
  expect(done.ok).toBe(true)
  if (done.ok) {
    expect(done.book.sheets[0].cells.B1.value).toBe('a😀b')
    expect((done.book.sheets[0].cells.B1 as any).richText).toBeUndefined()
  }
})

it('reads first-run, missing and empty run properties using Excel font defaults', async () => {
  const styles =
    '<fonts count="2"><font><b/><u/><sz val="14"/><color rgb="FF0000FF"/><name val="Courier New"/></font><font><b/><u/><sz val="18"/><color rgb="FF00FF00"/><name val="Arial"/></font></fonts><cellXfs count="2"><xf fontId="0"/><xf fontId="1"/></cellXfs>'
  const text = '<r><t>A</t></r><r><rPr><color rgb="FFFF0000"/></rPr><t>B</t></r><r><t>C</t></r><r><rPr/><t>D</t></r>'
  const f = parts({ styles, sheets: [{ name: 'S', body: '<sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is>' + text + '</is></c></row></sheetData>' }] })
  const read = await readXlsx(await pack(f))
  const rs = spans(read.book)!
  expect(rs[0].font).toMatchObject({ name: 'Arial', size: 18, color: '#00FF00', bold: true, underline: 'single' })
  expect(rs[1].font).toMatchObject({ name: 'Courier New', size: 14, color: '#FF0000' })
  expect(rs[1].font?.bold ?? false).toBe(false)
  expect(rs[1].font?.underline).toBeUndefined()
  expect(rs[2].font).toMatchObject({ name: 'Courier New', size: 14, color: '#0000FF', bold: true, underline: 'single' })
  expect(rs[3].font).toEqual(rs[0].font)
})

it('undoes and redoes a rich-text-only edit even though its inputHash is unchanged', () => {
  const s = createSession(rich())
  const hash = compute(s.getBook()).inputHash
  const next = structuredClone(runs)
  next[1].font!.color = '#0000FF'
  s.setCells([{ sheet: 'S', cell: 'A1', input: { richText: next } }])
  expect(compute(s.getBook()).inputHash).toBe(hash)
  s.undo()
  expect(spans(s.getBook())).toEqual(runs)
  s.redo()
  expect(spans(s.getBook())).toEqual(next)
  s.close()
})

it('returns detached runs from cell views', () => {
  const s = createSession(rich())
  const view = s.getCell('S', 'A1')!
  view.richText![1].font!.color = '#0000FF'
  expect(spans(s.getBook())?.[1].font?.color).toBe('#FF0000')
  s.close()
})

it("applies a font patch to every run while retaining each run's other properties", () => {
  const s = createSession(rich())
  s.patchCells([{ sheet: 'S', cell: 'A1', set: { style: { font: { italic: true } } } }])
  expect(spans(s.getBook())?.map((r) => r.font)).toEqual([{ bold: true, italic: true }, { color: '#FF0000', italic: true }, { italic: true }])
  s.undo()
  expect(spans(s.getBook())).toEqual(runs)
  s.close()
})

it('preserves runs when changing the cell fill or base style', () => {
  const s = createSession({ ...rich(), styles: [{ font: { bold: true } }] })
  s.patchCells([{ sheet: 'S', cell: 'A1', set: { style: { fill: { color: '#0000FF' } } } }])
  expect(spans(s.getBook())).toEqual(runs)
  s.patchCells([{ sheet: 'S', cell: 'A1', set: { style: 0 } }])
  expect(spans(s.getBook())).toEqual(runs)
  s.close()
})

it.each(['write', 'keep'] as const)('bounds rich-text font names in %s', async (mode) => {
  const { data } = await writeXlsx({ book: rich() }, { values: 'omit' })
  const b = mode === 'keep' ? (await openKept(data)).book : rich()
  b.sheets[0].cells.A1.richText![0].font = { name: 'x'.repeat(40) }
  b.sheets[0].cells.A1.richText![1].font = { name: '' }
  const out = mode === 'keep' ? await writeKept(data, { book: b }) : await writeXlsx({ book: b }, { values: 'omit' })
  const read = await readXlsx(out.data)
  expect(spans(read.book)?.[0].font?.name).toBe('x'.repeat(31))
  expect(spans(read.book)?.[1].font?.name).toBe('Calibri')
  expect(out.issues.map((i) => i.code)).toEqual(expect.arrayContaining(['style', 'textLength']))
})
