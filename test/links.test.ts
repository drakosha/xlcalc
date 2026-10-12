/**
 * Links to other workbooks without Excel: the forms a reference to one is read in and its canonical text, the check
 * against a book's links, the text shown and typed, the input check, the input hash, what structure operations do to a
 * link's references, the error reasons and origins, and the functions Excel answers otherwise over a closed book.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, editStructure, parseEntry, toInvariant, toLocal, validateFormula, validateNames, type BookInput, type ExternalBook } from '../src/index.js'
import { ruRU } from './locales.js'

const LINKS: ExternalBook[] = [
  { target: 'Book2.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 1, A2: 2, A3: 3, B1: 'x', B2: 'y', B3: 'x', C1: { error: 'NA' } } }, { name: 'My Sheet', cells: { A1: 11 } }], names: [{ name: 'Rate', ref: 'Sheet1!$A$2' }, { name: 'Local', ref: "'My Sheet'!$A$1", sheet: 0 }, { name: 'Konst', ref: '5' }] },
  { target: '../data/Data.xlsx', sheets: [{ name: 'Data', cells: { A1: 7 } }, { name: 'Gone' }] },
]

const book = (cells: Record<string, string>, extra: Partial<BookInput> = {}): BookInput => ({ sheets: [{ name: 'S', cells: Object.fromEntries(Object.entries(cells).map(([a, f]) => [a, { formula: f }])) }], externalBooks: LINKS, ...extra })
const inputHash = (b: BookInput) => compute(b, { display: false }).inputHash
const values = (b: BookInput) => Object.fromEntries(Object.entries(compute(b, { display: false }).sheets[0].cells).map(([a, c]) => [a, c.value]))

describe('the text of a reference to another workbook', () => {
  it('reads the forms a file and a person write and checks them to the canonical text by the link’s number', () => {
    const canon = (f: string) => {
      const r = validateFormula(f, { externalBooks: LINKS })
      return r.ok ? r.canonical : `${r.code}: ${r.message}`
    }
    expect(canon('[1]Sheet1!A1+1')).toBe("'[1]Sheet1'!A1+1")
    expect(canon("'[1]My Sheet'!$A$1:B9")).toBe("'[1]My Sheet'!$A$1:B9")
    expect(canon('[1]sheet1!A:A')).toBe("'[1]Sheet1'!A:A")
    expect(canon('[1]!Rate')).toBe('[1]!Rate')
    expect(canon('[1]Sheet1!Local')).toBe("'[1]Sheet1'!Local")
    expect(canon('[Book2.xlsx]Sheet1!A1')).toBe("'[1]Sheet1'!A1")
    expect(canon("'C:\\other\\[data.XLSX]Data'!A1")).toBe("'[2]Data'!A1")
    expect(canon('Book2.xlsx!Rate')).toBe('[1]!Rate')
    expect(canon("'../data/Data.xlsx'!Rate")).toBe('[2]!Rate')
    expect(canon('[3]Sheet1!A1')).toBe('external_ref: the book has no link to the workbook [3]')
    expect(canon('[Other.xlsx]Sheet1!A1')).toBe('external_ref: the book has no link to the workbook Other.xlsx')
    expect(canon('[1]Nope!A1')).toBe('external_ref: the workbook Book2.xlsx has no sheet Nope')
    expect(canon('[1]Sheet1:Sheet2!A1')).toBe('external_ref: a reference across the sheets of another workbook is not supported')
    expect(canon("[1]Sheet1!#REF!+'[Book2.xlsx]My Sheet'!#REF!")).toBe("'[1]Sheet1'!#REF!+'[1]My Sheet'!#REF!")
    expect(validateFormula('[1]Sheet1!A1').ok).toBe(false)
  })

  it('shows them as Excel shows them, in ru-RU too, and reads the shown text back', () => {
    const shown = (f: string, loc = 'en-US') => {
      const t = toLocal(f, loc === 'ru-RU' ? ruRU : loc, { externalBooks: LINKS })
      return t.ok ? t.formula : t.message
    }
    expect(shown("SUM('[1]Sheet1'!A1:A3)+[1]!Rate")).toBe('SUM([Book2.xlsx]Sheet1!A1:A3)+Book2.xlsx!Rate')
    expect(shown("SUM('[1]Sheet1'!A1:A3;1.5)".replace(';', ','), 'ru-RU')).toBe('СУММ([Book2.xlsx]Sheet1!A1:A3;1,5)')
    expect(shown("'[1]My Sheet'!A1*'[2]Data'!A1+'[2]Data'!Rate")).toBe("'[Book2.xlsx]My Sheet'!A1*'../data/[Data.xlsx]Data'!A1+'../data/[Data.xlsx]Data'!Rate")
    expect(shown("'[1]Sheet1'!#REF!")).toBe('[Book2.xlsx]Sheet1!#REF!')
    expect(toLocal("'[1]Sheet1'!A1", 'en-US')).toEqual({ ok: true, formula: "'[1]Sheet1'!A1" })
    expect(toInvariant('=СУММ([Book2.xlsx]Sheet1!A1:A3;Book2.xlsx!Rate)', ruRU, { externalBooks: LINKS })).toEqual({ ok: true, formula: "SUM('[1]Sheet1'!A1:A3,[1]!Rate)" })
    const r = toInvariant('=1+[Other.xlsx]Sheet1!A1', 'en-US', { externalBooks: LINKS })
    expect(r).toMatchObject({ ok: false, code: 'external_ref', at: 3 })
  })

  it('a typed one is numbered with the book’s links and refused without one of them', () => {
    expect(parseEntry('=[Book2.xlsx]Sheet1!A1*2', { externalBooks: LINKS })).toEqual({ ok: true, input: { formula: "'[1]Sheet1'!A1*2" } })
    expect(parseEntry('=[Other.xlsx]Sheet1!A1', { externalBooks: LINKS })).toMatchObject({ ok: false, code: 'external_ref', at: 2 })
    expect(parseEntry('=[Book2.xlsx]Sheet1!A1', {})).toEqual({ ok: true, input: { formula: '[Book2.xlsx]Sheet1!A1' } })
  })
})

describe('a book’s links', () => {
  it('are checked as input, and go into the input hash only when there are any', () => {
    const plain: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] }
    expect(inputHash({ ...plain, externalBooks: [] })).toBe(inputHash(plain))
    expect(inputHash({ ...plain, externalBooks: LINKS })).not.toBe(inputHash(plain))
    expect(inputHash({ ...plain, externalBooks: [{ ...LINKS[0], sheets: [{ name: 'Sheet1', cells: { A1: 2 } }] }] })).not.toBe(inputHash({ ...plain, externalBooks: LINKS }))
    expect(() => compute({ ...plain, externalBooks: [{ target: 'a', sheets: [{ name: 'S', cells: { A1: [1] } }] }] } as never)).toThrow('book.externalBooks[0].sheets[0].cells["A1"]: expected a number, text, a boolean, null or { error: code }')
  })

  it('compute: a stored value, a cell not stored blank and #REF! on a sheet marked or not stored, names, the book’s own names over links', () => {
    expect(values(book({ A1: "'[1]Sheet1'!A2*10", A2: "'[1]Sheet1'!A9", A3: "'[1]Sheet1'!C1", A4: '[1]!Rate', A5: "'[1]Sheet1'!Local", A6: '[1]!Konst', A7: '[1]!Nope', A8: "'[2]Gone'!A1", A9: "SUM('[1]Sheet1'!A:A)", A10: 'Ext*2' }, { names: [{ name: 'Ext', ref: "'[1]Sheet1'!$A$3" }] }))).toEqual({
      A1: 20, A2: 0, A3: { error: 'NA' }, A4: 2, A5: 11, A6: { error: 'REF' }, A7: { error: 'REF' }, A8: { error: 'REF' }, A9: 6, A10: 6,
    })
    // a name over a reference into a linked book that Excel lost is kept and gives #REF!
    const lost = compute(book({ A1: 'Lost+1' }, { names: [{ name: 'Lost', ref: "'[1]Sheet1'!#REF!" }] }), { display: false })
    expect([lost.sheets[0].cells.A1.value, lost.diagnostics.filter((d) => d.kind === 'name')]).toEqual([{ error: 'REF' }, []])
    const marked = LINKS.map((b, i) => (i ? b : { ...b, sheets: b.sheets.map((s) => ({ ...s, refreshError: true as const })) }))
    expect(values(book({ A1: "'[1]Sheet1'!A9", A2: "SUM('[1]Sheet1'!A:A)", A3: "'[1]Sheet1'!A1", A4: "MATCH(3,'[1]Sheet1'!A:A,0)" }, { externalBooks: marked }))).toEqual({ A1: { error: 'REF' }, A2: { error: 'REF' }, A3: 1, A4: 3 })
  })

  it('a position that takes an open book’s reference is #VALUE! over a link, directly, through a name or a LET name', () => {
    expect(values(book({ A1: "SUMIF('[1]Sheet1'!B1:B3,\"x\",'[1]Sheet1'!A1:A3)", A2: "COUNTIF(('[1]Sheet1'!B1:B3),\"x\")", A3: 'LET(r,\'[1]Sheet1\'!B1:B3,COUNTIF(r,"x"))', A4: 'COUNTIF(Ext,"x")', A5: "SUMPRODUCT(('[1]Sheet1'!B1:B3=\"x\")*'[1]Sheet1'!A1:A3)", A6: "COUNTIF(B1:B3,'[1]Sheet1'!B1)" }, { names: [{ name: 'Ext', ref: "'[1]Sheet1'!$B$1:$B$3" }] }))).toEqual({
      A1: { error: 'VALUE' }, A2: { error: 'VALUE' }, A3: { error: 'VALUE' }, A4: { error: 'VALUE' }, A5: 4, A6: 0,
    })
  })

  it('names of the book refer to links as the check takes them', () => {
    expect(validateNames([{ name: 'Ext', ref: "'[1]Sheet1'!$A$1" }, { name: 'Bad', ref: "'[3]Sheet1'!$A$1" }], { sheets: ['S'], externalBooks: LINKS })).toEqual([{ name: 'Bad', reason: 'the book has no link to the workbook [3]', dropped: false }])
  })
})

describe('structure operations and links', () => {
  const b = book({ A3: "'[1]Sheet1'!A3+A1", B1: "'[1]Sheet1'!A1" }, { sheets: [{ name: 'S', cells: { A1: { value: 1 }, A3: { formula: "'[1]Sheet1'!A3+A1" }, B1: { formula: "'[1]Sheet1'!A1" } } }, { name: 'Sheet1', cells: {} }] })
  const formulas = (x: BookInput) => Object.fromEntries(Object.entries(x.sheets[0].cells).flatMap(([a, c]) => (c.formula ? [[a, c.formula]] : [])))

  it('rows inserted, a sheet named like the link’s renamed or deleted leave a link’s reference as written', () => {
    const ins = editStructure(b, { op: 'insertRows', sheet: 'S', row: 1 })
    expect(ins.ok && formulas(ins.book)).toEqual({ A4: "'[1]Sheet1'!A3+A2", B2: "'[1]Sheet1'!A1" })
    const ren = editStructure({ ...b, names: [{ name: 'Lost', ref: "'[1]Sheet1'!#REF!" }] }, { op: 'renameSheet', sheet: 'Sheet1', name: 'Local' })
    expect(ren.ok && formulas(ren.book)).toEqual({ A3: "'[1]Sheet1'!A3+A1", B1: "'[1]Sheet1'!A1" })
    expect(ren.ok && ren.book.names).toEqual([{ name: 'Lost', ref: "'[1]Sheet1'!#REF!" }])
    const del = editStructure(b, { op: 'deleteSheet', sheet: 'Sheet1' })
    expect(del.ok && formulas(del.book)).toEqual({ A3: "'[1]Sheet1'!A3+A1", B1: "'[1]Sheet1'!A1" })
  })

  it('a copy moves its relative bounds and marks the link, as Excel does', () => {
    const c = editStructure({ ...b, sheets: [{ ...b.sheets[0], cells: { ...b.sheets[0].cells, C1: { formula: "'[1]Sheet1'!$A$1+'[1]Sheet1'!B1" } } }, b.sheets[1]] }, { op: 'copyRange', from: { sheet: 'S', range: 'C1' }, to: { sheet: 'S', range: 'C2:D2' } })
    expect(c.ok && [c.book.sheets[0].cells.C2.formula, c.book.sheets[0].cells.D2.formula]).toEqual(["'[1]Sheet1'!$A$1+'[1]Sheet1'!B2", "'[1]Sheet1'!$A$1+'[1]Sheet1'!C2"])
    expect(c.ok && c.book.externalBooks!.map((x) => x.sheets.every((s) => s.refreshError))).toEqual([true, false])
  })
})

describe('the reasons of errors from links', () => {
  it('name the linked cell and begin at the formula that reads it', () => {
    const s = createSession(book({ A1: "'[1]Sheet1'!C1", A2: 'A1+1' }))
    expect(s.getCellIssue('S', 'A2')).toEqual({ error: 'NA', reason: "value not found, from 'S'!A1", origin: { sheet: 'S', cell: 'A1' } })
    expect(s.getCellIssue('S', 'A1')).toEqual({ error: 'NA', reason: "value not found, from '[1]Sheet1'!C1", origin: { sheet: 'S', cell: 'A1' } })
    const e = s.explain('S', 'A1')
    expect(e.text).toContain('[Book2.xlsx]Sheet1!C1')
  })
})
