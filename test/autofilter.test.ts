/**
 * A sheet's AutoFilter (`SheetInput.autoFilter`): checked as input and kept out of inputHash; set and removed by a
 * session's layout edit, with undo and redo, in every kind of session and both execution modes; moved by the structure
 * operations as a range reference of its sheet; read and written by readXlsx and writeXlsx; written by writeKept over
 * the file, which keeps the file's criteria where the range is the workbook's.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createSession, editStructure, InputError, type BookInput, type Execution, type SheetInput, type StructureOp } from '../src/index.js'
import { inputHash } from '../src/canon.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { KEEP_ALL } from '../src/xlsx/write/sheet.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { pack } from './xlsx-fixtures.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-27T12:00:00Z', seed: 1 }

/** Data: a header row and four rows of amounts under A1:C5, with a total in E1; Other: a cell. */
const book = (data: Partial<SheetInput> = { autoFilter: { range: 'A1:C5' } }): BookInput => ({
  sheets: [
    {
      name: 'Data',
      cells: {
        A1: { value: 'Region' }, B1: { value: 'Amount' }, C1: { value: 'Month' }, E1: { formula: 'SUM(B2:B5)' },
        A2: { value: 'North' }, B2: { value: 1 }, A3: { value: 'South' }, B3: { value: 2 }, A4: { value: 'East' }, B4: { value: 3 }, A5: { value: 'West' }, B5: { value: 4 },
      },
      ...data,
    },
    { name: 'Other', cells: { A1: { value: 'x' } } },
  ],
})

const failure = (f: () => unknown): InputError => {
  try {
    f()
  } catch (e) {
    if (e instanceof InputError) return e
    throw e
  }
  throw new Error('no InputError')
}

describe('the AutoFilter as input', () => {
  it('a cell or a range from its top-left cell; anything else, and one over a table, is an InputError with its path', () => {
    expect(() => compute(book({ autoFilter: { range: 'B2' } }))).not.toThrow()
    for (const [autoFilter, path] of [
      [{ range: 'C5:A1' }, 'book.sheets[0].autoFilter.range'],
      [{ range: 'A:C' }, 'book.sheets[0].autoFilter.range'],
      [{ range: 7 }, 'book.sheets[0].autoFilter.range'],
      [{}, 'book.sheets[0].autoFilter.range'],
      [{ range: 'A1:C5', sort: true }, 'book.sheets[0].autoFilter'],
      ['A1:C5', 'book.sheets[0].autoFilter'],
    ] as const)
      expect(failure(() => compute(book({ autoFilter: autoFilter as never }))).path).toBe(path)
    const e = failure(() => compute(book({ autoFilter: { range: 'A2:C9' }, tables: [{ name: 'T', range: 'A1:A5', columns: [{ name: 'Region' }] }] })))
    expect([e.path, e.message]).toEqual(['book.sheets[0].autoFilter.range', 'book.sheets[0].autoFilter.range: the AutoFilter over A2:C9 overlaps the table T (a table has a filter of its own)'])
  })

  it('takes no part in calculation or inputHash', () => {
    expect(inputHash(book())).toBe(inputHash(book({})))
    expect(compute(book()).hash).toBe(compute(book({})).hash)
  })
})

describe.each(['closures', 'eval'] as const)('the AutoFilter in a session, execution %s', (execution: Execution) => {
  it('setLayout sets and removes it, recomputing nothing; undo and redo; one over a table is refused', () => {
    const s = createSession(book({}), { clock, execution })
    const r = s.setLayout([{ sheet: 'Data', autoFilter: { range: 'A1:C5' } }])
    expect(r.changes).toEqual([])
    expect(s.getLayout().sheets[0].autoFilter).toEqual({ range: 'A1:C5' })
    s.setLayout([{ sheet: 'Data', autoFilter: null }])
    expect('autoFilter' in s.getBook().sheets[0]).toBe(false)
    s.undo()
    expect(s.getBook().sheets[0].autoFilter).toEqual({ range: 'A1:C5' })
    s.undo()
    expect('autoFilter' in s.getBook().sheets[0]).toBe(false)
    s.redo()
    expect(s.getBook().sheets[0].autoFilter).toEqual({ range: 'A1:C5' })
    expect(failure(() => s.setLayout([{ sheet: 'Data', autoFilter: { range: 'A1:C5' }, frozen: null } as never])).path).toBe('edits[0]')
    s.editStructure({ op: 'addTable', sheet: 'Other', table: { name: 'T', range: 'A1:A2', columns: [{ name: 'x' }] } })
    expect(failure(() => s.setLayout([{ sheet: 'Other', autoFilter: { range: 'A1:B3' } }])).message).toBe('edits[0].autoFilter.range: the AutoFilter over A1:B3 overlaps the table T (a table has a filter of its own)')
    expect(compute(s.getBook(), { clock: s.clock, execution }).hash).toBe(s.getHash())
    s.close()
  })
})

describe.each(KINDS)('the AutoFilter in %s', (kind) => {
  it('setLayout sets it, editStructure moves it, undo puts it back', async () => {
    const s = await (await opener(kind))(book({}), { clock })
    await s.setLayout([{ sheet: 'Data', autoFilter: { range: 'A1:C5' } }])
    await s.editStructure({ op: 'insertRows', sheet: 'Data', row: 3 })
    expect((await s.getLayout()).sheets[0].autoFilter).toEqual({ range: 'A1:C6' })
    await s.undo()
    await s.undo()
    expect('autoFilter' in (await s.getLayout()).sheets[0]).toBe(false)
    await s.close()
  })
})

/** The AutoFilter of a sheet after an operation on the book (`A1:C5` on Data), or the refusal's code. */
function after(op: StructureOp, sheet = 'Data', start = book()): string | undefined {
  const r = editStructure(start, op)
  if (!r.ok) return r.code
  return r.book.sheets.find((s) => s.name === sheet)?.autoFilter?.range
}

describe('the structure operations move the AutoFilter as a range reference of its sheet, a move only when it takes all of it', () => {
  it.each([
    ['rows inserted inside it grow it', { op: 'insertRows', sheet: 'Data', row: 3, count: 2 }, 'A1:C7'],
    ['rows inserted right under it leave it', { op: 'insertRows', sheet: 'Data', row: 6 }, 'A1:C5'],
    ['rows inserted above its header move it', { op: 'insertRows', sheet: 'Data', row: 1 }, 'A2:C6'],
    ['its header row deleted: gone', { op: 'deleteRows', sheet: 'Data', row: 1 }, undefined],
    ['rows deleted under its header', { op: 'deleteRows', sheet: 'Data', row: 2, count: 2 }, 'A1:C3'],
    ['all its rows deleted: gone', { op: 'deleteRows', sheet: 'Data', row: 1, count: 5 }, undefined],
    ['a column deleted inside it', { op: 'deleteColumns', sheet: 'Data', column: 'B' }, 'A1:B5'],
    ['all its columns deleted: gone', { op: 'deleteColumns', sheet: 'Data', column: 'A', count: 3 }, undefined],
    ['a column inserted inside it', { op: 'insertColumns', sheet: 'Data', column: 'B' }, 'A1:D5'],
    ['a column inserted right of it', { op: 'insertColumns', sheet: 'Data', column: 'D' }, 'A1:C5'],
    ['moved whole on its sheet', { op: 'moveRange', from: { sheet: 'Data', range: 'A1:C5' }, to: { sheet: 'Data', cell: 'E7' } }, 'E7:G11'],
    ['moved whole to another sheet: gone from its own', { op: 'moveRange', from: { sheet: 'Data', range: 'A1:C5' }, to: { sheet: 'Other', cell: 'B2' } }, undefined],
    ['its header row moved away: left', { op: 'moveRange', from: { sheet: 'Data', range: 'A1:C1' }, to: { sheet: 'Data', cell: 'H1' } }, 'A1:C5'],
    ['cells moved over all of it: left', { op: 'moveRange', from: { sheet: 'Data', range: 'H1:J5' }, to: { sheet: 'Data', cell: 'A1' } }, 'A1:C5'],
    ['its last rows moved down: left', { op: 'moveRange', from: { sheet: 'Data', range: 'A4:C5' }, to: { sheet: 'Data', cell: 'A9' } }, 'A1:C5'],
    ['a copy over it leaves it', { op: 'copyRange', from: { sheet: 'Data', range: 'A2' }, to: { sheet: 'Data', range: 'A3:C3' } }, 'A1:C5'],
    ['its sheet renamed', { op: 'renameSheet', sheet: 'Data', name: 'Sales' }, 'A1:C5'],
  ] as [string, StructureOp, string | undefined][])('%s', (_, op, range) => {
    expect(after(op, op.op === 'renameSheet' ? 'Sales' : 'Data')).toBe(range)
  })

  it('no other sheet gets it, a moved sheet keeps it, a table over all of it takes its place, over part of it is refused', () => {
    expect(after({ op: 'moveRange', from: { sheet: 'Data', range: 'A1:C5' }, to: { sheet: 'Other', cell: 'B2' } }, 'Other')).toBeUndefined()
    expect(after({ op: 'moveSheet', sheet: 'Data' })).toBe('A1:C5')
    const r = editStructure(book(), { op: 'addTable', sheet: 'Data', table: { name: 'T', range: 'A1:C5', columns: [{ name: 'Region' }, { name: 'Amount' }, { name: 'Month' }] } })
    expect(r.ok && [r.book.sheets[0].tables?.[0].range, 'autoFilter' in r.book.sheets[0]]).toEqual(['A1:C5', false])
    expect(after({ op: 'addTable', sheet: 'Data', table: { name: 'T', range: 'G1:G2', columns: [{ name: 'g' }] } })).toBe('A1:C5')
    const part = editStructure(book(), { op: 'addTable', sheet: 'Data', table: { name: 'T', range: 'C1:D5', columns: [{ name: 'Month' }, { name: 'Total' }] } })
    expect(part.ok === false && [part.code, part.message]).toEqual(['table', "C1:D5 overlaps the sheet's AutoFilter over A1:C5: a table takes its place over all of it alone"])
  })
})

/** A part of a file, as text. */
async function partText(data: Uint8Array, name: string): Promise<string> {
  return (await openArchive(data).text(name)) ?? ''
}

/** A file's parts as text, each changed by `edits` (part, text found, its replacement). */
async function changed(data: Uint8Array, edits: [string, string, string][]): Promise<Uint8Array> {
  const a = openArchive(data)
  const parts: Record<string, string> = {}
  for (const n of a.names) parts[n] = (await a.text(n))!
  for (const [part, from, to] of edits) {
    if (!parts[part].includes(from)) throw new Error(`${part} has no ${from}`)
    parts[part] = parts[part].replace(from, to)
  }
  return pack(parts)
}

describe('the AutoFilter in xlsx files', () => {
  it('writeXlsx writes it over its range; a filtered sheet gets the criteria every row meets on it, not on a table', async () => {
    const plain = (await writeXlsx({ book: book() })).data
    const xml = await partText(plain, 'xl/worksheets/sheet1.xml')
    expect(xml).toContain('</sheetData><autoFilter ref="A1:C5"/></worksheet>')
    expect(xml).not.toContain('filterMode')
    // with its filter database, as Excel writes it when it makes an AutoFilter
    expect(await partText(plain, 'xl/workbook.xml')).toContain('<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$C$5</definedName></definedNames>')
    expect((await readXlsx(plain)).book.sheets[0].autoFilter).toEqual({ range: 'A1:C5' })
    const tables = [{ name: 'T', range: 'G1:G3', columns: [{ name: 'g' }] }]
    const filtered = book({ autoFilter: { range: 'A1:C5' }, filtered: true, rows: { 3: { hidden: true } }, tables })
    filtered.sheets[0].cells.G1 = { value: 'g' }
    const data = (await writeXlsx({ book: filtered })).data
    const fx = await partText(data, 'xl/worksheets/sheet1.xml')
    expect(fx).toContain('<sheetPr filterMode="1"/>')
    expect(fx).toContain(`<autoFilter ref="A1:C5">${KEEP_ALL}</autoFilter>`)
    expect(await partText(data, 'xl/tables/table1.xml')).toContain('<autoFilter ref="G1:G3"/>')
    const r = await readXlsx(data)
    expect([r.book.sheets[0].autoFilter, r.book.sheets[0].filtered]).toEqual([{ range: 'A1:C5' }, true])
  })

  it('readXlsx leaves out one over a table, and reports it', async () => {
    const data = (await writeXlsx({ book: book({ tables: [{ name: 'T', range: 'A1:C5', columns: [{ name: 'Region' }, { name: 'Amount' }, { name: 'Month' }] }] }) })).data
    const r = await readXlsx(await changed(data, [['xl/worksheets/sheet1.xml', '</sheetData>', '</sheetData><autoFilter ref="B2:B9"/>']]))
    expect(r.book.sheets[0].autoFilter).toBeUndefined()
    expect(r.issues.map((i) => `${i.code} ${i.sheet}: ${i.message}`)).toContain('autoFilter Data: an AutoFilter over B2:B9, which is no range of cells of the sheet or overlaps a table, is left out')
  })
})

/** A file with Data's AutoFilter over A1:C5 holding criteria on its second column, and its filter database, as Excel saves it. */
async function criteriaFile(): Promise<Uint8Array> {
  const data = (await writeXlsx({ book: book() })).data
  return changed(data, [['xl/worksheets/sheet1.xml', '<autoFilter ref="A1:C5"/>', '<autoFilter ref="A1:C5"><filterColumn colId="1"><filters><filter val="2"/></filters></filterColumn></autoFilter>']])
}

/** What writeKept writes over the criteria file after `edit`: Data's `<autoFilter>` and the workbook's names. */
async function keptAfter(edit: (s: ReturnType<typeof createSession>) => void): Promise<{ filter: string; names: string }> {
  const original = await criteriaFile()
  const read = await openKept(original)
  const s = createSession(read.book, { cached: read.cached, clock })
  edit(s)
  const data = (await writeKept(original, s.getSnapshot())).data
  s.close()
  const sheet = await partText(data, 'xl/worksheets/sheet1.xml')
  return { filter: /<autoFilter[^>]*\/>|<autoFilter[\s\S]*?<\/autoFilter>/.exec(sheet)?.[0] ?? '', names: /<definedNames>[\s\S]*<\/definedNames>/.exec(await partText(data, 'xl/workbook.xml'))?.[0] ?? '' }
}

describe('the AutoFilter written over the file', () => {
  const NAME = (ref: string) => `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">${ref}</definedName></definedNames>`

  it('read into the workbook with its criteria reported; where the workbook keeps its range, the file keeps its criteria', async () => {
    const read = await openKept(await criteriaFile())
    expect([read.book.sheets[0].autoFilter, read.book.sheets[0].filtered]).toEqual([{ range: 'A1:C5' }, true])
    expect(await keptAfter((s) => s.setCells([{ sheet: 'Data', cell: 'D9', input: { value: 1 } }]))).toEqual({
      filter: '<autoFilter ref="A1:C5"><filterColumn colId="1"><filters><filter val="2"/></filters></filterColumn></autoFilter>',
      names: NAME('Data!$A$1:$C$5'),
    })
    // moved by an operation as the workbook moves it, criteria and filter database along
    expect(await keptAfter((s) => s.editStructure({ op: 'insertColumns', sheet: 'Data', column: 'A' }))).toEqual({
      filter: '<autoFilter ref="B1:D5"><filterColumn colId="1"><filters><filter val="2"/></filters></filterColumn></autoFilter>',
      names: NAME('Data!$B$1:$D$5'),
    })
  })

  it("a host's range is written from the workbook, its criteria lost; one removed goes with its filter database; one deleted is no refusal; a new one gets one", async () => {
    expect(await keptAfter((s) => s.setLayout([{ sheet: 'Data', autoFilter: { range: 'A1:C9' } }]))).toEqual({ filter: `<autoFilter ref="A1:C9">${KEEP_ALL}</autoFilter>`, names: NAME('Data!$A$1:$C$9') })
    expect(await keptAfter((s) => s.setLayout([{ sheet: 'Data', autoFilter: null }, { sheet: 'Data', filtered: false }]))).toEqual({ filter: '', names: '' })
    expect(await keptAfter((s) => s.editStructure({ op: 'deleteRows', sheet: 'Data', row: 1, count: 5 }))).toEqual({ filter: '', names: '' })
    // moved to another sheet, it goes; the sheet, filtered still, gets an autofilter over its cells, as writeXlsx writes it
    expect(await keptAfter((s) => s.editStructure({ op: 'moveRange', from: { sheet: 'Data', range: 'A1:C5' }, to: { sheet: 'Other', cell: 'B2' } }))).toEqual({ filter: `<autoFilter ref="E1:E1">${KEEP_ALL}</autoFilter>`, names: '' })
    expect((await keptAfter((s) => s.setLayout([{ sheet: 'Other', autoFilter: { range: 'A1' } }]))).names).toBe(
      '<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$C$5</definedName><definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">Other!$A$1</definedName></definedNames>',
    )
  })
})
