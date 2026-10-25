/**
 * Hidden names (`NameInput.hidden`, xlsx `<definedName hidden="1">`): they calculate as any name, take part in the input
 * hash alone, and keep the flag through the xlsx reader and writer, the structure operations that rewrite names and the
 * sessions (synchronous, in the same thread and in the pool, with closures and eval), their undo and redo among them.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, editStructure, InputError, type BookInput, type Execution, type NameInput, type StructureOp } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import { read } from './xlsx-fixtures.js'

afterAll(closeCalcs)

const CLOCK = { now: '2026-10-25T12:00:00Z', seed: 1 }

/** Data with formulas over a hidden global name, a hidden local one, a hidden name over a table and a visible name; an empty sheet. */
const book = (): BookInput => ({
  sheets: [
    {
      name: 'Data',
      cells: { A1: { value: 2 }, A2: { value: 3 }, B1: { formula: 'rate*A1' }, B2: { formula: 'own+1' }, B3: { formula: 'SUM(qty)+shown' }, D1: { value: 'Qty' }, D2: { value: 4 } },
      tables: [{ name: 'T', range: 'D1:D2', columns: [{ name: 'Qty' }] }],
    },
    { name: 'Other', cells: {} },
  ],
  names: [
    { name: 'rate', ref: "'Data'!$A$2", hidden: true },
    { name: 'own', ref: "'Data'!$A$1", scope: 'Data', hidden: true },
    { name: 'qty', ref: 'T[Qty]', hidden: true },
    { name: 'shown', ref: '0.5' },
  ],
})

/** The book with every name visible. */
const shown = (b: BookInput): BookInput => ({ ...b, names: b.names!.map(({ hidden: _, ...n }) => n) })

/** Each name as `name@scope` with `!` when hidden. */
const flags = (names: NameInput[] | undefined) => (names ?? []).map((n) => `${n.name}${n.scope ? `@${n.scope}` : ''}${n.hidden ? '!' : ''}`)

describe.each(['closures', 'eval'] as const)('hidden names calculate as any name (%s)', (execution: Execution) => {
  it('the same values and values hash as visible names; the input hash tells them apart, and false is no flag', () => {
    const [hidden, visible] = [compute(book(), { clock: CLOCK, execution }), compute(shown(book()), { clock: CLOCK, execution })]
    expect(['B1', 'B2', 'B3'].map((a) => hidden.sheets[0].cells[a].value)).toEqual([6, 3, 4.5])
    expect(hidden.diagnostics).toEqual(visible.diagnostics)
    expect(hidden.hash).toBe(visible.hash)
    expect(hidden.inputHash).not.toBe(visible.inputHash)
    const unflagged = { ...book(), names: book().names!.map((n) => ({ ...n, hidden: false })) }
    expect(compute(unflagged, { clock: CLOCK, execution }).inputHash).toBe(visible.inputHash)
  })

  it('hidden is a boolean', () => {
    expect(() => compute({ ...book(), names: [{ name: 'x', ref: '1', hidden: 1 as never }] }, { execution })).toThrow(new InputError('book.names[0].hidden', 'expected a boolean, got number 1'))
  })
})

describe('hidden names in xlsx files', () => {
  it('a hidden name is read hidden, global or local, without an issue; a built-in hidden name is left out as before', async () => {
    const r = await read({
      sheets: [{ name: 'Data', body: '<sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>' }],
      workbook: {
        after:
          '<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$A$9</definedName><definedName name="secret" hidden="1">Data!$A$1</definedName>' +
          '<definedName name="own" localSheetId="0" hidden="1">Data!$A$1*2</definedName><definedName name="open" hidden="0">1</definedName></definedNames>',
      },
    })
    expect(r.book.names).toEqual([{ name: 'secret', ref: 'Data!$A$1', hidden: true }, { name: 'own', ref: 'Data!$A$1*2', scope: 'Data', hidden: true }, { name: 'open', ref: '1' }])
    expect(r.issues.map((i) => `${i.code} ${i.name}`)).toEqual(['name _xlnm._FilterDatabase'])
  })

  it('the writer writes hidden="1" after the scope, as Excel does, and the name reads back hidden; false writes nothing', async () => {
    const b = { ...book(), names: [...book().names!, { name: 'plain', ref: '1', hidden: false }] }
    const { data, issues } = await writeXlsx({ book: b })
    expect(issues.map((i) => i.code)).toEqual(['noCachedValues'])
    const xml = (await openArchive(data).text('xl/workbook.xml'))!
    expect(xml).toContain(
      '<definedName name="rate" hidden="1">\'Data\'!$A$2</definedName><definedName name="own" localSheetId="0" hidden="1">\'Data\'!$A$1</definedName><definedName name="qty" hidden="1">T[Qty]</definedName><definedName name="shown">0.5</definedName><definedName name="plain">1</definedName>',
    )
    const back = await readXlsx(data)
    expect(flags(back.book.names)).toEqual(['rate!', 'own@Data!', 'qty!', 'shown', 'plain'])
    expect(compute(back.book, { clock: CLOCK }).inputHash).toBe(compute(b, { clock: CLOCK }).inputHash)
  })
})

/** The operations that rewrite names, with the names after each. */
const OPS: [StructureOp, string[]][] = [
  [{ op: 'insertRows', sheet: 'Data', row: 1 }, ['rate!', 'own@Data!', 'qty!', 'shown']],
  [{ op: 'renameSheet', sheet: 'Data', name: 'Book' }, ['rate!', 'own@Book!', 'qty!', 'shown']],
  [{ op: 'moveSheet', sheet: 'Data' }, ['rate!', 'own@Data!', 'qty!', 'shown']],
  // a local name a formula moved or copied to another sheet uses is defined there too
  [{ op: 'copyRange', from: { sheet: 'Data', range: 'B2' }, to: { sheet: 'Other', range: 'B2' } }, ['rate!', 'own@Data!', 'qty!', 'shown', 'own@Other!']],
  [{ op: 'moveRange', from: { sheet: 'Data', range: 'B2' }, to: { sheet: 'Other', cell: 'C3' } }, ['rate!', 'own@Data!', 'qty!', 'shown', 'own@Other!']],
  [{ op: 'renameTable', table: 'T', name: 'Sales' }, ['rate!', 'own@Data!', 'qty!', 'shown']],
  [{ op: 'deleteTable', table: 'T' }, ['rate!', 'own@Data!', 'qty!', 'shown']],
  [{ op: 'deleteSheet', sheet: 'Data' }, ['rate!', 'qty!', 'shown']],
]

describe('structure operations keep a name hidden', () => {
  it.each(OPS)('%o', (op, names) => {
    const r = editStructure(book(), op)
    if (!r.ok) throw new Error(r.message)
    expect(flags(r.book.names)).toEqual(names)
    // the rewritten references are the visible names' own
    const plain = editStructure(shown(book()), op)
    expect(plain.ok && plain.book.names!.map((n) => n.ref)).toEqual(r.book.names!.map((n) => n.ref))
  })
})

/** Operations one after another in a session: rows above, a copy that carries a local name, a table and its sheet renamed. */
const SESSION_OPS: StructureOp[] = [
  { op: 'insertRows', sheet: 'Data', row: 1 },
  { op: 'copyRange', from: { sheet: 'Data', range: 'B3' }, to: { sheet: 'Other', range: 'B2' } },
  { op: 'renameTable', table: 'T', name: 'Sales' },
  { op: 'renameSheet', sheet: 'Data', name: 'Book' },
]

describe.each(['closures', 'eval'] as const)('a session keeps hidden names (%s)', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it('through setNames, structure operations, undo and redo, getBook, getLayout and a snapshot written to a file', async () => {
      const open = await opener(kind)
      const s = await open(book(), { clock: CLOCK, execution })
      const hashOf = async () => (await s.getResult()).inputHash
      expect(flags((await s.getBook()).names)).toEqual(['rate!', 'own@Data!', 'qty!', 'shown'])
      expect(await hashOf()).toBe(compute(book(), { clock: CLOCK }).inputHash)
      await s.setNames([...book().names!, { name: 'more', ref: "'Data'!$A$1*10", hidden: true }])
      expect(flags((await s.getBook()).names)).toEqual(['rate!', 'own@Data!', 'qty!', 'shown', 'more!'])
      expect(flags((await s.getLayout()).names)).toEqual(['rate!', 'own@Data!', 'qty!', 'shown', 'more!'])
      for (const op of SESSION_OPS) expect((await s.editStructure(op)).ok).toBe(true)
      const after = ['rate!', 'own@Book!', 'qty!', 'shown', 'more!', 'own@Other!']
      expect(flags((await s.getBook()).names)).toEqual(after)
      expect((await s.getBook()).names!.find((n) => n.scope === 'Other')).toEqual({ name: 'own', ref: "'Other'!$A$2", scope: 'Other', hidden: true })
      expect(await hashOf()).toBe(compute(await s.getBook(), { clock: CLOCK }).inputHash)
      // the operations and setNames undone
      for (let i = 0; i <= SESSION_OPS.length; i++) expect(await s.undo()).not.toBeNull()
      expect(await s.getBook()).toEqual(book())
      expect(await hashOf()).toBe(compute(book(), { clock: CLOCK }).inputHash)
      for (let i = 0; i <= SESSION_OPS.length; i++) expect(await s.redo()).not.toBeNull()
      expect(flags((await s.getBook()).names)).toEqual(after)
      const back = await readXlsx((await writeXlsx(await s.getSnapshot())).data)
      expect(flags(back.book.names)).toEqual(after)
      await s.close()
    })
  })
})
