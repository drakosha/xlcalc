/**
 * `Session.patchCells`: a cell's style, format, meta, note and link changed without its value or formula, on an empty
 * cell, a value, a formula, an old-kind array formula's area and a spill; nothing recomputed; the shown text a new
 * format changes; refused patches change nothing; in the same thread and in a thread as in a synchronous session, and
 * after a thread is lost, from the snapshot and the journal. With closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createInlineCalc, InputError, createSession, type BookInput, type Calc, type CellPatch, type Execution } from '../src/index.js'
import { CLOCK } from './async-books.js'
import { inside, pool } from './fixtures/pool.js'

const book = (): BookInput => ({
  styles: [{ font: { bold: true } }, { fill: { color: '#FFFF00' } }],
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 0.25, style: 0, meta: { id: 1 } },
        A2: { value: 2 },
        A3: { formula: 'A1*2', numberFormat: '0.00' },
        C1: { formula: 'A1:A2*10', array: 'C1:C3' },
        E1: { formula: 'SEQUENCE(3)' },
        G1: { value: 'Qty' }, G2: { value: 5 },
      },
      tables: [{ name: 'T', range: 'G1:G2', columns: [{ name: 'Qty' }] }],
    },
  ],
})

/** Patches over every kind of cell: empty, a value, a formula, an old-kind array's top-left and its area, a spill's formula and its cells, a table's header. */
const PATCHES: CellPatch[][] = [
  [{ sheet: 'S', cell: 'B5', set: { style: 1 } }],
  [{ sheet: 'S', cell: 'A1', set: { numberFormat: '0%', meta: null } }],
  [{ sheet: 'S', cell: 'A3', set: { numberFormat: '[Red]0.0;[Blue]-0.0', note: { author: 'Ann', text: 'checked' } } }],
  [{ sheet: 'S', cell: 'C1', set: { style: 0, numberFormat: '0.000' } }, { sheet: 'S', cell: 'C2', set: { numberFormat: '0.00' } }, { sheet: 'S', cell: 'C3', set: { style: 1 } }],
  [{ sheet: 'S', cell: 'E2', set: { numberFormat: '0.000', link: { url: 'https://example.com/' } } }, { sheet: 'S', cell: 'E1', set: { meta: [1, 2] } }],
  [{ sheet: 'S', cell: 'G1', set: { style: 0 } }],
  [{ sheet: 'S', cell: 'B5', set: { style: null } }, { sheet: 'S', cell: 'A3', set: { note: null, numberFormat: null } }],
]

const calcs: Calc[] = []
afterAll(() => Promise.all(calcs.map((c) => c.close())))

describe.each(['closures', 'eval'] as const)('patches of cells, execution %s', (execution: Execution) => {
  const opts = { clock: CLOCK, execution }

  it("the look changes, the value and formula stay, whatever the cell; nothing else is computed", () => {
    const s = createSession(book(), opts)
    const hash = s.getHash()
    const answers = PATCHES.map((p) => s.patchCells(p))
    // each patch one revision more; the cells whose own style or format changed, with the text they show by them
    expect(answers.map((a) => a.revision)).toEqual([1, 2, 3, 4, 5, 6, 7])
    expect(answers.map((a) => a.changes)).toEqual([
      [{ sheet: 'S', cell: 'B5', value: null, text: '', style: 1 }],
      [{ sheet: 'S', cell: 'A1', value: 0.25, text: '25%', numberFormat: '0%', numberFormatCode: '0%' }],
      [{ sheet: 'S', cell: 'A3', value: 0.5, text: '0.5', color: '#FF0000', numberFormat: '[Red]0.0;[Blue]-0.0', numberFormatCode: '[Red]0.0;[Blue]-0.0' }],
      [
        { sheet: 'S', cell: 'C1', value: 2.5, text: '2.500', spill: 'C1:C3', style: 0, numberFormat: '0.000', numberFormatCode: '0.000' },
        { sheet: 'S', cell: 'C2', value: 20, text: '20.00', spillOf: 'C1', numberFormat: '0.00', numberFormatCode: '0.00' },
        { sheet: 'S', cell: 'C3', value: { error: 'NA' }, text: '#N/A', spillOf: 'C1', style: 1 },
      ],
      [{ sheet: 'S', cell: 'E2', value: 2, text: '2.000', spillOf: 'E1', numberFormat: '0.000', numberFormatCode: '0.000' }],
      [{ sheet: 'S', cell: 'G1', value: 'Qty', text: 'Qty', style: 0 }],
      [
        { sheet: 'S', cell: 'A3', value: 0.5, text: '0.5', numberFormat: null },
        { sheet: 'S', cell: 'B5', value: null, text: '', style: null },
      ],
    ])
    const cells = s.getBook().sheets[0].cells
    expect(cells).toEqual({
      A1: { value: 0.25, style: 0, numberFormat: '0%' },
      A2: { value: 2 },
      A3: { formula: 'A1*2' },
      C1: { formula: 'A1:A2*10', array: 'C1:C3', style: 0, numberFormat: '0.000' },
      C2: { numberFormat: '0.00' },
      C3: { style: 1 },
      E1: { formula: 'SEQUENCE(3)', meta: [1, 2] },
      E2: { numberFormat: '0.000', link: { url: 'https://example.com/' } },
      G1: { value: 'Qty', style: 0 },
      G2: { value: 5 },
    })
    // the values hash is the book's, as a full computation gives it, and the table's column keeps its name
    expect(s.getHash()).toBe(hash)
    expect(s.getResult().sheets).toEqual(compute(s.getBook(), opts).sheets)
    expect(s.getBook().sheets[0].tables![0].columns[0].name).toBe('Qty')
    // the arrays answer as before and change as before: the area edited whole, the spill blocked
    expect(s.getCells('S', 'C1:C3')).toMatchObject({ C1: { value: 2.5, text: '2.500', array: 'C1:C3', style: 0 }, C3: { value: { error: 'NA' }, style: 1 } })
    expect(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 1 } }]).changes.map((c) => c.cell).sort()).toEqual(['A1', 'A3', 'C1'])
    expect(() => s.setCells([{ sheet: 'S', cell: 'C2', input: { value: 1 } }])).toThrow(InputError)
    s.setCells([{ sheet: 'S', cell: 'E3', input: { value: 'x' } }])
    expect(s.getCells('S', 'E1:E2')).toEqual({ E1: { value: { error: 'SPILL' }, text: '#SPILL!', formula: 'SEQUENCE(3)', meta: [1, 2] }, E2: { value: null, text: '', numberFormat: '0.000', numberFormatCode: '0.000', link: { url: 'https://example.com/' } } })
    expect(s.getHash()).toBe(compute(s.getBook(), opts).hash)
  })

  it('a live session keeps its clock and its volatile values: a patch computes nothing', () => {
    const s = createSession({ sheets: [{ name: 'S', cells: { A1: { formula: 'RAND()' }, A2: { formula: 'NOW()' } } }] }, { live: true, execution })
    const [clock, before] = [s.clock, s.getCells('S', 'A1:A2')]
    expect(s.patchCells([{ sheet: 'S', cell: 'B1', set: { note: { text: 'n' } } }]).changes).toEqual([])
    // a format whose text no General text has: the cell is listed, with the value it had
    expect(s.patchCells([{ sheet: 'S', cell: 'A1', set: { numberFormat: '0.0000" u"' } }]).changes[0].value).toEqual(before.A1.value)
    expect(s.clock).toEqual(clock)
    expect(s.getCells('S', 'A2').A2.value).toEqual(before.A2.value)
  })

  it('without the display a patch lists the cells whose look changed, without their text; patches of one cell apply in order; nothing left removes the cell', () => {
    const s = createSession(book(), { ...opts, display: false })
    expect(s.patchCells(PATCHES[1]).changes).toEqual([{ sheet: 'S', cell: 'A1', value: 0.25, numberFormat: '0%', numberFormatCode: '0%' }])
    expect(s.patchCells([{ sheet: 'S', cell: 'K9', set: { style: 0 } }, { sheet: 'S', cell: 'K9', set: { numberFormat: '0.0', style: null } }, { sheet: 'S', cell: 'K9', set: { meta: 'm' } }]).changes).toEqual([{ sheet: 'S', cell: 'K9', value: null, numberFormat: '0.0', numberFormatCode: '0.0' }])
    expect(s.getBook().sheets[0].cells.K9).toEqual({ numberFormat: '0.0', meta: 'm' })
    s.patchCells([{ sheet: 'S', cell: 'K9', set: { numberFormat: null, meta: null } }, { sheet: 'S', cell: 'K10', set: { style: null } }])
    expect('K9' in s.getBook().sheets[0].cells || 'K10' in s.getBook().sheets[0].cells).toBe(false)
  })

  it('a patch the session refuses changes nothing', () => {
    const s = createSession(book(), opts)
    const before = JSON.stringify(s.getBook())
    const refused: [unknown, string][] = [
      [[{ sheet: 'S', cell: 'A1', set: { style: 0 } }, { sheet: 'S', cell: 'A2', set: { style: 2 } }], 'patches[1].set.style'],
      [[{ sheet: 'S', cell: 'A1', set: { value: 3 } }], 'patches[0].set: unknown field "value"'],
      [[{ sheet: 'S', cell: 'A1', set: { formula: 'A2' } }], 'patches[0].set: unknown field "formula"'],
      [[{ sheet: 'Nope', cell: 'A1', set: {} }], 'patches[0].sheet: unknown sheet "Nope"'],
      [[{ sheet: 'S', cell: 'a1', set: {} }], 'patches[0].cell'],
      [[{ sheet: 'S', cell: 'A1' }], 'patches[0].set'],
      [[{ sheet: 'S', cell: 'A1', set: { link: { tooltip: 'x' } } }], 'patches[0].set.link: a hyperlink has a url, a location or both'],
      [[{ sheet: 'S', cell: 'A1', set: { numberFormat: 99 } }], 'patches[0].set.numberFormat'],
      [{}, 'patches: expected an array'],
    ]
    for (const [p, message] of refused) expect(() => s.patchCells(p as CellPatch[]), message).toThrow(message)
    expect(JSON.stringify(s.getBook())).toBe(before)
  })

  it('in the same thread and in a thread: the changes, the book and the hash of a synchronous session', async () => {
    const p = await pool({ threads: 1 })
    calcs.push(p)
    const inline = createInlineCalc({ sliceMs: 2 })
    calcs.push(inline)
    const y = createSession(book(), opts)
    const [a, b] = [await inline.createSession(book(), opts), await p.createSession(book(), opts)]
    for (const patches of PATCHES) {
      const want = y.patchCells(patches)
      expect(await a.patchCells(patches)).toEqual(want)
      expect(await b.patchCells(patches)).toEqual(want)
    }
    expect(await a.getBook()).toEqual(y.getBook())
    expect(await b.getBook()).toEqual(y.getBook())
    expect(await b.getCells('S', 'A1:G5')).toEqual(y.getCells('S', 'A1:G5'))
    expect(await b.getHash()).toBe(y.getHash())
    await expect(b.patchCells([{ sheet: 'S', cell: 'A1', set: { style: 9 } }])).rejects.toThrow(InputError)
    expect([a.revision, b.revision]).toEqual([7, 7])
  })

  it('a thread lost: the session comes back with its patches, from the journal and from a snapshot taken after them', async () => {
    for (const snapshotEvery of [64, 3]) {
      const { calc, inside: x, crash } = await inside({ threads: 1, snapshotEvery })
      calcs.push(calc)
      const s = await calc.createSession(book(), opts)
      const y = createSession(book(), opts)
      for (const patches of PATCHES.slice(0, 5)) {
        await s.patchCells(patches)
        y.patchCells(patches)
      }
      await s.setCells([{ sheet: 'S', cell: 'A2', input: { value: 3 } }])
      y.setCells([{ sheet: 'S', cell: 'A2', input: { value: 3 } }])
      await s.getHash()
      const r = [...x.recs][0]
      // the patches are in the journal, or in a snapshot taken after them
      if (snapshotEvery === 64) expect(r.journal.map((e) => e.op)).toEqual(['patchCells', 'patchCells', 'patchCells', 'patchCells', 'patchCells', 'setCells'])
      else expect(r.snapshot.seq).toBe(6)
      crash(0)
      expect(await s.getBook()).toEqual(y.getBook())
      expect(await s.getCells('S', 'A1:G5')).toEqual(y.getCells('S', 'A1:G5'))
      expect(await s.patchCells(PATCHES[5])).toEqual(y.patchCells(PATCHES[5]))
      expect(await s.getHash()).toBe(y.getHash())
    }
  }, 60_000)
})
