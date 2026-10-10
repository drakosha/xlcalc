/**
 * What a session's cells show: changes carry the text, edits that recompute nothing (a format, a width, a row's or
 * a column's format, merged areas) and structure operations return the cells whose look changed, and over seeded
 * edits the looks a host keeps from the changes equal a full computation's.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput, CellChange, CellDisplay, CellEdit, LayoutEdit, StructureOp } from '../src/contract.js'
import { compute, openSession } from '../src/index.js'
import { mulberry32 } from '../src/synth.js'

const clock = { now: '2026-10-07T12:00:00Z', seed: 1 }

const book = (): BookInput => ({
  sheets: [{ name: 'S', cells: { A1: { value: 1234.5 }, A2: { formula: 'A1*2' }, B1: { value: 'note' }, C1: { value: 0.25, format: '0%' } }, columns: { A: { width: 12 }, B: { width: 3 } } }],
  styles: [{ font: { bold: true } }],
})

const looks = (r: { changes: CellChange[] }) => r.changes.map((c) => [c.cell, c.value, c.text, ...(c.color ? [c.color] : [])])

describe('a session shows its cells', () => {
  it('changes carry the text of the new values', () => {
    const s = openSession(book(), { clock })
    expect(looks(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 0.5 } }]))).toEqual([
      ['A1', 0.5, '0.5'],
      ['A2', 1, '1'],
    ])
  })

  it('a new format returns the cell with its text; a new style alone returns the cell with it', () => {
    const s = openSession(book(), { clock })
    expect(looks(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 1234.5, format: '#,##0.00' } }]))).toEqual([['A1', 1234.5, '1,234.50']])
    expect(s.setCells([{ sheet: 'S', cell: 'B1', input: { value: 'note', style: 0 } }]).changes).toEqual([{ sheet: 'S', cell: 'B1', value: 'note', text: 'note', style: 0 }])
    expect(looks(s.setCells([{ sheet: 'S', cell: 'A2', input: { formula: 'A1*2', format: '[Red]0' } }]))).toEqual([['A2', 2469, '2469', '#FF0000']])
  })

  it('setLayout: a narrower column, a row format, a merged area', () => {
    const s = openSession(book(), { clock })
    expect(looks(s.setLayout([{ sheet: 'S', columns: 'A', set: { width: 3 } }]))).toEqual([
      ['A1', 1234.5, '##'],
      ['A2', 2469, '##'],
    ])
    expect(looks(s.setLayout([{ sheet: 'S', rows: '2', set: { format: '0.0' } }]))).toEqual([])
    expect(looks(s.setLayout([{ sheet: 'S', columns: 'A', set: { width: null } }, { sheet: 'S', rows: '1:2', set: { format: '0.0' } }]))).toEqual([
      ['A1', 1234.5, '1234.5'],
      ['A2', 2469, '2469.0'],
    ])
    expect(looks(s.setLayout([{ sheet: 'S', columns: 'D', set: { width: 1 } }, { sheet: 'S', merges: ['C1:D1'] }]))).toEqual([])
    expect(looks(s.setLayout([{ sheet: 'S', columns: 'C', set: { width: 1 } }]))).toEqual([['C1', 0.25, '#']])
    expect(looks(s.setLayout([{ sheet: 'S', merges: [] }]))).toEqual([['C1', 0.25, '']])
  })

  it('a structure operation returns the cells that look different where they land', () => {
    const s = openSession(book(), { clock })
    const r = s.restructure({ op: 'moveRange', from: { sheet: 'S', range: 'A1' }, to: { sheet: 'S', cell: 'B3' } })
    // the formula over the moved cell is listed with its new text
    expect(r.ok && looks(r)).toEqual([['A2', 2469, '2469'], ['B3', 1234.5, '##']])
    expect(r.ok && r.changes[0].formula).toBe('B3*2')
    const back = s.restructure({ op: 'moveRange', from: { sheet: 'S', range: 'B3' }, to: { sheet: 'S', cell: 'A1' } })
    expect(back.ok && looks(back)).toEqual([['A1', 1234.5, '1234.5'], ['A2', 2469, '2469']])
  })

  it('display: false: changes and results without text; a format alone lists the cell with it', () => {
    const s = openSession(book(), { clock, display: false })
    expect(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 1234.5, format: '0.0' } }]).changes).toEqual([{ sheet: 'S', cell: 'A1', value: 1234.5, numberFormat: '0.0' }])
    expect(s.setLayout([{ sheet: 'S', columns: 'A', set: { width: 1 } }]).changes).toEqual([])
    expect(s.result().sheets[0].cells.A1).toEqual({ value: 1234.5, format: '0.0' })
  })
})

describe('seeded edits: the looks a host keeps from the changes equal a full computation', () => {
  const formats = [undefined, '0.00', '#,##0', '[Red]0;[Blue]-0', 'dd/mm/yyyy', '0%', '@', '* #,##0']
  const values = [1234.5, -0.5, 123456789, 0.000123, 'text', true, 45688.75, 0]

  for (const seed of [1, 2, 3]) {
    it(`seed ${seed}, 150 edits`, () => {
      const rnd = mulberry32(seed)
      const pick = <T>(xs: readonly T[]) => xs[Math.floor(rnd() * xs.length)]
      const cell = () => `${pick(['A', 'B', 'C', 'D'])}${1 + Math.floor(rnd() * 6)}`
      const s = openSession(
        { sheets: [{ name: 'S', cells: { A1: { value: 5 }, B2: { formula: 'A1*1000' }, C3: { formula: 'SEQUENCE(2,2)*A1' } } }, { name: 'T', cells: { A1: { formula: "S!B2&\"\"" } } }] },
        { clock },
      )
      const kept = new Map<string, CellDisplay>()
      const keep = (r: { changes: CellChange[] }) => r.changes.forEach((c) => kept.set(`${c.sheet}!${c.cell}`, { text: c.text!, ...(c.color ? { color: c.color } : {}) }))
      const shown = (cells: Record<string, { text?: string; color?: string }>, sheet: string) =>
        Object.entries(cells).map(([a, c]): [string, CellDisplay] => [`${sheet}!${a}`, { text: c.text!, ...(c.color ? { color: c.color } : {}) }])
      for (const sh of s.result().sheets) for (const [k, d] of shown(sh.cells, sh.name)) kept.set(k, d)
      for (let i = 0; i < 150; i++) {
        const op = rnd()
        if (op < 0.35) {
          const e: CellEdit = { sheet: 'S', cell: cell(), input: rnd() < 0.2 ? null : { value: pick(values), ...(rnd() < 0.5 ? { format: pick(formats) } : {}) } }
          keep(s.setCells([e]))
        } else if (op < 0.55) {
          const at = cell()
          const now = s.book().sheets[0].cells[at]
          keep(s.setCells([{ sheet: 'S', cell: at, input: { ...(now ?? { value: pick(values) }), format: pick(formats) } }]))
        } else if (op < 0.8) {
          const edit: LayoutEdit =
            rnd() < 0.5
              ? { sheet: 'S', columns: pick(['A', 'B', 'C', 'A:B']), set: { width: rnd() < 0.2 ? null : 1 + Math.floor(rnd() * 14) } }
              : rnd() < 0.5
                ? { sheet: 'S', rows: pick(['1', '2:3', '5']), set: { format: rnd() < 0.3 ? null : pick(formats.filter((f) => f !== undefined)) } }
                : { sheet: 'S', merges: rnd() < 0.5 ? [] : [pick(['A1:B1', 'B2:C2', 'D1:D3'])] }
          keep(s.setLayout([edit]))
        } else {
          const op: StructureOp = pick<StructureOp>([
            { op: 'insertColumns', sheet: 'S', column: 'B' },
            { op: 'deleteColumns', sheet: 'S', column: 'D' },
            { op: 'moveRange', from: { sheet: 'S', range: 'A1:A2' }, to: { sheet: 'S', cell: 'B5' } },
          ])
          const before = new Map(kept)
          const r = s.restructure(op)
          if (!r.ok) continue
          kept.clear()
          for (const [k, d] of before) {
            const [sheet, addr] = k.split('!')
            const to = r.mapCell({ sheet, cell: addr })
            if (to) kept.set(`${to.sheet}!${to.cell}`, d)
          }
          keep(r)
        }
        const full = compute(s.book(), { clock: s.clock }).sheets.flatMap((sh) => shown(sh.cells, sh.name))
        // a blank cell and a cell the result does not list both show nothing
        const visible = (m: Iterable<[string, CellDisplay]>) => new Map([...m].filter(([, d]) => d.text !== '' || d.color))
        expect(visible(kept)).toEqual(visible(full))
        expect(s.result().sheets.flatMap((sh) => shown(sh.cells, sh.name))).toEqual(full)
      }
    })
  }
})
