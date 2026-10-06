/**
 * Session: recomputing only the dependents must give the same as a full computation in order. 200 seeded
 * edits (values of the lookup table, criteria and data, formulas with and without arrays, invalid formulas,
 * clearing, a cycle and breaking it); after each one the session hash equals the hash of a full `compute`
 * of the edited book.
 */
import { describe, expect, it } from 'vitest'
import { compute, debugGeneratedCode, openSession, type BookInput, type CellEdit, type CellInput } from '../src/index.js'
import { classAEdits, classBEdit, mulberry32, S, synthBook } from '../src/synth.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

function apply(book: BookInput, edits: CellEdit[]): BookInput {
  return {
    names: book.names,
    sheets: book.sheets.map((s) => {
      const mine = edits.filter((e) => e.sheet === s.name)
      if (!mine.length) return s
      const cells = { ...s.cells }
      for (const e of mine) {
        if (e.input === null) delete cells[e.cell]
        else cells[e.cell] = e.input
      }
      return { name: s.name, cells }
    }),
  }
}

/** Seeded edits: a mix of values and formulas. */
function edits(seed: number, count: number, synth: ReturnType<typeof synthBook>): CellEdit[][] {
  const rnd = mulberry32(seed)
  const a = classAEdits(synth, seed, count)
  const formulas: CellInput[] = [
    { formula: `SUMIFS('${S.data}'!D:D,'${S.data}'!B:B,$A3,'${S.data}'!E:E,"расход")` },
    { formula: 'SUM(B3:M3)*2' },
    { formula: 'SUM(B3:M3),1' }, // invalid: syntax
    { formula: 'NOPE(1)' },
    { formula: 'ROUND(B3/3,2)' },
    // arrays: lifted operators and functions, functions with an array answer, an array result (#SPILL!)
    { formula: 'SUMPRODUCT((B4:M4>0)*B4:M4)' },
    { formula: 'MAX(IF(B5:M5>100,B5:M5,0))+SUM(B6:M6*{1,0,1,0,1,0,1,0,1,0,1,0})' },
    { formula: 'INDEX(SORT(B4:M4,1,-1,TRUE),1,1)+COUNTA(UNIQUE(B5:M5,TRUE))' },
    { formula: `SUM(FILTER('${S.data}'!D2:D40,'${S.data}'!D2:D40>0,0))` },
    { formula: 'B4:M4*2' },
  ]
  const out: CellEdit[][] = []
  for (let i = 0; i < count; i++) {
    const x = rnd()
    if (x < 0.4) out.push([a[i]])
    else if (x < 0.6) out.push([classBEdit(synth, seed * 1000 + i)])
    else if (x < 0.8) out.push([{ sheet: S.months, cell: `${'BCDEFGHIJKLMN'[Math.floor(rnd() * 12)]}${3 + Math.floor(rnd() * 5)}`, input: formulas[Math.floor(rnd() * formulas.length)] }])
    else if (x < 0.85) out.push([{ sheet: S.months, cell: 'P1', input: { formula: 'P2+1' } }, { sheet: S.months, cell: 'P2', input: { formula: 'P1+1' } }]) // a cycle
    else if (x < 0.9) out.push([{ sheet: S.months, cell: 'P2', input: { value: 1 } }]) // breaks the cycle
    else if (x < 0.95) out.push([{ sheet: S.params, cell: 'B3', input: null }]) // clearing
    else out.push([a[i], classBEdit(synth, seed * 1000 + i), { sheet: S.items, cell: 'B2', input: { value: 3 } }]) // several at once
  }
  return out
}

describe('session', () => {
  it('200 edits: after each one the session hash equals the full computation hash', () => {
    const synth = synthBook({ seed: 1, rows: 340 })
    const session = openSession(synth.book, { clock })
    let book = synth.book
    for (const batch of edits(3, 200, synth)) {
      const changes = session.setCells(batch)
      book = apply(book, batch)
      const full = compute(book, { clock })
      const r = session.result()
      expect(r.hash).toBe(full.hash)
      // the reported changes agree with the values
      for (const ch of changes) expect(r.sheets.find((s) => s.name === ch.sheet)!.cells[ch.cell]?.value ?? null).toEqual(ch.value)
    }
    session.close()
  }, 300_000)

  it('setClock recomputes volatile cells; setNames rebuilds', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: 'TODAY()' }, A2: { formula: 'A1+1' }, A3: { value: 5 }, B1: { formula: 'x*2' } } }], names: [{ name: 'x', ref: "'S'!$A$3" }] }
    const s = openSession(book, { clock })
    expect(s.setClock({ now: '2026-10-05T00:00:00Z', seed: 1 }).map((c) => c.cell).sort()).toEqual(['A1', 'A2'])
    expect(s.result().hash).toBe(compute(book, { clock: { now: '2026-10-05T00:00:00Z', seed: 1 } }).hash)
    expect(s.setNames([{ name: 'x', ref: "'S'!$A$2" }])).toEqual([{ sheet: 'S', cell: 'B1', value: 2 * (46300 + 1) }])
  })

  it('a live session takes a clock on every recomputation; a full computation with its clock gives the same hash', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: 'NOW()' }, A2: { formula: 'RAND()' }, A3: { value: 1 }, A4: { formula: 'A3*2' } } }] }
    const s = openSession(book, { live: true })
    const c1 = s.clock
    expect(Math.abs(Date.parse(c1.now) - Date.now())).toBeLessThan(60_000)
    const ch = s.setCells([{ sheet: 'S', cell: 'A3', input: { value: 2 } }])
    const c2 = s.clock
    expect(c2.seed).not.toBe(c1.seed)
    expect(ch.map((c) => c.cell).sort()).toEqual(['A1', 'A2', 'A3', 'A4'].filter((c) => c !== 'A1' || c1.now !== c2.now))
    const saved = s.result()
    expect(saved.clock).toEqual(c2)
    expect(compute(apply(book, [{ sheet: 'S', cell: 'A3', input: { value: 2 } }]), { clock: c2 }).hash).toBe(saved.hash)
    expect(() => s.setClock(c1)).toThrow()
  })

  it('explain: the precedents tree with depth, node limit and text', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { formula: 'SUM(A1:A2)' }, A4: { formula: 'A3*2+A1' } } }] }
    const s = openSession(book, { clock })
    const e = s.explain('S', 'A4')
    expect(e.root).toMatchObject({ ref: "'S'!A4", value: 7, formula: "A3*2+A1" })
    expect(e.root.children.map((c) => c.ref)).toEqual(["'S'!A3", "'S'!A1"])
    expect(e.root.children[0].children[0]).toMatchObject({ ref: "'S'!A1:A2", cells: 2, children: [] })
    expect(e.text).toBe(["'S'!A4 = 7  ← A3*2+A1", "  'S'!A3 = 3  ← SUM(A1:A2)", "    'S'!A1:A2 (2 cells)", "  'S'!A1 = 1"].join('\n'))
    const x = s.explain('S', 'A4', { expand: true })
    expect(x.root.children[0].children[0].children.map((c) => c.ref)).toEqual(["'S'!A1", "'S'!A2"])
    const t = s.explain('S', 'A4', { depth: 1, maxNodes: 2 })
    expect(t.truncated).toBe(true)
    expect(t.nodes).toBeLessThanOrEqual(3)
  })

  it('debug dump of the code in evaluation order', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A2: { formula: 'A1*2' }, A1: { formula: '1+1' } } }] }
    const code = debugGeneratedCode(book, { clock })
    expect(code.indexOf("// 'S'!A1 = 1+1")).toBeLessThan(code.indexOf("// 'S'!A2 = A1*2"))
    expect(code).toContain('return ()=>add(1,1)')
  })
})
