/** API conveniences: the default clock of `compute` and `openSession`, clear errors on wrong input. */
import { describe, expect, it } from 'vitest'
import { compute, InputError, openSession, validateFormula, validateName, type BookInput } from '../src/index.js'

const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: 'NOW()' }, A2: { formula: 'RAND()' }, A3: { value: 1 }, A4: { formula: 'A3*2' } } }] }

describe('default clock', () => {
  it('compute without a clock uses the current moment and a random seed; its clock reproduces the result', () => {
    const r = compute(book)
    expect(Math.abs(Date.parse(r.clock.now) - Date.now())).toBeLessThan(60_000)
    expect(Number.isInteger(r.clock.seed) && r.clock.seed >= 0 && r.clock.seed < 2 ** 32).toBe(true)
    const again = compute(book, { clock: r.clock })
    expect(again.hash).toBe(r.hash)
    expect(again.sheets[0].cells.A2.value).toBe(r.sheets[0].cells.A2.value)
    expect(compute(book).clock.seed).not.toBe(r.clock.seed)
  })

  it('openSession without options keeps the clock of opening; a full computation with it gives the same hash', () => {
    const s = openSession(book)
    const clock = s.clock
    s.setCells([{ sheet: 'S', cell: 'A3', input: { value: 5 } }])
    expect(s.clock).toEqual(clock)
    const edited: BookInput = { sheets: [{ name: 'S', cells: { ...book.sheets[0].cells, A3: { value: 5 } } }] }
    expect(s.result().hash).toBe(compute(edited, { clock }).hash)
    expect(s.setClock({ now: '2026-01-01T00:00:00Z', seed: 1 }).length).toBeGreaterThan(0)
  })
})

describe('input errors', () => {
  const sheet = (cells: unknown) => ({ sheets: [{ name: 'S', cells }] }) as unknown as BookInput
  const fails = (f: () => unknown, message: string | RegExp) => {
    let err: unknown
    try {
      f()
    } catch (e) {
      err = e
    }
    expect(err).toBeInstanceOf(InputError)
    expect(err).toBeInstanceOf(TypeError)
    expect((err as InputError).name).toBe('InputError')
    expect((err as Error).message).toMatch(message)
  }

  it('the book: wrong types, cell addresses, duplicate sheets, cell fields, dates', () => {
    fails(() => compute(null as unknown as BookInput), 'book: expected a book such as { sheets: [{ name: "Sheet1", cells: {} }] }, got null')
    fails(() => compute({ sheets: 'S' } as unknown as BookInput), 'book.sheets: expected an array of sheets, got the string "S"')
    fails(() => compute(sheet({ A1: 5 })), 'book.sheets[0].cells["A1"]: expected a cell object such as { value: 1 } or { formula: "A1*2" }, got number 5')
    fails(() => compute(sheet({ b2: { value: 1 } })), 'book.sheets[0].cells["b2"]: invalid cell address "b2": expected A1 style such as "B12" (column letters must be upper-case: "B2")')
    fails(() => compute(sheet({ A0: { value: 1 } })), 'invalid cell address "A0"')
    fails(() => compute(sheet({ XFE1: { value: 1 } })), 'cell address "XFE1" is outside the sheet (the last cell is XFD1048576)')
    fails(() => compute(sheet({ A1048577: { value: 1 } })), 'is outside the sheet')
    fails(() => compute(sheet({ A1: { val: 1 } })), 'book.sheets[0].cells["A1"]: unknown field "val" (expected value, formula, format, style, meta)')
    fails(() => compute(sheet({ A1: { value: { date: '31.01.2025' } } })), 'book.sheets[0].cells["A1"].value.date: invalid date "31.01.2025"')
    fails(() => compute(sheet({ A1: { value: [1] } })), 'book.sheets[0].cells["A1"].value: expected a number, string, boolean, null or { date: "YYYY-MM-DD" }, got an array')
    fails(() => compute(sheet({ A1: { formula: 1 } })), 'book.sheets[0].cells["A1"].formula: expected a string, got number 1')
    fails(() => compute({ sheets: [{ name: 'Data', cells: {} }, { name: 'data', cells: {} }] }), 'book.sheets[1].name: duplicate sheet name "data": book.sheets[0] is "Data" (sheet names are case-insensitive)')
    fails(() => compute({ sheets: [{ name: 'S', cells: {} }], names: [{ name: 'x', ref: "'S'!A1", scop: 'S' }] } as unknown as BookInput), 'book.names[0]: unknown field "scop"')
  })

  it('options and the clock', () => {
    const book = sheet({ A1: { value: 1 } })
    fails(() => compute(book, { now: '2026-01-01T00:00:00Z', seed: 1 } as never), 'opts: unknown field "now" (expected clock, externals, execution, display) (pass the clock as { clock: { now, seed } })')
    fails(() => compute(book, { clock: { now: 'yesterday', seed: 1 } }), 'opts.clock.now: invalid moment "yesterday"')
    fails(() => compute(book, { clock: { now: '2026-01-01T00:00:00Z', seed: '1' } as never }), 'opts.clock.seed: expected a finite number, got the string "1"')
    fails(() => openSession(book, { live: true, clock: { now: '2026-01-01T00:00:00Z', seed: 1 } } as never), 'opts: a live session takes its clock from the system')
    fails(() => compute(book, { execution: 'fast' } as never), 'opts.execution: expected "closures" or "eval", got the string "fast"')
    fails(() => openSession(book, { execution: 'Eval' } as never), 'opts.execution: expected "closures" or "eval", got the string "Eval"')
  })

  it('session calls and validation', () => {
    const s = openSession(sheet({ A1: { value: 1 } }), { display: false })
    fails(() => s.setCells([{ sheet: 'T', cell: 'A1', input: { value: 1 } }]), 'edits[0].sheet: unknown sheet "T"')
    fails(() => s.setCells([{ sheet: 's', cell: 'A1', input: { value: 1 } }, { sheet: 'S', cell: '1A', input: null }]), 'edits[1].cell: invalid cell address "1A"')
    fails(() => s.setCells([{ sheet: 'S', cell: 'A1' } as never]), 'edits[0].input: expected a cell object, or null to clear the cell')
    fails(() => s.setClock({ now: '2026-01-01', seed: Number.NaN }), 'clock.seed: expected a finite number, got number NaN')
    fails(() => s.explain('S', 'a1'), 'cell: invalid cell address "a1"')
    fails(() => s.setNames([{ name: 'x' } as never]), 'names[0].ref: expected a string, got undefined')
    fails(() => validateFormula(42 as never), 'text: expected a string, got number 42')
    fails(() => validateFormula('A1', { sheets: 'S' } as never), 'ctx.sheets: expected an array of strings, got the string "S"')
    fails(() => validateName(undefined as never), 'name: expected a string, got undefined')
    // a failed call leaves the session intact
    expect(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 2 } }])).toEqual([{ sheet: 'S', cell: 'A1', value: 2 }])
  })
})
