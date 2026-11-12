import { describe, expect, it } from 'vitest'
import { compute, createSession, type BookInput, type CellInput } from '../src/index.js'

const book = (input: CellInput): BookInput => ({
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 3 },
        A2: { value: 1 },
        A3: { value: 2 },
        B1: { value: 30 },
        B2: { value: 10 },
        B3: { value: 20 },
        C1: { value: 300 },
        C2: { value: 100 },
        C3: { value: 200 },
        D1: { value: 1 },
        D2: { value: 2 },
        G1: input,
      },
    },
  ],
})

// Expected values were measured with Excel Formula2 / FormulaArray, rather than
// inferred from the dynamic result. In CSE, XLOOKUP and INDEX select each lookup
// result before the enclosing SUM; VLOOKUP and the other controls still aggregate.
const lookups = [
  { name: 'XLOOKUP reference', formula: 'SUM(XLOOKUP({1;2},A1:A3,B1:B3))' },
  { name: 'XLOOKUP key expression', formula: 'SUM(XLOOKUP({1;2}+0,A1:A3,B1:B3))' },
  { name: 'INDEX reference', formula: 'SUM(INDEX(B1:B3,{2;3}))' },
  { name: 'INDEX value expression', formula: 'SUM(INDEX(B1:B3+0,{2;3}))' },
  { name: 'INDEX literal values', formula: 'SUM(INDEX({30;10;20},{2;3}))' },
]

const singleCellLookups = [
  { name: 'XLOOKUP reference', formula: 'SUM(XLOOKUP({1;2},A1:A3,A1:A3))' },
  { name: 'XLOOKUP result coercion', formula: 'SUM(XLOOKUP({1;2},A1:A3,A1:A3)+0)' },
  { name: 'XLOOKUP return coercion', formula: 'SUM(XLOOKUP({1;2},A1:A3,A1:A3+0))' },
  { name: 'XLOOKUP under SUMPRODUCT', formula: 'SUMPRODUCT(XLOOKUP({1;2},A1:A3,A1:A3))' },
  { name: 'XLOOKUP literal values', formula: 'SUM(XLOOKUP({1;2},{3;1;2},{3;1;2}))' },
]

const controls = [
  { name: 'ABS', formula: 'SUM(ABS({-1;-2}))', value: 3 },
  { name: 'VLOOKUP', formula: 'SUM(VLOOKUP({1;2},A1:B3,2,FALSE))', value: 30 },
  { name: 'XMATCH', formula: 'SUM(XMATCH({1;2},A1:A3))', value: 5 },
  { name: 'FILTER', formula: 'SUM(FILTER(B1:B3,A1:A3<3))', value: 30 },
  { name: 'SEQUENCE', formula: 'SUM(SEQUENCE(2))', value: 3 },
  { name: 'SORT', formula: 'SUM(SORT(A1:A3))', value: 6 },
]

describe.each(['closures', 'eval'] as const)('CSE lookup evaluation (%s)', (execution) => {
  const values = (input: CellInput) => {
    const cells = compute(book(input), { execution, display: false }).sheets[0].cells
    return ['G1', 'G2', 'G3'].map((cell) => cells[cell]?.value)
  }

  it.each(lookups)('$name aggregates all selected values in a dynamic formula', ({ formula }) => {
    expect(values({ formula })).toEqual([30, undefined, undefined])
  })

  it.each(lookups)('$name selects each value before SUM in a two-cell CSE area', ({ formula }) => {
    expect(values({ formula, array: 'G1:G2' })).toEqual([10, 20, undefined])
  })

  it.each(singleCellLookups)('$name aggregates both keys in a dynamic formula', ({ formula }) => {
    expect(values({ formula })).toEqual([3, undefined, undefined])
  })

  it.each(singleCellLookups)('$name selects the first key before aggregation in a one-cell CSE area', ({ formula }) => {
    expect(values({ formula, array: 'G1' })).toEqual([1, undefined, undefined])
  })

  it('uses both lookup positions when the same CSE formula occupies two cells', () => {
    expect(values({ formula: singleCellLookups[0].formula, array: 'G1:G2' })).toEqual([1, 2, undefined])
  })

  it('selects the first lookup position rather than the smallest key', () => {
    expect(values({ formula: 'SUM(XLOOKUP({2;1},A1:A3,A1:A3))', array: 'G1' })).toEqual([2, undefined, undefined])
  })

  it.each(controls)('$name aggregates the full array in a dynamic formula', ({ formula, value }) => {
    expect(values({ formula })).toEqual([value, undefined, undefined])
  })

  it.each(controls)('$name repeats the aggregate in a two-cell CSE area', ({ formula, value }) => {
    expect(values({ formula, array: 'G1:G2' })).toEqual([value, value, undefined])
  })

  it.each([lookups[0], lookups[2]])('$name switches evaluation mode on edit, undo and redo', ({ formula }) => {
    const session = createSession(book({ formula }), { execution, display: false })
    const current = () => {
      const cells = session.getResult().sheets[0].cells
      return ['G1', 'G2', 'G3'].map((cell) => cells[cell]?.value)
    }
    expect(current()).toEqual([30, undefined, undefined])

    session.setCells([{ sheet: 'S', cell: 'G1', input: { formula, array: 'G1:G2' } }])
    expect(current()).toEqual([10, 20, undefined])

    session.undo()
    expect(current()).toEqual([30, undefined, undefined])

    session.redo()
    expect(current()).toEqual([10, 20, undefined])
  })
  it.each([
    ['SUM(ABS(XLOOKUP({1;2},A1:A3,B1:B3)))', [30, 30]],
    ['SUM(MOD(XLOOKUP({1;2},A1:A3,B1:B3),{2;3}))', [2, 2]],
    ['SUM(MOD(XLOOKUP({1;2},A1:A3,B1:B3),3))', [3, 3]],
    ['SUM(IFERROR(XLOOKUP({9;1},A1:A3,B1:B3),{7;8}))', [17, 17]],
    ['SUM(IF({TRUE;FALSE},XLOOKUP({1;2},A1:A3,B1:B3),0))', [10, 20]],
    ['SUM(IF(XLOOKUP({1;2},A1:A3,B1:B3)>0,XLOOKUP({1;2},A1:A3,B1:B3),0))', [10, 20]],
    ['SUM(XLOOKUP({1;2},A1:A3,B1:B3)+SUM(A1:A3))', [42, 42]],
    ['SUM(XLOOKUP({0;4},A1:A3,B1:B3,,{1;-1}))', [10, 30]],
    ['SUM(XLOOKUP({9;1},A1:A3,B1:B3,{7;8}))', [15, 10]],
    ['LET(x,XLOOKUP({1;2},A1:A3,B1:B3),SUM(TRANSPOSE(x)))', [10, 20]],
    ['LET(x,XLOOKUP({1;2},A1:A3,B1:B3),SUM(x)+SUM(TRANSPOSE(x)))', [20, 40]],
    ['LET(x,XLOOKUP({1;2},A1:A3,B1:B3),SUM(LAMBDA(t,SUM(x))(0)))', [10, 20]],
    ['SUM(LAMBDA(t,t)(XLOOKUP({1;2},A1:A3,B1:B3)))', [10, 20]],
    ['SUM(LAMBDA(k,XLOOKUP(k,A1:A3,B1:B3))({1;2}))', [30, 30]],
    ['SUM(TRANSPOSE(XLOOKUP({1;2},A1:A3,B1:B3)))', [30, 30]],
  ] as const)('%s keeps its native CSE argument context', (formula, expected) => {
    expect(values({ formula, array: 'G1:G2' }).slice(0, 2)).toEqual(expected)
  })

  // test/excel/spill/cse-lookup-held*.json
  it.each([
    ['SUM(OFFSET(B1,{1;2},0))', 'G1:G2', [10, 20]],
    ['SUM(OFFSET(B1,{1;2},0))', 'G1', [10]],
    ['SUM(OFFSET(B1,0,0,{1;2}))', 'G1:G2', [30, 40]],
    ['SUM(INDIRECT({"B2";"B3"}))', 'G1:G2', [10, 20]],
    ['AVERAGE(XLOOKUP({1;2},A1:A3,B1:B3))', 'G1:G2', [10, 20]],
    ['ROWS(XLOOKUP({1;2},A1:A3,B1:B3))', 'G1:G2', [1, 1]],
    ['SUMPRODUCT(XLOOKUP({1;2},A1:A3,B1:B3))', 'G1:G2', [30, 30]],
    ['MODE(XLOOKUP({1;2},A1:A3,B1:B3),10)', 'G1:G2', [10, 10]],
    ['SUMX2PY2(XLOOKUP({1;2},A1:A3,B1:B3),{0;0})', 'G1:G2', [500, 500]],
    ['SUM(CHOOSE({1;2},B2,B3))', 'G1:G2', [30, 30]],
  ] as const)('%s over %s selects or aggregates as Excel does', (formula, array, expected) => {
    expect(values({ formula, array }).slice(0, expected.length)).toEqual(expected)
  })

  // test/excel/spill/cse-lookup-held-4.json, -6.json: the dereferencing idioms add every value, as in 0.89.0 and before
  it.each([
    ['SUM(INDEX(B1:B3,N(IF(1,{2;3}))))', 'G1', [30]],
    ['SUM(INDEX(B1:B3,N(IF(1,{2;3}))))', 'G1:G2', [30, 30]],
    ['SUM(INDEX(B1:B3,N(IF(TRUE,{2;3}))))', 'G1', [30]],
    ['SUM(XLOOKUP(N(IF(1,{1;2})),A1:A3,B1:B3))', 'G1', [30]],
    ['SUM(N(OFFSET(B1,{1;2},0)))', 'G1:G2', [30, 30]],
    ['SUM(N(INDIRECT({"B2";"B3"})))', 'G1', [30]],
    ['SUM(INDEX(B1:B3,IF(1,{2;3})))', 'G1', [10]],
    ['SUM(INDEX(B1:B3,N({2;3})))', 'G1', [10]],
  ] as const)('%s over %s keeps the old dereferencing idiom', (formula, array, expected) => {
    expect(values({ formula, array }).slice(0, expected.length)).toEqual(expected)
  })

  it.each([
    ['ROW(G1)', 'G1', [1]],
    ['ROWS(1:2)', 'G1', [2]],
    ['ROW(1:2)', 'G1:G2', [1, 2]],
    ['INDEX(B1:B3,SMALL(IF(A1:A3>1,ROW(A1:A3)),ROW(1:2)))', 'G1:G2', [30, 20]],
  ] as const)('%s over %s reads its own cells for their place alone, no circular reference', (formula, array, expected) => {
    expect(values({ formula, array }).slice(0, expected.length)).toEqual(expected)
  })

  it('pads positions past the key range with NA', () => {
    expect(values({ formula: 'SUM(XLOOKUP(D1:D2,A1:A3,B1:B3))', array: 'G1:G3' })).toEqual([10, 20, { error: 'NA' }])
  })

  it.each([
    ['SUM(XLOOKUP(D1:D2,A1:A3,B1:B3))', 'G5:G6', ['G5', 'G6'], [10, 20]],
    ['SUM(XLOOKUP(D1:D2,A1:A3,B1:B3))', 'G5', ['G5'], [10]],
    ['SUM(INDEX(B1:B3,D1:D2))', 'G5:G6', ['G5', 'G6'], [30, 10]],
  ] as const)('%s samples range selectors relative to the CSE area %s', (formula, area, addresses, expected) => {
    const b = book({ value: null })
    b.sheets[0].cells.G5 = { formula, array: area }
    const cells = compute(b, { execution, display: false }).sheets[0].cells
    expect(addresses.map((cell) => cells[cell].value)).toEqual(expected)
  })

  it('keeps SUM reference coercion for each selected INDEX cell', () => {
    const b = book({ formula: 'SUM(INDEX(B1:B7,SEQUENCE(7)))', array: 'G1:G7' })
    const input: CellInput[] = [{ value: 42 }, { value: '7' }, { value: true }, { value: null }, { formula: 'NA()' }, { value: 1e-308 }, { value: -0 }]
    input.forEach((cell, i) => { b.sheets[0].cells[`B${i + 1}`] = cell })
    const cells = compute(b, { execution, display: false }).sheets[0].cells
    expect(input.map((_, i) => cells[`G${i + 1}`].value)).toEqual([42, 0, 0, 0, { error: 'NA' }, 0, 0])
  })

  it('keeps the matched row or column as a whole reference before SUM', () => {
    const result = compute(book({ formula: 'SUM(XLOOKUP({1;2},A1:A3,B1:C3))', array: 'G1:G2' }), { execution }).sheets[0].cells
    expect([result.G1.value, result.G2.value]).toEqual([110, 220])
    const columns = compute(book({ formula: 'SUM(INDEX(B1:C3,0,{1,2}))', array: 'G1:H1' }), { execution }).sheets[0].cells
    expect([columns.G1.value, columns.H1.value]).toEqual([60, 600])
  })

  it('broadcasts row and column selectors independently', () => {
    const cells = compute(book({ formula: 'SUM(INDEX(B1:C3,{2;3},{1,2}))', array: 'G1:H2' }), { execution }).sheets[0].cells
    expect(['G1', 'H1', 'G2', 'H2'].map((a) => cells[a].value)).toEqual([10, 100, 20, 200])
  })

  it('keeps the scalar type of a singleton selector', () => {
    expect(values({ formula: 'TYPE(XLOOKUP({1},A1:A3,B1:B3))', array: 'G1' })[0]).toBe(1)
  })

  it.each([
    ['SUM(XLOOKUP({1;2},A1:A3,B1:B3)+1)', 11],
    ['SUM(ABS(XLOOKUP({1;2},A1:A3,B1:B3)))', 10],
    ['SUM(MOD(XLOOKUP({1;2},A1:A3,B1:B3),3))', 1],
    ['SUM(XLOOKUP({1;2},A1:A3,B1:B3)+SUM(A1:A3))', 16],
    ['SUM(XLOOKUP({1;2},A1:A3,B1:B3)+{7;8})', 35],
    ['SUM(TRANSPOSE(XLOOKUP({1;2},A1:A3,B1:B3)))', 10],
    ['SUM(IFERROR(XLOOKUP({9;1},A1:A3,B1:B3),{7;8}))', 15],
    ['SUM(LAMBDA(k,XLOOKUP(k,A1:A3,B1:B3))({1;2}))', 30],
  ] as const)('%s keeps the single-cell CSE context', (formula, expected) => {
    expect(values({ formula, array: 'G1' })[0]).toBe(expected)
  })

  it.each([
    ['SUM(XLOOKUP({1;2},A1:A3,B1:B3)+0)', [30, 30]],
    ['SUM(XLOOKUP({1;2},A1:A3,B1:B3)+1)', [32, 32]],
    ['SUM(XLOOKUP({1;2},A1:A3,B1:B3)+B1)', [90, 90]],
    ['SUM(XLOOKUP({1;2},A1:A3,B1:B3)+SUM(1,2,3))', [42, 42]],
    ['SUM(XLOOKUP({1;2},A1:A3,B1:B3)+IF(TRUE,6,7))', [42, 42]],
    ['SUM(IFERROR(XLOOKUP({1;2},A1:A3,B1:B3),0))', [30, 30]],
    ['SUM(N(XLOOKUP({1;2},A1:A3,B1:B3)))', [10, 20]],
  ] as const)('%s uses the multi-cell CSE context', (formula, expected) => {
    expect(values({ formula, array: 'G1:G2' }).slice(0, 2)).toEqual(expected)
  })

  it.each([
    ['SUM(XLOOKUP(1,A1:A3,XLOOKUP({3;1;2},A1:A3,B1:B3)))', 'VALUE'],
    ['SUM(INDEX(XLOOKUP({3;1;2},A1:A3,B1:B3),2))', 'REF'],
  ] as const)('%s retains the native single-cell table error', (formula, error) => {
    expect(values({ formula, array: 'G1' })[0]).toEqual({ error })
  })

  it.each([
    'LET(k,1,x,XLOOKUP({1;2},A1:A3,B1:B3),SUM(x)+SUM(k))',
    'LET(k,1,x,XLOOKUP({1;2},A1:A3,B1:B3),SUM(x)+SUM(LAMBDA(t,k)(0)))',
    'LET(n,2,k,SEQUENCE(n),SUM(XLOOKUP(k,A1:A3,B1:B3)))',
    'SUM(XLOOKUP(LET(n,2,SEQUENCE(n)),A1:A3,B1:B3))',
  ])('%s preserves lexical bindings in both backends', (formula) => {
    const b = book({ formula, array: 'G1:G2' })
    const reference = compute(b, { execution: 'closures', display: false })
    expect(compute(b, { execution, display: false }).sheets).toEqual(reference.sheets)
  })
})
