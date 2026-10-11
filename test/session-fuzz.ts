/**
 * Seeded random books and edits of the session fuzz test (`session-fuzz.test.ts`), shaped where the range caches
 * can go wrong (`docs/cache.md`): a data sheet of one column read through rectangles wider than it and through whole
 * columns, by formulas that also count cells edited often, so that cached ranges are read again after a write past
 * the sheet's edge grew it; spills that grow, shrink and are blocked; formulas replaced; rows inserted and deleted.
 */
import type { BookInput, CellEdit, CellInput, CellPatch, CellReplace, ClearKind, ComputeResult, RangeRef, Scalar, StructureOp } from '../src/contract.js'
import { indexToCol } from '../src/model.js'
import { mulberry32 } from '../src/synth.js'

/** Data rows of the main sheet `S` and of the data sheet `D`. */
const N = 300
/** The rows of `S!A` edited most: each edit there recomputes every formula over `D` (`+COUNT(A1:A60)`). */
const HOT = 60

export type Step = { edits: CellEdit[] } | { op: StructureOp }

export function fuzz(seed: number): { book: BookInput; step: () => Step } {
  const rnd = mulberry32(seed)
  const int = (n: number) => Math.floor(rnd() * n)
  const pick = <T>(xs: readonly T[]): T => xs[int(xs.length)]
  const value = (): Scalar => pick([() => int(50), () => -int(50) / 4, () => pick(['a', 'b', 'Б', 'x y']), () => rnd() < 0.5, () => int(5)])()
  const input = (): CellInput | null => (rnd() < 0.15 ? null : { value: value() })
  const k = () => int(50)
  const r = () => 1 + int(N)
  const hot = `COUNT(A1:A${HOT})`
  const formulas: (() => string)[] = [
    () => `SUM(A1:A${N})`, () => 'COUNT(A:A)', () => 'MIN(B:B)', () => `MAX(A1:B${N})`,
    () => `XLOOKUP(${k()},A1:A${N},B1:B${N},"none")`, () => `MATCH(${k()},A:A,0)`, () => `VLOOKUP(${k()},A1:B${N},2,FALSE)`,
    () => `XLOOKUP("${pick(['a', 'b', 'Б'])}",A:A,B:B,"none",0,-1)`, () => 'COUNTIF(A:A,">10")', () => `SUMIFS(B:B,A:A,"<${k()}")`,
    () => `INDEX(A1:D${N}*1,${r()},1)`, () => `VAR(A1:A${N})`, () => `SLOPE(B1:B${N},A1:A${N})`,
    () => 'SUM(E:E)+MAX(E:E)', () => `COUNT(F1:F${N + 200})`, () => 'SUM(G:G)+COUNTA(G:G)', () => `SUM(I:I)+COUNTA(I1:I${N + 50})`,
    // equality criteria and ranges one formula reads twice: kept from their second read
    () => `COUNTIF(A:A,${k()})+COUNTIF(A:A,${k()})`, () => `SUMIFS(B:B,A:A,${k()},B:B,"<>a")`,
    () => `AVERAGE(A1:A${N})+COUNTA(A1:A${N})+MATCH(${k()},A1:A${N},0)`,
    // the data sheet through rectangles wider than it and whole columns, recomputed by each edit of S!A1:A60
    () => `INDEX(D!A1:D${N}*1,${r()},1)+${hot}`, () => `SUM(D!A1:D${N}*1)+${hot}`, () => `SUMPRODUCT(D!A:B*1)+${hot}`,
    () => `COUNTA(D!A1:C${N})+${hot}`, () => `MATCH(${k()},D!A1:A${N},0)+${hot}`, () => `INDEX(D!A1:J${N}*1,${r()},1)+COUNTA(D!A1:J${N})`,
    () => `COUNTIFS(D!A:A,">${k()}")+${hot}`, () => `INDEX(D!A:C*1,${r()},1)+${hot}`, () => 'MAX(D!A:B)', () => 'COUNT(D!B:B)',
    // criteria over rectangles of the data sheet, a blank criterion counting past its rows
    () => `COUNTIFS(D!A:C,"")+${hot}`, () => `SUMIFS(D!A1:C${N},D!A1:C${N},"<>${k()}")+${hot}`,
  ]

  const cells: Record<string, CellInput> = {}
  for (let row = 1; row <= N; row++) {
    if (rnd() < 0.8) cells[`A${row}`] = { value: int(50) }
    if (rnd() < 0.6) cells[`B${row}`] = { value: value() }
  }
  // spills: one as tall as the data, one of a set size, one as tall as the count of the hot rows
  cells.E1 = { formula: `A1:A${N}*2` }
  cells.F1 = { formula: `SEQUENCE(${50 + int(200)})` }
  cells.G1 = { formula: `SEQUENCE(MAX(1,${hot}))` }
  formulas.forEach((f, i) => (cells[`H${i + 1}`] = { formula: f() }))
  const data: Record<string, CellInput> = {}
  for (let row = 1; row <= N; row++) if (rnd() < 0.7) data[`A${row}`] = { value: int(50) }
  const other: Record<string, CellInput> = {
    A1: { formula: 'SUM(S!A:A)+COUNT(S!E:E)' },
    A2: { formula: `XLOOKUP(7,S!A1:A${N},S!B1:B${N},"none")` },
    A3: { formula: `SUM(D!A1:D${N}*1)+S!A1` },
  }
  const book: BookInput = { sheets: [{ name: 'S', cells }, { name: 'D', cells: data }, { name: 'T', cells: other }] }

  // right of the data sheet's column: past the rectangles over A:D and A:C, inside the one over A:J
  const right = () => `${indexToCol(4 + int(8))}${r()}`
  const edit = (): CellEdit => {
    const x = rnd()
    if (x < 0.33) return { sheet: 'S', cell: `A${1 + int(HOT)}`, input: input() }
    if (x < 0.45) return { sheet: 'S', cell: `${pick(['A', 'B'])}${1 + int(N + 100)}`, input: input() }
    // past the data sheet's edge, right of it or below its rows: the sheet grows, the cached ranges the write misses stay
    if (x < 0.58) return { sheet: 'D', cell: rnd() < 0.5 ? right() : `${indexToCol(1 + int(9))}${N + 1 + int(2 * N)}`, input: input() }
    if (x < 0.64) return { sheet: 'D', cell: `${indexToCol(int(4))}${r()}`, input: input() }
    // a value in a spill's way blocks it (#SPILL!), clearing it lets the spill back
    if (x < 0.74) return { sheet: 'S', cell: `${pick(['E', 'F', 'G', 'I'])}${1 + int(N)}`, input: rnd() < 0.5 ? { value: 'block' } : null }
    // a new spill on S, where values and other spills may block it, or past the data sheet's edge; some as tall as the hot rows' count
    if (x < 0.8) return { sheet: pick(['S', 'D']), cell: rnd() < 0.5 ? `I${r()}` : right(), input: { formula: pick([`SEQUENCE(${1 + int(30)})`, `FILTER(S!A1:A${HOT},S!A1:A${HOT}>${k()})`, `S!A1:A${HOT}*1`]) } }
    if (x < 0.92) return { sheet: 'S', cell: `H${1 + int(formulas.length)}`, input: { formula: pick(formulas)() } }
    return { sheet: 'S', cell: `${indexToCol(9 + int(8))}${N + 200 + int(400)}`, input: { value: int(9) } }
  }
  const step = (): Step => {
    const x = rnd()
    if (x < 0.9) return { edits: [edit()] }
    if (x < 0.97) return { edits: Array.from({ length: 2 + int(2) }, edit) }
    return { op: pick<StructureOp>([{ op: 'insertRows', sheet: 'S', row: r(), count: 1 + int(3) }, { op: 'deleteRows', sheet: 'S', row: 2 + int(N), count: 1 }, { op: 'insertRows', sheet: 'D', row: r(), count: 1 }]) }
  }
  return { book, step }
}

/**
 * Seeded patches of cells' looks (`Session.patchCells`), from a stream of their own so that the steps of `fuzz` stay
 * as they were: on spill ranges, values and empty cells of `S`, on cells of `D`, and far past the data, where they move
 * a sheet's extent and are removed again; a patch now and then holds several of them.
 */
export function patches(seed: number): () => CellPatch[] {
  const rnd = mulberry32(seed * 7919 + 1)
  const int = (n: number) => Math.floor(rnd() * n)
  const pick = <T>(xs: readonly T[]): T => xs[int(xs.length)]
  const far: { sheet: string; cell: string }[] = []
  const set = (): CellPatch['set'] => {
    const out: CellPatch['set'] = {}
    if (rnd() < 0.5) out.numberFormat = pick(['0.00', '0%', '@', null])
    if (rnd() < 0.3) out.meta = pick([int(9), 'm', null])
    if (rnd() < 0.3) out.note = rnd() < 0.7 ? { text: `n${int(9)}` } : null
    if (rnd() < 0.2) out.link = rnd() < 0.7 ? { url: 'https://example.com/' } : null
    return out
  }
  const one = (): CellPatch => {
    const x = rnd()
    if (x < 0.35) return { sheet: 'S', cell: `${pick(['E', 'F', 'G', 'I'])}${1 + int(N)}`, set: set() }
    if (x < 0.6) return { sheet: 'S', cell: `${pick(['A', 'B', 'H'])}${1 + int(N + 100)}`, set: set() }
    if (x < 0.7) return { sheet: 'D', cell: `${indexToCol(int(10))}${1 + int(2 * N)}`, set: set() }
    // far past the data: the extent grows, and shrinks when the cell is left with nothing
    if (x < 0.85 || !far.length) {
      const at = { sheet: pick(['S', 'D', 'T']), cell: `${indexToCol(10 + int(20))}${N + 300 + int(3000)}` }
      far.push(at)
      return { ...at, set: { numberFormat: '0.0' } }
    }
    return { ...far.splice(int(far.length), 1)[0], set: { numberFormat: null, meta: null, note: null, link: null } }
  }
  return () => Array.from({ length: rnd() < 0.8 ? 1 : 2 + int(3) }, one)
}

/** An edit of the calls that keep a cell's look or set it: cells replaced whole, cleared, formatted by a style's fields, text typed. */
export type LookEdit =
  | { replace: CellReplace[] }
  | { clear: RangeRef[]; what: ClearKind }
  | { patch: CellPatch[] }
  | { text: [sheet: string, cell: string, text: string] }

/**
 * Seeded edits of the calls that keep or set a cell's look (`replaceCells`, `clearCells`, `patchCells` over ranges with a
 * style by its fields, `setCellText`), from a stream of their own: over the data, spills and formulas of `S` and `D`,
 * rectangles and whole columns and rows, text read as numbers, dates, percents and formulas, some of them refused.
 */
export function lookEdits(seed: number): () => LookEdit {
  const rnd = mulberry32(seed * 4093 + 5)
  const int = (n: number) => Math.floor(rnd() * n)
  const pick = <T>(xs: readonly T[]): T => xs[int(xs.length)]
  const cell = () => `${pick(['A', 'B', 'E', 'H', 'I', 'K'])}${1 + int(N + 20)}`
  const rect = () => {
    const [c, r] = [int(10), 1 + int(N)]
    return `${indexToCol(c)}${r}:${indexToCol(c + int(3))}${r + int(5)}`
  }
  const lines = () => pick([`${indexToCol(9 + int(5))}:${indexToCol(9 + int(8))}`, `${N + 10 + int(20)}:${N + 30 + int(20)}`, 'B:B', `${1 + int(N)}:${1 + int(N)}`])
  const style = () => pick([{ font: { bold: true } }, { font: { italic: pick([true, null]) } }, { fill: { color: '#FFEE99' } }, { fill: null, font: { bold: null } }, { border: { bottom: { style: 'thin' as const } } }])
  return () => {
    const x = rnd()
    if (x < 0.2) return { replace: [{ sheet: pick(['S', 'D']), cell: cell(), input: rnd() < 0.3 ? null : { value: int(50), ...(rnd() < 0.5 ? { numberFormat: pick(['0.00', '0%']) } : {}) } }] }
    if (x < 0.4) return { clear: [{ sheet: pick(['S', 'D']), range: rnd() < 0.7 ? rect() : lines() }], what: pick(['contents', 'formats', 'all'] as const) }
    if (x < 0.7) {
      const range = rnd() < 0.6 ? rect() : lines()
      return { patch: [{ sheet: pick(['S', 'D']), range, set: { ...(rnd() < 0.8 ? { style: style() } : {}), ...(rnd() < 0.4 ? { numberFormat: pick(['0.0', '#,##0', 'yyyy-mm-dd', null]) } : {}) } }] }
    }
    return { text: [pick(['S', 'D']), cell(), pick([String(int(50)), '-3,5', '1/2/2026', '15%', '=A1*2', `=SUM(A1:A${1 + int(N)})`, '', 'text', '=SUM(A1', '1e3'])] }
  }
}

/** The cells whose values differ between a session's result and a full computation's, with their formulas. */
export function differences(session: ComputeResult, full: ComputeResult): string[] {
  const out: string[] = []
  full.sheets.forEach((f, i) => {
    const s = session.sheets[i]?.cells ?? {}
    for (const a of new Set([...Object.keys(f.cells), ...Object.keys(s)])) {
      const [x, y] = [JSON.stringify(s[a]?.value), JSON.stringify(f.cells[a]?.value)]
      if (x !== y) out.push(`${f.name}!${a} ${f.cells[a]?.formula ?? s[a]?.formula ?? ''}: session ${x}, full ${y}`)
    }
  })
  return out
}
