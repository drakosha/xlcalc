/**
 * Formulas without their prefix repaired by `writeKept` (`prefixes: 'repair'`), against Excel: the files openpyxl wrote
 * (`prefixes.py`, kept in `test/excel/keep/prefixes-made/`), each sent as it is (`<book>-original`) and as `writeKept`
 * writes it over itself with the option, nothing else changed (`<book>-repaired`). `keep prefixes-predict` writes
 * what Excel is predicted to do with each (`prefixes.predicted.json`), before `keep prefixes` has Excel for Mac open,
 * compute and save them (`test/excel/keep/prefixes.json`, `prefixes/<case>.xlsx`) and `bench/windows/run.ts keep
 * prefixes` Excel for Windows (`test/excel/win/keep/`). Test: `test/xlsx-no-prefix-repair.test.ts`.
 *
 * The prediction: Excel opens every repaired file without its repair window, and each cell read shows the engine's
 * value, the same before the repair and after it; the formulas Excel saves read as those of the file repaired, of the old
 * kind where they were (`@FILTER(…)`, an array formula of the old kind over its range), so none spills. Of the
 * originals, hand-written from `links/no-prefix.ts`: Excel refuses those with FILTER or SORT in a cell, and in the
 * others shows `#NAME?` for a call without its prefix, a defined name's too, but where `IFERROR` gives its fallback.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { compute } from '../../../src/index.js'
import type { CellValue } from '../../../src/contract.js'
import { openKept, writeKept } from '../../../src/xlsx-keep.js'
import { readXlsx } from '../../../src/xlsx.js'
import { ROOT } from '../paths.js'

/** A value as the answers hold it: `{ error }` for an error, as a result has it. */
export type Shown = CellValue | { error: string } | null

/** A book of `prefixes.py`: the cells read, sheet and address (the first sheet's without a sheet). */
export interface PrefixBook {
  id: string
  read: string[]
  /** What Excel shows in the original file, by cell; `refused` where it does not open it. */
  original: { refused: true } | { values: Record<string, Shown> }
  /** Excel for Windows refuses the original, which Excel for Mac opens: a LAMBDA's parameters `r` and `c` without their prefix read as no names there (`links/no-prefix.ts`, MAKEARRAY). */
  refusedOnWindows?: true
}

const NAME = { error: 'NAME' }

export const PREFIX_BOOKS: PrefixBook[] = [
  { id: 'lookup', read: ['C1', 'C2', 'C3', 'D1'], original: { values: { C1: NAME, C2: 'fallback', C3: NAME, D1: 1 } } },
  { id: 'filter', read: ['C1', 'C2', 'C3', 'C4'], original: { refused: true } },
  { id: 'sort', read: ['C2', 'C3', 'C4'], original: { refused: true } },
  { id: 'shared', read: ['E1', 'E2', 'E3', 'E4', 'F1', 'F2', 'F3', 'F4'], original: { values: { E1: NAME, E2: NAME, E3: NAME, E4: NAME, F1: 'none', F2: 'none', F3: 'none', F4: 'none' } } },
  { id: 'array', read: ['C1', 'C2', 'C3', 'E1', 'E2', 'E3', 'G1'], original: { refused: true } },
  { id: 'name', read: ['C1', 'C2', 'C3'], original: { values: { C1: NAME, C2: NAME, C3: NAME } } },
  { id: 'mixed', read: ['C1', 'C2', 'C3', 'C4', 'C5'], original: { values: { C1: NAME, C2: NAME, C3: NAME, C4: NAME, C5: 5 } }, refusedOnWindows: true },
  { id: 'rules', read: ['C1'], original: { values: { C1: 1 } } },
  { id: 'table', read: ['B2', 'B3'], original: { values: { B2: NAME, B3: NAME } } },
  { id: 'pivot', read: ['E2', 'G1', 'G2'], original: { values: { E2: 1610, G1: NAME, G2: NAME } } },
]

export const prefixMade = (id: string) => new Uint8Array(readFileSync(join(ROOT, 'test/excel/keep/prefixes-made', `${id}.xlsx`)))

/** A book's file repaired: written over itself by `writeKept` with `prefixes: 'repair'`, the engine's values with it, as a host writes it. */
export async function repaired(original: Uint8Array): Promise<Uint8Array> {
  const { book, cached } = await openKept(original)
  return (await writeKept(original, { book, result: compute(book, { cached }) }, { prefixes: 'repair' })).data
}

/** The files sent: each book's original and its repair. */
export async function prefixFiles(): Promise<{ id: string; data: Uint8Array }[]> {
  const out: { id: string; data: Uint8Array }[] = []
  for (const b of PREFIX_BOOKS) {
    const original = prefixMade(b.id)
    out.push({ id: `${b.id}-original`, data: original }, { id: `${b.id}-repaired`, data: await repaired(original) })
  }
  return out
}

const shown = (v: CellValue | undefined | null): Shown => (v === undefined ? null : v)

/** The first sheet's cells read in a file: the values it stores (Excel's, in a file Excel saved) and their formulas as the reader reads them. */
export async function readCells(data: Uint8Array, cells: string[]): Promise<{ values: Record<string, Shown>; formulas: Record<string, string | null> }> {
  const r = await readXlsx(data)
  const s = r.book.sheets[0]
  const c = r.cached.sheets[0].cells
  return {
    values: Object.fromEntries(cells.map((a) => [a, shown(c[a]?.value)])),
    formulas: Object.fromEntries(cells.map((a) => [a, s.cells[a]?.formula ?? null])),
  }
}

/** What the engine gives the cells of a file read, `compute` over it. */
export async function engineCells(data: Uint8Array, cells: string[]): Promise<Record<string, Shown>> {
  const r = await readXlsx(data)
  const s = compute(r.book, { cached: r.cached }).sheets[0].cells
  return Object.fromEntries(cells.map((a) => [a, shown(s[a]?.value)]))
}

/** The prediction of each case: whether Excel opens the file, the values it shows and, for a repaired one, the formulas its saved file reads as. */
export async function prefixPredictions(): Promise<Record<string, { opens: boolean; opensOnWindows?: false; values?: Record<string, Shown>; formulas?: Record<string, string | null> }>> {
  const out: Awaited<ReturnType<typeof prefixPredictions>> = {}
  for (const b of PREFIX_BOOKS) {
    const original = prefixMade(b.id)
    out[`${b.id}-original`] = 'refused' in b.original ? { opens: false } : { opens: true, ...(b.refusedOnWindows ? { opensOnWindows: false as const } : {}), values: b.original.values }
    const data = await repaired(original)
    const before = await engineCells(original, b.read)
    const after = await engineCells(data, b.read)
    if (JSON.stringify(before) !== JSON.stringify(after)) throw new Error(`${b.id}: the repair changed the engine's values: ${JSON.stringify(before)} → ${JSON.stringify(after)}`)
    out[`${b.id}-repaired`] = { opens: true, values: after, formulas: (await readCells(data, b.read)).formulas }
  }
  return out
}
