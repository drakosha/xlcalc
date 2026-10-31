/**
 * What each cell of a PivotTable is (`pivotCells`) against Excel for Windows' `Range.PivotCell` over the books Excel
 * saved for other sets (the sets `cells` and `cells-held`, answers in `test/excel/win/pivot/`), each read with
 * `readXlsx` as `bench/coverage/pivot/cells-predict.ts` reads it. Compared by Excel's terms: the `PivotCellType` of
 * every cell of `TableRange2`, and by its type the field (`SourceName`), the item's label, the value field's name and
 * the items of a value's row and column.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { BookInput, PivotCell, PivotItemRef } from '../src/contract.js'
import { labeler } from '../src/pivot/layout.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { predictedCells } from '../bench/coverage/pivot/cells-predict.js'
import type { PivotAnswer, PivotCase, PivotCellKind, PivotSet } from '../bench/coverage/pivot/cases.js'

const SETS = ['cells', 'cells-held']

/** `PivotCellType` by the library's kind. */
const TYPE: Record<PivotCell['kind'], number> = { value: 0, item: 1, subtotal: 2, grand: 3, valueName: 4, field: 5, pageItem: 6, valuesField: 8, blank: 9 }

/** What a cell is in Excel's terms, by its type: [type, field, item, value field's name, row items, column items]. */
type Seen = [number, string?, string?, string?, string[]?, string[]?]

/** Excel's answer for a cell in the terms compared: a list Excel could not read (`[["",""]]`) is none. */
function excelSeen(k: PivotCellKind): Seen | null {
  const [, type, field, item, data, rows, columns] = k
  if (type === undefined) return null
  const pairs = (xs: [string, string][] | null | undefined) => (xs && !(xs.length === 1 && xs[0][0] === '') ? xs.map(([f, i]) => `${f}=${i}`) : [])
  switch (type) {
    case 0:
      return [0, undefined, undefined, data ?? '', pairs(rows), pairs(columns)]
    case 1:
    case 2:
    case 7:
      return [type === 7 ? 2 : type, field ?? '', item ?? '']
    case 3:
    case 4:
      return [type, undefined, undefined, data ?? '']
    case 5:
    case 6:
      return [type, field ?? '']
    default:
      return [type]
  }
}

/** The library's answer for a cell in Excel's terms. */
function oursSeen(book: BookInput, sheet: string, c: PivotCell): Seen {
  const t = book.sheets.find((s) => s.name === sheet)!.pivotTables!.find((p) => p.name === c.pivot)!
  const cache = book.pivotCaches!.find((x) => x.id === t.cache)!
  const { labelText, valueName } = labeler(t, cache, book)
  const fieldName = (f: PivotCell['field']) => (f === undefined || f === 'values' ? '' : cache.fields[f].name)
  const pairs = (xs: PivotItemRef[] | undefined) => (xs ?? []).map((x) => `${fieldName(x.field)}=${labelText(x.field, x.item)}`)
  const type = TYPE[c.kind]
  switch (c.kind) {
    case 'value':
      return [0, undefined, undefined, c.value === undefined ? '' : valueName(c.value), pairs(c.rowItems), pairs(c.columnItems)]
    case 'item':
    case 'subtotal':
      return [type, fieldName(c.field), labelText(c.field!, c.item!)]
    case 'grand':
    case 'valueName':
      return [type, undefined, undefined, c.value === undefined ? '' : valueName(c.value)]
    case 'field':
    case 'pageItem':
      return [type, fieldName(c.field)]
    default:
      return [type]
  }
}

/** The cells of each sheet a book's file gives Excel's filter button (`pivotButton` on their style), by sheet name. */
async function fileButtons(bytes: Uint8Array): Promise<Record<string, string[]>> {
  const a = openArchive(bytes)
  const xfs = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec((await a.text('xl/styles.xml')) ?? '')?.[1] ?? ''
  const buttons = new Set([...xfs.matchAll(/<xf\b[^>]*>/g)].flatMap((m, i) => (/pivotButton="1"/.test(m[0]) ? [String(i)] : [])))
  const book = (await a.text('xl/workbook.xml')) ?? ''
  const rels = (await a.text('xl/_rels/workbook.xml.rels')) ?? ''
  const target = (id: string) => new RegExp(`<Relationship[^>]*Id="${id}"[^>]*Target="([^"]+)"`).exec(rels)?.[1] ?? new RegExp(`<Relationship[^>]*Target="([^"]+)"[^>]*Id="${id}"`).exec(rels)?.[1]
  const out: Record<string, string[]> = {}
  for (const m of book.matchAll(/<sheet\b[^>]*name="([^"]*)"[^>]*r:id="([^"]+)"/g)) {
    const xml = (await a.text(`xl/${target(m[2])!.replace(/^\/?xl\//, '')}`)) ?? ''
    out[m[1]] = [...xml.matchAll(/<c r="([A-Z]+\d+)"[^>]*?\ss="(\d+)"/g)].flatMap((c) => (buttons.has(c[2]) ? [c[1]] : []))
  }
  return out
}

describe.each(SETS)('what the cells of the pivot tables of the set %s are', (set) => {
  const answers = JSON.parse(readFileSync(new URL(`excel/win/pivot/${set}.json`, import.meta.url), 'utf8')) as { answers: Record<string, PivotAnswer> }
  it("is what Excel's PivotCell says", async () => {
    const { buildPivotCases } = (await import(`../bench/coverage/pivot/${set}.ts`)) as PivotSet
    const misses: string[] = []
    let compared = 0
    for (const c of buildPivotCases() as PivotCase[]) {
      const book = (await readXlsx(readFileSync(new URL(`../${c.file}`, import.meta.url)))).book
      const ours = await predictedCells(c)
      for (const [key, got] of Object.entries(answers.answers[c.id].phases.at(-1)!.kinds ?? {})) {
        if ('error' in got) {
          misses.push(`${c.id} ${key}: Excel ${got.error}`)
          continue
        }
        const [sheet] = key.split('!')
        for (const k of got) {
          const want = excelSeen(k)
          const cell = ours[key]?.cells[k[0]]
          const have = cell ? oursSeen(book, sheet, cell) : null
          compared++
          if (JSON.stringify(have) !== JSON.stringify(want)) misses.push(`${c.id} ${key} ${k[0]}: ours ${JSON.stringify(have)}, Excel ${JSON.stringify(want)}`)
        }
      }
    }
    expect(misses).toEqual([])
    expect(compared).toBeGreaterThan(1000)
  }, 120_000)

  it("gives the filter buttons Excel's file marks (pivotButton)", async () => {
    const { buildPivotCases } = (await import(`../bench/coverage/pivot/${set}.ts`)) as PivotSet
    const misses: string[] = []
    for (const c of buildPivotCases() as PivotCase[]) {
      const bytes = new Uint8Array(readFileSync(new URL(`../${c.file}`, import.meta.url)))
      const file = await fileButtons(bytes)
      const ours = await predictedCells(c)
      const bySheet: Record<string, string[]> = {}
      for (const [key, cells] of Object.entries(ours)) (bySheet[key.split('!')[0]] ??= []).push(...cells.buttons)
      for (const [sheet, want] of Object.entries(file)) {
        const got = (bySheet[sheet] ?? []).slice().sort()
        if (JSON.stringify(got) !== JSON.stringify(want.slice().sort())) misses.push(`${c.id} ${sheet}: ours ${got.join(' ')}, file ${want.slice().sort().join(' ')}`)
      }
    }
    expect(misses).toEqual([])
  }, 120_000)

  // Excel opening a file without them draws none until a refresh (Excel for Windows saved each such book without, and with them after RefreshTable)
  it("writes the filter buttons where Excel's file has them, in writeXlsx and in writeKept over the sheet written again", async () => {
    const { buildPivotCases } = (await import(`../bench/coverage/pivot/${set}.ts`)) as PivotSet
    const misses: string[] = []
    for (const c of buildPivotCases() as PivotCase[]) {
      const bytes = new Uint8Array(readFileSync(new URL(`../${c.file}`, import.meta.url)))
      const want = await fileButtons(bytes)
      const book = (await readXlsx(bytes)).book
      // writeKept over a sheet written again: a cell beside the PivotTable on its sheet
      const kept = (await openKept(bytes)).book
      const edited = { ...kept, sheets: kept.sheets.map((s) => (s.pivotTables?.length ? { ...s, cells: { ...s.cells, Z99: { value: 1 } } } : s)) }
      for (const [how, written] of [['writeXlsx', (await writeXlsx({ book })).data], ['writeKept', (await writeKept(bytes, { book: edited })).data]] as const) {
        const got = await fileButtons(written)
        for (const [sheet, cells] of Object.entries(want))
          if (JSON.stringify((got[sheet] ?? []).slice().sort()) !== JSON.stringify(cells.slice().sort())) misses.push(`${how} ${c.id} ${sheet}: written ${(got[sheet] ?? []).join(' ')}, Excel's ${cells.join(' ')}`)
      }
    }
    expect(misses).toEqual([])
  }, 240_000)
})
