/**
 * The engine's prediction of a set of GETPIVOTDATA formulas (`gpd-1`), written before Excel's answers are read: each
 * case's book (the saved book of `from`, its values put and its formulas typed as the harness types them) computed,
 * and the cells it watches. A formula over a pivot table in another book is predicted at each read of the harness:
 * the engine does not open the other book, so a computation gives what it gives without it, and the file opened again
 * keeps the value Excel saved (`held`). Run: `npx tsx bench/coverage/pivot/gpd-predict.ts gpd-1`; the cases' book
 * and answers: `pivotGpdBook`, `test/pivot-gpd.test.ts`.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { compute, type BookInput, type CellValue } from '../../../src/index.js'
import { readXlsx } from '../../../src/xlsx.js'
import { formatAddress } from '../../../src/model.js'
import { parseRange } from '../../../src/structure/area.js'
import type { PivotCase, PivotSet } from './cases.js'

const ROOT = new URL('../../../', import.meta.url)

/**
 * The book of a case: its `from` book as Excel saved it, with the values and formulas of its operations; a case that
 * builds its PivotTable (`gpdOver`), the book Excel saved for it in the set (`set`), its formulas computed anew, the
 * values it stores for them not read; null for one over another book.
 */
export async function pivotGpdBook(c: PivotCase, set?: string, locale?: string): Promise<BookInput | null> {
  if (c.ops.some((o) => 'external' in o) || (!c.from && !set)) return null
  // a set of Excel ru-RU reads the books of its own run (`ru-RU/basic/…`), in ru-RU
  const dir = locale ? `${locale}/` : ''
  const read = (await readXlsx(readFileSync(new URL(`test/excel/win/pivot/${dir}${c.from ?? `${set}/${c.id}`}.xlsx`, ROOT)))).book
  const book = locale ? { ...read, locale } : read
  for (const o of c.ops) {
    if (!('edit' in o)) continue
    const [sheet, cell] = o.edit.split('!')
    const s = book.sheets.find((x) => x.name === sheet)!
    if (s.pivotTables && typeof o.to === 'string') renameItem(book, s, cell, o.to)
    // a value typed as Excel reads it: text of a number a number
    const typed = typeof o.to === 'string' && o.to.trim() !== '' && Number.isFinite(Number(o.to)) ? Number(o.to) : o.to
    s.cells[cell] = typeof o.to === 'object' && o.to !== null && 'formula' in o.to ? { formula: o.to.formula } : { value: typed as string | number }
  }
  return book
}

/**
 * Text typed over an item's label in a PivotTable renames the item, as Excel's entry does (the set edit): the item of a
 * field on the rows or columns whose label is the cell's gives its caption.
 */
function renameItem(book: BookInput, sheet: BookInput['sheets'][number], cell: string, to: string) {
  const was = sheet.cells[cell]?.value
  for (const t of sheet.pivotTables ?? []) {
    const c = book.pivotCaches!.find((x) => x.id === t.cache)!
    for (const f of [...(t.rows ?? []), ...(t.columns ?? [])]) {
      if (f === 'values') continue
      const it = t.fields[f].items?.find((i) => (i.caption ?? String(c.fields[f].items?.[i.item!] ?? '(blank)')) === String(was))
      if (it) it.caption = to
    }
  }
}

/** The cells of a case's watched ranges (`G!A1:A9`), as `G!A1`. */
export function watched(c: PivotCase): string[] {
  return (c.watch ?? []).flatMap((w) => {
    const [sheet, range] = w.split('!')
    const x = parseRange(range)!
    const out: string[] = []
    for (let r = x.r1; r <= x.r2; r++) for (let col = x.c1; col <= x.c2; col++) out.push(`${sheet}!${formatAddress(r, col)}`)
    return out
  })
}

/** The engine's values of a case's watched cells. */
export async function predictCase(c: PivotCase, set?: string, locale?: string): Promise<Record<string, CellValue>> {
  const book = await pivotGpdBook(c, set, locale)
  if (!book) return {}
  const r = compute(book)
  const out: Record<string, CellValue> = {}
  for (const k of watched(c)) {
    const [sheet, cell] = k.split('!')
    out[k] = r.sheets[book.sheets.findIndex((x) => x.name === sheet)].cells[cell]?.value ?? null
  }
  return out
}

/** A formula over another book: computed without it, #REF! (`closed, computed`, `reopened, computed`); the file opened again keeps the value Excel saved (`reopened`, as `open`). */
function predictExternal(c: PivotCase): Record<string, string> {
  const o = c.ops.find((x) => 'external' in x) as { external: string; cells: Record<string, string> }
  const book: BookInput = { sheets: [{ name: 'G', cells: {} }], externalBooks: [{ target: 'Src.xlsx', sheets: [{ name: 'P', cells: {} }] }] }
  const cells = Object.keys(o.cells)
  for (const k of cells) book.sheets[0].cells[k.split('!')[1]] = { formula: o.cells[k].slice(1).replace('[Src.xlsx]', '[1]') }
  const r = compute(book)
  const held = compute(book, { cached: { sheets: [{ name: 'G', cells: Object.fromEntries(cells.map((k) => [k.split('!')[1], { value: 7 }])) }] } })
  return Object.fromEntries(
    cells.flatMap((k) => {
      const a = k.split('!')[1]
      const v = r.sheets[0].cells[a].value
      return [
        [`closed, computed ${k}`, JSON.stringify(v)],
        [`reopened, computed ${k}`, JSON.stringify(v)],
        [`reopened ${k}`, held.sheets[0].cells[a].value === 7 ? 'as open' : JSON.stringify(held.sheets[0].cells[a].value)],
      ]
    }),
  )
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const set = process.argv[2]
  const { buildPivotCases } = (await import(`./${set}.ts`)) as PivotSet
  const out: Record<string, Record<string, unknown>> = {}
  for (const c of buildPivotCases()) out[c.id] = c.ops.some((o) => 'external' in o) ? predictExternal(c) : await predictCase(c, set)
  writeFileSync(new URL(`bench/coverage/pivot/${set}.predicted.json`, ROOT), JSON.stringify(out, null, 1) + '\n')
  console.log(Object.entries(out).map(([id, v]) => `${id}: ${Object.keys(v).length}`).join('\n'))
}
