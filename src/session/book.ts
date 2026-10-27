/** Changes of a session's book that recompute nothing: styles, formats, meta, notes and links of cells, the layout of sheets, their data validation. */
import type { BookInput, BookLayout, CellInput, CellReplace, ColumnInput, ColumnsFit, LayoutEdit, RowInput, RowsFit, SheetInput, SheetLayout } from '../contract.js'
import { setInput, settled, sheetIndexOf, type EngineState } from '../engine.js'
import { InputError } from '../input-error.js'
import { hasLooks, inputIn, layerOf } from '../looks/sheet.js'
import { lines, parseSpan, patchIn, toMap, type Line } from '../layout/spans.js'
import { bookScalar, parseAddress } from '../model.js'
import { parseMerge } from '../layout/merges.js'
import { meets, parseRange, type Rect } from '../structure/area.js'
import { autoFilterOverlap } from '../tables/check.js'
import { tablesOf } from '../tables/model.js'

/** A scalar as JSON writes it: a number's -0 as 0, NaN and the infinities as null. */
const json = (v: unknown) => (typeof v === 'number' ? (Number.isFinite(v) ? v + 0 : null) : v)

/** Whether two values are written alike in JSON: scalars compared without writing them. */
const alike = (a: unknown, b: unknown) => (typeof a === 'object' && a !== null) || (typeof b === 'object' && b !== null) ? JSON.stringify(a) === JSON.stringify(b) : json(a) === json(b)

/** Whether two inputs compute from the same: their formula, array area and the value it holds (`held.ts`), or their value (a date as its serial number). */
function sameContent(a: CellInput | null | undefined, b: CellInput | null | undefined, epoch: number): boolean {
  if (a?.formula !== undefined) return b?.formula !== undefined && a.formula === b.formula && a.array === b.array && alike(a.held ?? null, b.held ?? null)
  return b?.formula === undefined && alike(bookScalar(a?.value, epoch), bookScalar(b?.value, epoch))
}

/**
 * The edits that change a cell's value or formula, the last one per cell; the others (a new style, format or
 * meta) are applied to the book's input here.
 */
export function contentEdits(s: EngineState, edits: CellReplace[]): CellReplace[] {
  const last = new Map<string, [CellReplace, number]>()
  for (const e of edits) {
    const i = sheetIndexOf(s, e.sheet)!
    last.set(`${i}:${e.cell}`, [e, i])
  }
  const out: CellReplace[] = []
  const quiet: CellReplace[] = []
  for (const [e, i] of last.values()) {
    if (!sameContent(s.input.sheets[i].cells[e.cell], e.input, s.ctx.rules.epoch)) out.push(e)
    else {
      const { row, col } = parseAddress(e.cell)
      setInput(s, i, e.cell, row, col, e.input)
      quiet.push(e)
    }
  }
  settled(s, quiet)
  return out
}

/** The inputs of cells now (null: none), a look of the sheet's looks as an input of a look alone, once per cell: the edits that put them back. */
export function inputsOf(s: EngineState, refs: { sheet: string; cell: string }[]): CellReplace[] {
  const out = new Map<string, CellReplace>()
  for (const { sheet, cell } of refs) {
    const i = sheetIndexOf(s, sheet)!
    const k = `${i}!${cell}`
    if (!out.has(k)) out.set(k, { sheet, cell, input: inputAt(s, i, cell) ?? null })
  }
  return [...out.values()]
}

/** A cell's input, or its look from the sheet's looks as an input of a look alone (`inputIn`). */
export function inputAt(s: EngineState, i: number, cell: string): CellInput | undefined {
  const sheet = s.input.sheets[i]
  const input = sheet.cells[cell]
  if (input || !hasLooks(sheet)) return input
  const { row, col } = parseAddress(cell)
  return inputIn(sheet, cell, row, col)
}

/** The session's book without its cells and their looks (`Session.getLayout`): new objects for the book and its sheets, the rest the book's own; the looks are not read. */
export function bookLayout(book: BookInput): BookLayout {
  return { ...book, sheets: book.sheets.map((s) => Object.fromEntries(Object.keys(s).filter((k) => k !== 'cells' && k !== 'looks').map((k) => [k, s[k as keyof SheetInput]])) as unknown as SheetLayout) }
}

/** Every style that cells, rows and columns use must be in a table of `count` styles. */
export function checkStylesUsed(book: BookInput, count: number) {
  const check = (style: number | undefined, where: string) => {
    if (style !== undefined && style >= count) throw new InputError('styles', `${where} uses style ${style}, but the new style table has ${count}`)
  }
  for (const s of book.sheets) {
    for (const [addr, c] of Object.entries(s.cells)) check(c.style, `${s.name}!${addr}`)
    for (const x of layerOf(s)?.used() ?? []) check(x.style, `the formatting of ${s.name}`)
    for (const [k, c] of Object.entries(s.columns ?? {})) check(c.style, `columns ${k} of ${s.name}`)
    for (const [k, r] of Object.entries(s.rows ?? {})) check(r.style, `rows ${k} of ${s.name}`)
  }
}

/** Merged areas and AutoFilters of checked layout edits refused over a table, as `checkBook` refuses them in a book. */
export function checkOutsideTables(book: BookInput, edits: LayoutEdit[], sheetIndex: (name: string) => number): void {
  const tables = tablesOf(book).list
  if (!tables.length) return
  edits.forEach((e, i) => {
    const s = sheetIndex(e.sheet)
    const over = (x: Rect) => tables.find((t) => t.s === s && meets(x, t))
    if ('autoFilter' in e && e.autoFilter) {
      const t = over(parseRange(e.autoFilter.range)!)
      if (t) throw new InputError(`edits[${i}].autoFilter.range`, autoFilterOverlap(e.autoFilter.range, t.name))
    }
    if (!('merges' in e)) return
    e.merges.forEach((m, j) => {
      const t = over(parseMerge(m)!)
      if (t) throw new InputError(`edits[${i}].merges[${j}]`, `the merged area ${m} overlaps the table ${t.name} (Excel does not allow merged cells in tables)`)
    })
  })
}

/** A layout edit that sets what it gives, AutoFit turned into the sizes it sets (`fitEdits`). */
export type SetLayoutEdit = Exclude<LayoutEdit, ColumnsFit | RowsFit>

/** Layout edits (checked) applied to the book: each changed sheet is a new object with the same cells; the changed sheets by index. */
export function applyLayout(book: BookInput, edits: SetLayoutEdit[], sheetIndex: (name: string) => number): Map<number, SheetInput> {
  const changed = new Map<number, SheetInput>()
  // each sheet's columns and rows patched as one list, read and keyed once however many edits there are; the key is
  // put in its place in the sheet at the first edit, as a map patched at once would be
  const patching = new Map<SheetInput, { columns?: Line<ColumnInput>[]; rows?: Line<RowInput>[] }>()
  const listOf = <K extends 'columns' | 'rows'>(sheet: SheetInput, key: K) => {
    const x = patching.get(sheet) ?? patching.set(sheet, {}).get(sheet)!
    return (x[key] ??= lines((sheet[key] = sheet[key]) as never, key === 'columns' ? 'c' : 'r')) as NonNullable<(typeof x)[K]>
  }
  for (const e of edits) {
    const i = sheetIndex(e.sheet)
    const sheet = changed.get(i) ?? { ...book.sheets[i] }
    changed.set(i, sheet)
    if ('merges' in e) sheet.merges = [...e.merges]
    else if ('filtered' in e) sheet.filtered = e.filtered
    else if ('frozen' in e) sheet.frozen = e.frozen ? { ...e.frozen } : undefined
    else if ('autoFilter' in e) sheet.autoFilter = e.autoFilter ? { range: e.autoFilter.range } : undefined
    else if ('conditionalFormats' in e) sheet.conditionalFormats = e.conditionalFormats.length ? [...e.conditionalFormats] : undefined
    else if ('validations' in e) sheet.validations = e.validations.map((v) => ({ ...v }))
    else if ('outline' in e) sheet.outline = e.outline ? { ...e.outline } : undefined
    else if ('protection' in e) sheet.protection = e.protection ? { ...e.protection } : undefined
    // a width or height assigned leaves AutoFit; an explicit bestFit still describes a fitted size
    else if ('columns' in e) patchIn(listOf(sheet, 'columns'), ...parseSpan(e.columns, 'c')!, e.set.width !== undefined && e.set.bestFit === undefined ? { ...e.set, bestFit: null } : e.set)
    else patchIn(listOf(sheet, 'rows'), ...parseSpan(e.rows, 'r')!, e.set.height !== undefined && e.set.bestFit === undefined ? { ...e.set, bestFit: null } : e.set)
  }
  for (const [sheet, { columns, rows }] of patching) {
    if (columns) sheet.columns = toMap(columns, 'c')
    if (rows) sheet.rows = toMap(rows, 'r')
  }
  for (const [i, sheet] of changed) {
    for (const k of ['columns', 'rows', 'merges', 'validations'] as const) if (sheet[k] && !Object.keys(sheet[k]).length) delete sheet[k]
    if (sheet.filtered === false) delete sheet.filtered
    for (const k of ['frozen', 'autoFilter', 'outline', 'conditionalFormats', 'protection'] as const) if (sheet[k] === undefined) delete sheet[k]
    book.sheets[i] = sheet
  }
  return changed
}
