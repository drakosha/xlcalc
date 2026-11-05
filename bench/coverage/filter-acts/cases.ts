/**
 * What an AutoFilter does to the rows of its sheet, for Excel (`run.ts filter-acts <set>`, `bench/windows/run.ts
 * filter-acts <set>`): each case is a book of one sheet `S`, its steps taken in turn on the header's button or through
 * the object model (`via`), and the sheet read as it opened and after each step: the range the AutoFilter acts on, the
 * filter mode, the hidden rows, where the total row is and what it shows; at the end what the saved file holds. A step
 * Excel for Mac and for Windows both take on the screen, as a person does, where `via` is `button`.
 */
import JSZip from 'jszip'
import type { BookInput, CellInput, RowInput, SheetInput } from '../../../src/contract.js'
import { createSession, registerLocale } from '../../../src/index.js'
import { ruRU } from '../../../src/locales/ru-RU.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { openKept, writeKept } from '../../../src/xlsx-keep.js'
import { zip } from '../../../src/xlsx/zip/write.js'

export type Via = 'model' | 'button'

export type ActStep =
  /** Criteria on a column of the AutoFilter, keeping one value: `Range.AutoFilter` with a field and a value, or the column's button. */
  | { criteria: string; column: string; via: Via }
  /** The AutoFilter sorted by a column, largest first: `AutoFilter.Sort` with that column of `AutoFilter.Range` as its key, or the column's button. */
  | { sort: string; via: Via }
  /** The AutoFilter removed: `AutoFilterMode = false`, or Data > Filter with its header's first cell selected. */
  | { off: Via }
  /** A range given to the filter: `Range.AutoFilter` with no arguments (`plain`), with a field and no value (`field`), or the range selected and Data > Filter (`button`). */
  | { over: string; via: 'plain' | 'field' | 'button' }
  /** A cell's formula or text written through the object model. */
  | { set: string; at: string }
  /** The table's totals row shown, its last column summed (`ListObject.ShowTotals`, `ListColumn.TotalsCalculation`). */
  | { totals: true }

export interface ActCase {
  id: string
  cells: Record<string, CellInput>
  /** The AutoFilter's range in the book, none without it. */
  filter?: string
  /** The sheet's table over this range, with its filter, in place of an AutoFilter. */
  table?: string
  /** Rows the book hides (from 1), and rows of height 0. */
  hidden?: number[]
  zero?: number[]
  /** Criteria the file holds (a `<filterColumn>`) over rows hidden in `hidden`, the sheet marked filtered, as a file written by another program. */
  fileCriteria?: string
  /** `_xlnm._FilterDatabase` in the file, as Excel leaves it. */
  database?: boolean
  /** The file is the book through `writeKept` with these rows shown and the AutoFilter removed in one `setLayout`. */
  keptOff?: string
  /** Rows whose heights are read. */
  watch?: number[]
  steps: ActStep[]
}

/** The sheet as read: the rows from 1 to `ACT_ROWS`. */
export interface ActState {
  /** The range the AutoFilter acts on, empty without one. */
  range: string
  filterMode: boolean
  hidden: number[]
  /** The heights of the rows the case watches. */
  heights: Record<string, number>
  /** The sheet's standard row height. */
  standard: number
  /** The row labelled `TOTAL`: its number, what its cell in column C shows and that cell's formula. */
  total: { row: number; text: string; formula: string } | null
  /** Column C from row 2, as it shows. */
  column: string[]
  /** What a button that was opened lists. */
  listed?: string[]
}

/** The saved file's filter. */
export interface SavedFilter {
  /** The AutoFilter's range, the table's filter's where the sheet has a table. */
  ref: string | null
  tableRef: string | null
  totalsRows: number
  /** A `<filterColumn>`'s values by its column, or its text where it holds another form. */
  criteria: Record<string, string>
  sortRef: string | null
  filterMode: boolean
  hidden: number[]
  /** `ht` of each row that has one, with `customHeight` set or not. */
  heights: Record<string, string>
  /** `_xlnm._FilterDatabase`, null for none. */
  database: string | null
}

export interface ActAnswer {
  /** The window Excel would repair the file through, or its refusal; no other field then. */
  refused?: string
  opened?: ActState
  after?: ActState[]
  saved?: SavedFilter
  /** What the file sent holds, for a case of `keptOff`. */
  written?: SavedFilter
  /** The windows Excel showed meanwhile, answered. */
  dialog?: string
}

export interface ActSet {
  buildActCases(): ActCase[]
  /** The set's own model, with the range the engine reads for an AutoFilter on a range; the common one without it (`predict.ts`). */
  predictAct?(c: ActCase, acted: (own: string) => string): PredictedAct
}

/** The label of the total row, and the rows read. */
export const TOTAL = 'Итого'
export const ACT_ROWS = 30

/** Column A's values in turn, each five times over rows 2 to 21. */
export const REGIONS = ['Север', 'Юг', 'Запад', 'Восток']
export const regionOf = (row: number) => REGIONS[(row - 2) % 4]
export const amountOf = (row: number) => 1000 + 137 * (row - 2)

/** The book of the questions: a header on A1:C1 and 20 rows, regions repeating in column A. */
export function sales(): Record<string, CellInput> {
  const out: Record<string, CellInput> = { A1: { value: 'Регион' }, B1: { value: 'Менеджер' }, C1: { value: 'Сумма' } }
  for (let r = 2; r <= 21; r++) Object.assign(out, { [`A${r}`]: { value: regionOf(r) }, [`B${r}`]: { value: `Менеджер ${r - 1}` }, [`C${r}`]: { value: amountOf(r) } })
  return out
}

/** The total row at `row`: the label, and a formula in column C if given. */
export const total = (row: number, formula?: string): Record<string, CellInput> => ({ [`A${row}`]: { value: TOTAL }, ...(formula ? { [`C${row}`]: { formula } } : {}) })

const SHEET = 'xl/worksheets/sheet1.xml'

/** The book of a case. */
export function actBook(c: ActCase): BookInput {
  const rows: Record<string, RowInput> = Object.fromEntries([...(c.hidden ?? []).map((r): [string, RowInput] => [String(r), { hidden: true }]), ...(c.zero ?? []).map((r): [string, RowInput] => [String(r), { height: 0 }])])
  const s: SheetInput = {
    name: 'S',
    cells: c.cells,
    ...(c.filter ? { autoFilter: { range: c.filter } } : {}),
    ...(c.table ? { tables: [{ name: 'Sales', range: c.table, columns: [{ name: 'Регион' }, { name: 'Менеджер' }, { name: 'Сумма' }] }] } : {}),
    ...(Object.keys(rows).length ? { rows } : {}),
    ...(c.fileCriteria ? { filtered: true } : {}),
  }
  return { sheets: [s] }
}

/** The file of a case: the book as `writeXlsx` writes it, with the criteria and the name put in by hand, and through `writeKept` for a case of `keptOff`. */
export async function actFile(c: ActCase): Promise<Uint8Array> {
  const { data, issues } = await writeXlsx({ book: actBook(c) }, { values: 'omit' })
  if (issues.length) throw new Error(`filter-acts ${c.id}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
  let file = data
  if (c.fileCriteria || c.database) {
    const parts = await JSZip.loadAsync(data)
    const files = new Map<string, string>()
    for (const [name, f] of Object.entries(parts.files)) if (!f.dir) files.set(name, await f.async('string'))
    if (c.fileCriteria) files.set(SHEET, files.get(SHEET)!.replace(/<autoFilter ref="([^"]+)"\/>/, `<autoFilter ref="$1">${c.fileCriteria}</autoFilter>`))
    if (c.database) files.set('xl/workbook.xml', files.get('xl/workbook.xml')!.replace('</sheets>', `</sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">S!${c.filter!.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}</definedName></definedNames>`))
    const enc = new TextEncoder()
    file = await zip([...files].map(([name, text]) => ({ name, data: enc.encode(text) })))
  }
  if (!c.keptOff) return file
  registerLocale(ruRU)
  const r = await openKept(file, { locale: 'ru-RU' })
  const s = createSession(r.book, { cached: r.cached })
  s.setLayout([
    { sheet: 'S', filtered: false },
    { sheet: 'S', rows: c.keptOff, set: { hidden: null } },
    { sheet: 'S', autoFilter: null },
  ])
  return (await writeKept(file, s.getSnapshot(), { locale: 'ru-RU' })).data
}

/** The arguments of a case's script and job: what it reads and does. */
export const actSpec = (c: ActCase) => ({ table: !!c.table, label: TOTAL, rows: ACT_ROWS, watch: c.watch ?? [], steps: c.steps })

const attr = (tag: string | undefined, name: string) => tag?.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1] ?? null

/** What a file holds of its filter. */
export async function savedFilter(data: Uint8Array): Promise<SavedFilter> {
  const parts = await JSZip.loadAsync(data)
  const text = async (name: string) => (await parts.file(name)?.async('string')) ?? ''
  const sheet = await text(SHEET)
  const table = await text('xl/tables/table1.xml')
  const where = table || sheet
  const filter = where.match(/<autoFilter\b[^>]*?(\/>|>[\s\S]*?<\/autoFilter>)/)?.[0]
  const criteria: Record<string, string> = {}
  for (const m of filter?.matchAll(/<filterColumn colId="(\d+)"[^>]*>([\s\S]*?)<\/filterColumn>/g) ?? []) {
    const values = [...m[2].matchAll(/<filter val="([^"]*)"/g)].map((v) => v[1])
    criteria[m[1]] = values.length ? values.join(', ') : m[2]
  }
  const hidden: number[] = []
  const heights: Record<string, string> = {}
  for (const m of sheet.matchAll(/<row\b[^>]*>/g)) {
    const r = attr(m[0], 'r')!
    if (attr(m[0], 'hidden') === '1') hidden.push(Number(r))
    const ht = attr(m[0], 'ht')
    if (ht !== null) heights[r] = `${ht}${attr(m[0], 'customHeight') === '1' ? ' custom' : ''}`
  }
  const database = (await text('xl/workbook.xml')).match(/<definedName\b[^>]*name="_xlnm\._FilterDatabase"[^>]*>([^<]*)</)?.[1] ?? null
  return {
    ref: attr(filter, 'ref'),
    tableRef: table ? attr(table.match(/<table\b[^>]*>/)?.[0], 'ref') : null,
    totalsRows: Number(attr(table.match(/<table\b[^>]*>/)?.[0], 'totalsRowCount') ?? 0),
    criteria,
    sortRef: attr(where.match(/<sortState\b[^>]*>/)?.[0], 'ref'),
    filterMode: attr(sheet.match(/<sheetPr\b[^>]*>/)?.[0], 'filterMode') === '1',
    hidden,
    heights,
    database,
  }
}


/** What a set's prediction says of a state: heights as 0 or `standard`, the total without its formula. */
export interface PredictedState {
  range: string
  filterMode: boolean
  hidden: number[]
  heights: Record<string, number | 'standard'>
  total: { row: number; text: string } | null
}

/** What a set's prediction says of a saved file. */
export interface PredictedSaved {
  ref: string | null
  tableRef: string | null
  criteria: Record<string, string>
  sortRef?: string | null
  filterMode: boolean
  hidden: number[]
}

export interface PredictedAct {
  /** Whether Excel opens the file without its repair window. */
  opens: boolean
  opened?: PredictedState
  after?: PredictedState[]
  saved?: PredictedSaved
}

/** Whether an answer holds all a prediction says, which may be part of what the answer has: the differences, by path. */
export function missedBy(predicted: unknown, got: unknown, path = ''): string[] {
  if (predicted && got && typeof predicted === 'object' && typeof got === 'object' && !(Array.isArray(predicted) && typeof predicted[0] !== 'object'))
    return Object.keys(predicted).flatMap((k) => missedBy((predicted as Record<string, unknown>)[k], (got as Record<string, unknown>)[k], `${path}.${k}`))
  return JSON.stringify(predicted) === JSON.stringify(got) ? [] : [`${path}: predicted ${JSON.stringify(predicted)}, got ${JSON.stringify(got)}`]
}

const predictedState = (s: ActState): PredictedState => ({
  range: s.range,
  filterMode: s.filterMode,
  hidden: s.hidden,
  heights: Object.fromEntries(Object.entries(s.heights).map(([r, h]) => [r, h === s.standard ? 'standard' : h])),
  total: s.total && { row: s.total.row, text: s.total.text },
})

/** An answer of Excel in the shape of a prediction. */
export function asPredicted(a: ActAnswer): PredictedAct {
  if (a.refused) return { opens: false }
  const { ref, tableRef, criteria, sortRef, filterMode, hidden } = a.saved!
  return { opens: true, opened: predictedState(a.opened!), after: a.after!.map(predictedState), saved: { ref, tableRef, criteria, sortRef, filterMode, hidden } }
}
