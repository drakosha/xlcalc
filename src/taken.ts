/**
 * Input as the engine takes it. A formula's leading `=` is dropped: Excel shows a formula with one in its formula bar
 * and in its boxes of names, conditional formats and data validation, and a host passes what it shows: one `=` goes
 * (`=SUM(A1:A2)` is `SUM(A1:A2)`, the same input and the same `inputHash`), a second stays and is an error of the
 * formula (`==1`). Protection, a style's and a sheet's, holds only what differs from Excel's defaults
 * (`input/protection.ts`): `{ protection: { locked: true } }` is the style `{}`. What needs none is given back as it is, what needs one as a copy: the caller's input is not changed,
 * and input without one costs no copy. A pivot cache's time of day on Excel's 1899-12-30, as 0.71 read it, is at the model's
 * day 0 (`pivot-time.ts`). With `xlcalc/pivot` loaded, a PivotTable of two value fields or more has Σ Values on its
 * axes (`pivot/hook.ts`).
 */
import type { BookInput, CellEdit, CellFields, CellInput, CellReplace, CellStyle, ConditionalFormat, ConditionalValue, DataValidation, LayoutEdit, NameInput, PivotCache, PivotCacheField, SheetInput, StructureOp, TableInput } from './contract.js'
import { modelValue } from './pivot-time.js'
import { protectionTaken, styleTaken } from './input/protection.js'
import { pivotTaken } from './pivot/hook.js'

/** A formula with a leading `=`. */
export const leadingEquals = (f: unknown): f is string => typeof f === 'string' && f.charCodeAt(0) === 61

const drop = (f: string) => (leadingEquals(f) ? f.slice(1) : f)

/** The object with its formula fields `keys` without their leading `=`; itself when none has one. */
function fields<T extends object>(o: T, keys: readonly (keyof T)[]): T {
  if (typeof o !== 'object' || o === null || !keys.some((k) => leadingEquals(o[k]))) return o
  const out = { ...o }
  for (const k of keys) if (leadingEquals(o[k])) (out as Record<keyof T, unknown>)[k] = (o[k] as string).slice(1)
  return out
}

/** The items each given by `f`; the list itself when none changed. */
function each<T>(list: T[], f: (x: T, i: number) => T): T[] {
  let out: T[] | null = null
  list.forEach((x, i) => {
    const y = f(x, i)
    if (y !== x) (out ??= list.slice())[i] = y
  })
  return out ?? list
}

export const cellTaken = <C extends Pick<CellInput, 'formula'> | CellFields>(c: C): C => fields(c, ['formula'] as (keyof C)[])

export const namesTaken = (names: NameInput[]): NameInput[] => each(names, (n) => fields(n, ['ref']))

function tableTaken(t: TableInput): TableInput {
  const columns = each(t.columns, (c) => fields(c, ['formula']))
  return columns === t.columns ? t : { ...t, columns }
}

const validationsTaken = (rules: DataValidation[]): DataValidation[] => each(rules, (r) => fields(r, ['formula1', 'formula2']))

const threshold = (v: ConditionalValue): ConditionalValue => fields(v, ['value'])

/** A conditional format's formulas and thresholds without a leading `=`. */
function formatTaken(f: ConditionalFormat): ConditionalFormat {
  let out = 'formula' in f ? fields(f, ['formula']) : f
  if (out.type === 'cellIs' && out.formulas.some(leadingEquals)) out = { ...out, formulas: out.formulas.map(drop) }
  if (out.type === 'dataBar' && (leadingEquals(out.min.value) || leadingEquals(out.max.value))) out = { ...out, min: threshold(out.min), max: threshold(out.max) }
  if (out.type === 'colorScale' || out.type === 'iconSet') {
    const points = each<{ value: ConditionalValue }>(out.points, (p) => (leadingEquals(p.value.value) ? { ...p, value: threshold(p.value) } : p))
    if (points !== out.points) out = { ...out, points } as ConditionalFormat
  }
  return out
}

const formatsTaken = (formats: ConditionalFormat[]): ConditionalFormat[] => each(formats, formatTaken)

/** A sheet as taken: `cells`, the addresses of the cells that need it, and its other parts. */
function sheetTaken(s: SheetInput, cells: string[] | undefined): SheetInput {
  const out: Partial<SheetInput> = {}
  if (cells?.length) {
    out.cells = { ...s.cells }
    for (const a of cells) out.cells[a] = cellTaken(s.cells[a])
  }
  const tables = s.tables && each(s.tables, tableTaken)
  if (tables !== s.tables) out.tables = tables
  const formats = s.conditionalFormats && formatsTaken(s.conditionalFormats)
  if (formats !== s.conditionalFormats) out.conditionalFormats = formats
  const rules = s.validations && validationsTaken(s.validations)
  if (rules !== s.validations) out.validations = rules
  const protection = s.protection && protectionTaken(s.protection)
  if (protection !== s.protection) out.protection = protection
  // Σ Values of a PivotTable put on its axes by `xlcalc/pivot`
  const pivot = pivotTaken()
  const pivots = pivot && s.pivotTables && each(s.pivotTables, pivot)
  if (pivots && pivots !== s.pivotTables) out.pivotTables = pivots
  return Object.keys(out).length ? { ...s, ...out } : s
}

/** A number or a date of a cache field's facts or a group's range, in the model's spelling. */
const bound = <T extends { min?: unknown; max?: unknown } | { start?: unknown; end?: unknown }>(o: T, keys: readonly (keyof T)[]): T => {
  if (!keys.some((k) => modelValue(o[k]) !== o[k])) return o
  const out = { ...o }
  for (const k of keys) out[k] = modelValue(o[k])
  return out
}

/** A cache field's items, facts and group in the model's spelling of a time of day (`modelValue`); itself when none needs it. */
function fieldTaken(f: PivotCacheField): PivotCacheField {
  const items = f.items && each(f.items, modelValue)
  const facts = f.facts && bound(f.facts, ['min', 'max'])
  const range = f.group?.range && bound(f.group.range, ['start', 'end'])
  const groupItems = f.group?.items && each(f.group.items, modelValue)
  if (items === f.items && facts === f.facts && range === f.group?.range && groupItems === f.group?.items) return f
  const group = f.group && { ...f.group, ...(range ? { range } : {}), ...(groupItems ? { items: groupItems } : {}) }
  return { ...f, ...(items ? { items } : {}), ...(facts ? { facts } : {}), ...(group ? { group } : {}) }
}

/**
 * Whether a field's records may hold a time on 1899-12-30: a field without items (whose records hold values, not
 * indexes) whose least date is one, as 0.71's reader gave every field its facts from the file's.
 */
const timesInRecords = (f: PivotCacheField): boolean => !f.items && f.facts?.min !== undefined && modelValue(f.facts.min) !== f.facts.min

/** Whether a record holds a time on 1899-12-30 in one of `columns`. */
function oldIn(r: PivotCache['records'] extends (infer R)[] | undefined ? R : never, columns: number[]): boolean {
  for (let i = 0; i < columns.length; i++) {
    const v = r[columns[i]]
    if (v !== null && typeof v === 'object' && modelValue(v) !== v) return true
  }
  return false
}

/** A cache with its times of day on Excel's 1899-12-30, as 0.71 read them, at the model's day 0; itself when it has none. */
function cacheTaken(c: PivotCache): PivotCache {
  const fields = each(c.fields, fieldTaken)
  // only the columns that may hold one are read, as there can be many records
  const columns = c.records ? c.fields.flatMap((f, k) => (timesInRecords(f) ? [k] : [])) : []
  const records = c.records && columns.length ? each(c.records, (r) => (oldIn(r, columns) ? r.map(modelValue) : r)) : c.records
  return fields === c.fields && records === c.records ? c : { ...c, fields, ...(records ? { records } : {}) }
}

/** A checked book as taken; `cells`, by sheet, the addresses of the cells that need it (a formula with `=`). */
export function bookTaken(book: BookInput, cells: (string[] | undefined)[]): BookInput {
  const sheets = each(book.sheets, (s, i) => sheetTaken(s, cells[i]))
  const names = book.names && namesTaken(book.names)
  const styles = book.styles && stylesTaken(book.styles)
  const caches = book.pivotCaches && each(book.pivotCaches, cacheTaken)
  if (sheets === book.sheets && names === book.names && styles === book.styles && caches === book.pivotCaches) return book
  return { ...book, sheets, ...(names ? { names } : {}), ...(styles ? { styles } : {}), ...(caches ? { pivotCaches: caches } : {}) }
}

/** A style table as taken: each style's protection without the defaults (`styleTaken`). */
export const stylesTaken = (styles: CellStyle[]): CellStyle[] => each(styles, styleTaken)

/** Edits of cells (`setCells`'s by fields, `replaceCells`' whole inputs) as taken. */
export const editsTaken = <E extends CellEdit | CellReplace>(edits: E[]): E[] =>
  each(edits, (e) => {
    const input = e.input && cellTaken(e.input)
    return input === e.input ? e : { ...e, input }
  })

/** Layout edits as taken: the conditional formats, the rules of data validation and a sheet's protection. */
export const layoutTaken = (edits: LayoutEdit[]): LayoutEdit[] =>
  each(edits, (e) => {
    if ('conditionalFormats' in e) {
      const conditionalFormats = formatsTaken(e.conditionalFormats)
      return conditionalFormats === e.conditionalFormats ? e : { ...e, conditionalFormats }
    }
    if ('validations' in e) {
      const validations = validationsTaken(e.validations)
      return validations === e.validations ? e : { ...e, validations }
    }
    if ('protection' in e && e.protection) {
      const protection = protectionTaken(e.protection)
      return protection === e.protection ? e : { ...e, protection }
    }
    return e
  })

/** A structure operation without a leading `=` in the formulas of a table it adds; anything else as it is (checked later). */
export function opTaken(op: StructureOp): StructureOp {
  if (op?.op !== 'addTable' || !Array.isArray(op.table?.columns)) return op
  const table = tableTaken(op.table)
  return table === op.table ? op : { ...op, table }
}
