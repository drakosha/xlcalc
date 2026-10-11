/**
 * Input as the engine takes it. A formula's leading `=` is dropped: Excel shows a formula with one in its formula bar
 * and in its boxes of names, conditional formats and data validation, and a host passes what it shows: one `=` goes
 * (`=SUM(A1:A2)` is `SUM(A1:A2)`, the same input and the same `inputHash`), a second stays and is an error of the
 * formula (`==1`). What needs none is given back as it is, what needs one as a copy: the caller's input is not changed,
 * and input without one costs no copy.
 */
import type { BookInput, CellEdit, CellFields, CellInput, CellReplace, ConditionalFormat, ConditionalValue, DataValidation, LayoutEdit, NameInput, SheetInput, StructureOp, TableInput } from './contract.js'

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
  return Object.keys(out).length ? { ...s, ...out } : s
}

/** A checked book as taken; `cells`, by sheet, the addresses of the cells that need it (a formula with `=`). */
export function bookTaken(book: BookInput, cells: (string[] | undefined)[]): BookInput {
  const sheets = each(book.sheets, (s, i) => sheetTaken(s, cells[i]))
  const names = book.names && namesTaken(book.names)
  return sheets === book.sheets && names === book.names ? book : { ...book, sheets, ...(names ? { names } : {}) }
}

/** Edits of cells (`setCells`'s by fields, `replaceCells`' whole inputs) as taken. */
export const editsTaken = <E extends CellEdit | CellReplace>(edits: E[]): E[] =>
  each(edits, (e) => {
    const input = e.input && cellTaken(e.input)
    return input === e.input ? e : { ...e, input }
  })

/** Layout edits as taken: the conditional formats and the rules of data validation. */
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
    return e
  })

/** A structure operation without a leading `=` in the formulas of a table it adds; anything else as it is (checked later). */
export function opTaken(op: StructureOp): StructureOp {
  if (op?.op !== 'addTable' || !Array.isArray(op.table?.columns)) return op
  const table = tableTaken(op.table)
  return table === op.table ? op : { ...op, table }
}
