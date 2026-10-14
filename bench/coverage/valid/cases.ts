/**
 * Data validation and outlines for Excel (`run.ts valid <set>`). A case's book is written by `writeXlsx`; Excel opens
 * it, makes what the case makes through its object model (`Range.Validation.Add`, `Rows.Group`, `ShowDetail`), applies
 * the case's structure operation, reads each cell of a window that has a rule (the rule with the cell selected, and
 * whether the cell's value meets it: `Validation.Value`), the window's outline levels, hidden and collapsed lines and the
 * sheet's summary settings, and saves the book. The answer is what Excel read, and the file it saved goes beside it.
 */
import type { BookInput, CellInput, DataValidation, OutlineSettings, StructureOp, ValidationOperator, ValidationType } from '../../../src/contract.js'
import { LOCALES, createSession, parseEntry, toInvariant, validateFormula } from '../../../src/index.js'
import { isHiddenRow } from '../../../src/layout/hidden.js'
import { lines, parseSpan, patchLines } from '../../../src/layout/spans.js'
import { formatAddress, indexToCol } from '../../../src/model.js'
import { parseRange } from '../../../src/structure/area.js'
import { sheetKey } from '../../../src/syntax.js'
import { anchorOf, areasOf } from '../../../src/validation/rules.js'
import { formulaAt } from '../../../src/validation/structure.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'
import { rawFile, type RawEdit } from '../xlsx/sent.js'

/**
 * What Excel makes through its object model: a rule (with the range's first cell active; `excel`, its formulas as typed
 * into the object model in the user's language, when not those of the rule), a group, a collapsed group, summary settings.
 */
export type ValidMake =
  | { rule: Omit<DataValidation, 'range'>; range: string; sheet: string; excel?: [string, string?] }
  | { group: string; sheet: string }
  | { collapse: string; sheet: string }
  | { summary: OutlineSettings; sheet: string }

export interface ValidCase {
  id: string
  /** The key of the case's book in the set's `books`. */
  book: string
  make?: ValidMake[]
  op?: StructureOp
  /** The file changed by hand after `writeXlsx`, as another writer leaves it: the engine holds what the reader reads from it. */
  raw?: RawEdit[]
}

export interface ValidCases {
  books: Record<string, BookInput>
  cases: ValidCase[]
  /** The rectangle of each sheet that is read (`A1:J30` by default). */
  window?: string
}

export interface ValidSet {
  buildValidCases(): ValidCases
}

export const WINDOW = 'A1:J30'

/** A case's file: its book as `writeXlsx` writes it, changed by hand for a raw case. */
export async function validCaseFile(set: ValidCases, c: ValidCase): Promise<Uint8Array> {
  return rawFile((await writeXlsx({ book: set.books[c.book] })).data, c.raw)
}

/** The book the engine holds for a case before what Excel makes: its book, or for a raw case the book the reader reads from its file. */
export const validCaseBook = async (set: ValidCases, c: ValidCase): Promise<BookInput> => (c.raw ? (await readXlsx(await validCaseFile(set, c))).book : set.books[c.book])

/**
 * A cell with a rule as Excel's object model gives it, with the cell selected: address, type, operator, formula1,
 * formula2, ignore blank, dropdown, show input, show error, alert style, input title and message, error title and
 * message, and whether the cell's value meets the rule.
 */
export type ExcelRule = [string, string, string, string, string, boolean, boolean, boolean, boolean, string, string, string, string, string, boolean]

/** A row or a column of the window that is grouped, hidden or collapsed: its number or letters, level (0 for none), hidden, collapsed. */
export type ExcelLine = [number | string, number, boolean, boolean]

export interface ValidSheetAnswer {
  name: string
  rules: ExcelRule[]
  rows: ExcelLine[]
  columns: ExcelLine[]
  /** The sheet's outline: summary row and summary column, as Excel names them. */
  summary: [string, string]
}

export interface ValidAnswer {
  /** Excel's error when it refused the operation or a step. */
  error?: string
  sheets: ValidSheetAnswer[]
}

/** A cell's rule in one shape for Excel, the engine and the reader: formulas as at the cell, canonical; texts left out when empty. */
export interface CellRule {
  type: ValidationType
  operator?: ValidationOperator
  formula1?: string
  formula2?: string
  allowBlank: boolean
  /** Lists only. */
  dropdown?: boolean
  showPrompt: boolean
  showError: boolean
  errorStyle: 'stop' | 'warning' | 'information'
  promptTitle?: string
  prompt?: string
  errorTitle?: string
  error?: string
}

/** A sheet's data validation and outline over the window, in one shape for Excel, the engine and the reader. */
export interface SheetView {
  name: string
  rules: Record<string, CellRule>
  /** Whether each cell with a rule holds a value that meets it. */
  meets: Record<string, boolean>
  /** Grouped, hidden or collapsed lines: level, hidden, collapsed. */
  rows: Record<string, [number, boolean, boolean]>
  columns: Record<string, [number, boolean, boolean]>
  summary: { below: boolean; right: boolean }
}

const TYPES: Record<string, ValidationType> = {
  'validate input only': 'any',
  'validate whole number': 'whole',
  'validate decimal': 'decimal',
  'validate list': 'list',
  'validated date': 'date',
  'validate time': 'time',
  'validate text length': 'textLength',
  'validate custom': 'custom',
}
const OPERATORS: Record<string, ValidationOperator> = {
  'operator between': 'between',
  'operator not between': 'notBetween',
  'operator equal': 'equal',
  'operator not equal': 'notEqual',
  'operator greater': 'greaterThan',
  'operator less': 'lessThan',
  'operator greater equal': 'greaterThanOrEqual',
  'operator less equal': 'lessThanOrEqual',
}
const ALERTS: Record<string, CellRule['errorStyle']> = { 'valid alert stop': 'stop', 'valid alert warning': 'warning', 'valid alert information': 'information' }
const COMPARED = new Set<ValidationType>(['whole', 'decimal', 'date', 'time', 'textLength'])

/** Formula text as the engine spells it (canonical), or as written when it does not parse. */
export function canonicalText(text: string): string {
  const v = validateFormula(text)
  return v.ok ? v.canonical : text
}

/**
 * A rule's formula as Excel's object model gives it, as stored text: with `=`, a formula (in the user's language for a
 * locale other than en-US); without, a list's items (a text constant) or another type's value as written.
 */
export function excelFormula(text: string, type: ValidationType, locale = 'en-US'): string | undefined {
  if (text === '') return undefined
  if (!text.startsWith('=')) {
    if (type === 'list') return `"${text.replace(/"/g, '""')}"`
    // a value Excel shows as a date, a time or a number in the user's locale, as a cell would show it
    const e = parseEntry(text, { locale })
    return e.ok && typeof e.input?.value === 'number' ? String(e.input.value) : canonicalText(text)
  }
  const body = text.slice(1)
  if (locale === 'en-US') return canonicalText(body)
  const t = toInvariant(body, LOCALES.get(locale)!)
  return t.ok ? canonicalText(t.formula) : body
}

/** A rule of Excel's in the shape of `CellRule`. */
export function excelRule(x: ExcelRule, locale = 'en-US'): CellRule {
  const [, typeText, opText, f1, f2, ignore, dropdown, showInput, showError, alert, inputTitle, inputMessage, errorTitle, errorMessage] = x
  const type = TYPES[typeText] ?? (typeText as ValidationType)
  const out: CellRule = { type, allowBlank: ignore, showPrompt: showInput, showError, errorStyle: ALERTS[alert] ?? 'stop' }
  if (COMPARED.has(type)) out.operator = OPERATORS[opText] ?? (opText as ValidationOperator)
  const a = type === 'any' ? undefined : excelFormula(f1, type, locale)
  if (a !== undefined) out.formula1 = a
  const b = COMPARED.has(type) && (out.operator === 'between' || out.operator === 'notBetween') ? excelFormula(f2, type, locale) : undefined
  if (b !== undefined) out.formula2 = b
  if (type === 'list') out.dropdown = dropdown
  for (const [k, v] of [['promptTitle', inputTitle], ['prompt', inputMessage], ['errorTitle', errorTitle], ['error', errorMessage]] as const) if (v !== '') out[k] = v
  return out
}

/** A rule of the model at a cell, its formulas as at the cell (`at`), in the shape of `CellRule`. */
export function modelRule(v: DataValidation, at: (text: string) => string): CellRule {
  const out: CellRule = { type: v.type, allowBlank: v.allowBlank !== false, showPrompt: v.showPrompt !== false, showError: v.showError !== false, errorStyle: v.errorStyle ?? 'stop' }
  if (COMPARED.has(v.type)) out.operator = v.operator ?? 'between'
  if (v.formula1 !== undefined) out.formula1 = canonicalText(at(v.formula1))
  if (v.formula2 !== undefined) out.formula2 = canonicalText(at(v.formula2))
  if (v.type === 'list') out.dropdown = v.dropdown !== false
  for (const k of ['promptTitle', 'prompt', 'errorTitle', 'error'] as const) if (v[k]) out[k] = v[k]
  return out
}

/** What Excel read on a sheet, in the shape of `SheetView`. */
export function excelView(s: ValidSheetAnswer, locale = 'en-US'): SheetView {
  const view: SheetView = { name: s.name, rules: {}, meets: {}, rows: {}, columns: {}, summary: { below: s.summary[0] !== 'summary above', right: s.summary[1] !== 'summary on left' } }
  for (const x of s.rules) {
    const addr = x[0].replaceAll('$', '')
    view.rules[addr] = excelRule(x, locale)
    view.meets[addr] = x[14]
  }
  for (const [k, level, hidden, collapsed] of s.rows) view.rows[String(k)] = [level, hidden, collapsed]
  for (const [k, level, hidden, collapsed] of s.columns) view.columns[String(k)] = [level, hidden, collapsed]
  return view
}

/**
 * A book's sheets over the window, as the engine holds them: each cell's rule with its formulas as at the cell, whether
 * its value meets the rule (`Session.validateEntry` with the cell's own input), the lines' levels, hidden and collapsed.
 */
export function bookViews(book: BookInput, window = WINDOW): SheetView[] {
  const w = parseRange(window)!
  const session = createSession(book, { display: false, clock: { now: '2026-10-10T12:00:00Z', seed: 1 } })
  const names = new Map(book.sheets.map((s) => [sheetKey(s.name), s.name]))
  try {
    return book.sheets.map((s) => {
      const view: SheetView = { name: s.name, rules: {}, meets: {}, rows: {}, columns: {}, summary: { below: s.outline?.summaryBelow !== false, right: s.outline?.summaryRight !== false } }
      for (const v of s.validations ?? [])
        for (const x of areasOf(v.range)!)
          for (let r = Math.max(x.r1, w.r1); r <= Math.min(x.r2, w.r2); r++)
            for (let c = Math.max(x.c1, w.c1); c <= Math.min(x.c2, w.c2); c++) {
              const addr = formatAddress(r, c)
              view.rules[addr] = modelRule(v, (text) => formulaAt(text, s.name, anchorOf(v), { r, c }, names))
              view.meets[addr] = session.validateEntry(s.name, addr, s.cells[addr] ?? null).ok
            }
      for (const [lo, hi, e] of lines(s.rows, 'r'))
        for (let r = Math.max(lo, w.r1); r <= Math.min(hi, w.r2); r++) {
          const line: [number, boolean, boolean] = [e.outlineLevel ?? 0, isHiddenRow(e), e.collapsed === true]
          if (line[0] || line[1] || line[2]) view.rows[String(r + 1)] = line
        }
      for (const [lo, hi, e] of lines(s.columns, 'c'))
        for (let c = Math.max(lo, w.c1); c <= Math.min(hi, w.c2); c++) {
          const line: [number, boolean, boolean] = [e.outlineLevel ?? 0, e.hidden === true || e.width === 0, e.collapsed === true]
          if (line[0] || line[1] || line[2]) view.columns[indexToCol(c)] = line
        }
      return view
    })
  } finally {
    session.close()
  }
}

const show = (v: unknown) => (v === undefined ? 'none' : JSON.stringify(v))

/** The lines where two views of the same sheets differ: a sheet missing on either side, a rule, a check of a value, a line, the summary settings. */
export function viewDifferences(excel: SheetView[], book: SheetView[]): string[] {
  const out: string[] = []
  for (const name of [...new Set([...excel.map((s) => s.name), ...book.map((s) => s.name)])]) {
    const [x, y] = [excel.find((s) => s.name === name), book.find((s) => s.name === name)]
    if (!x || !y) {
      out.push(`${name}: ${x ? 'only in Excel' : 'only in the book'}`)
      continue
    }
    const keys = <T>(a: Record<string, T>, b: Record<string, T>) => [...new Set([...Object.keys(a), ...Object.keys(b)])].sort((p, q) => p.localeCompare(q, 'en', { numeric: true }))
    for (const addr of keys(x.rules, y.rules)) {
      const [a, b] = [show(x.rules[addr]), show(y.rules[addr])]
      if (a !== b) out.push(`${name} rule ${addr}: Excel ${a}, book ${b}`)
      else if (x.meets[addr] !== y.meets[addr]) out.push(`${name} value ${addr}: Excel ${x.meets[addr] ? 'meets' : 'fails'}, book ${y.meets[addr] ? 'meets' : 'fails'}`)
    }
    for (const kind of ['rows', 'columns'] as const)
      for (const k of keys(x[kind], y[kind])) {
        const [a, b] = [show(x[kind][k]), show(y[kind][k])]
        if (a !== b) out.push(`${name} ${kind === 'rows' ? 'row' : 'column'} ${k}: Excel ${a}, book ${b}`)
      }
    if (show(x.summary) !== show(y.summary)) out.push(`${name} summary: Excel ${show(x.summary)}, book ${show(y.summary)}`)
  }
  return out
}

/**
 * The book with what Excel makes, as the engine holds it: a rule over its range; a group one level deeper on each of its
 * lines; a collapsed group's lines (those before its summary line, below or right by the sheet's settings, one level
 * deeper than it at least) hidden, the summary line marked collapsed; summary settings.
 */
export function made(book: BookInput, make: ValidMake[] = []): BookInput {
  const sheets = book.sheets.map((s) => ({ ...s }))
  const sheet = (name: string) => sheets.find((s) => s.name === name)!
  for (const m of make) {
    const s = sheet(m.sheet)
    if ('rule' in m) s.validations = [...(s.validations ?? []), { range: m.range, ...m.rule }]
    else if ('summary' in m) s.outline = { ...s.outline, ...m.summary }
    else {
      const text = 'group' in m ? m.group : m.collapse
      const axis = /^\d/.test(text) ? 'r' : 'c'
      const key = axis === 'r' ? 'rows' : 'columns'
      let map = s[key] as Record<string, { outlineLevel?: number; hidden?: boolean; collapsed?: boolean }> | undefined
      const at = (i: number) => lines(map, axis).find(([lo, hi]) => i >= lo && i <= hi)?.[2]
      if ('group' in m) {
        const [lo, hi] = parseSpan(text, axis)!
        for (let i = lo; i <= hi; i++) map = patchLines(map, axis, i, i, { outlineLevel: (at(i)?.outlineLevel ?? 0) + 1 })
      } else {
        const [i] = parseSpan(text, axis)!
        const level = at(i)?.outlineLevel ?? 0
        const before = axis === 'r' ? s.outline?.summaryBelow !== false : s.outline?.summaryRight !== false
        for (let k = before ? i - 1 : i + 1; k >= 0 && (at(k)?.outlineLevel ?? 0) > level; k += before ? -1 : 1) map = patchLines(map, axis, k, k, { hidden: true })
        map = patchLines(map, axis, i, i, { collapsed: true })
      }
      ;(s as Record<string, unknown>)[key] = map
    }
  }
  return { ...book, sheets }
}

/** A group's or a collapse's range as Excel's object model takes it: rows `3:5` as `A3:A5`, columns `B:D` as `B1:D1`. */
export function lineRange(text: string): string {
  const m = /^(\d+)(?::(\d+))?$/.exec(text)
  if (m) return `A${m[1]}:A${m[2] ?? m[1]}`
  const [a, b = a] = text.split(':')
  return `${a}1:${b}1`
}

/** A column of a set of values: a rule and the values below its title. */
export type ValueColumn = { rule: Omit<DataValidation, 'range'>; values: (CellInput | null)[] }

/** A value cell and a formula cell. */
export const v = (value: CellInput['value']): CellInput => ({ value })
export { formulaCell as f } from '../stored.js'

/**
 * The columns laid out on a sheet from A, into `cells`: row 1 the rule's title, the values below it (a null one left
 * blank); the rules over the values, returned.
 */
export function laidOut(columns: readonly ValueColumn[], cells: Record<string, CellInput>): DataValidation[] {
  return columns.map((col, i) => {
    const letter = String.fromCharCode(65 + i)
    cells[`${letter}1`] = { value: `${col.rule.type} ${i + 1}` }
    col.values.forEach((x, k) => {
      if (x) cells[`${letter}${k + 2}`] = x
    })
    return { range: `${letter}2:${letter}${col.values.length + 1}`, ...col.rule }
  })
}
