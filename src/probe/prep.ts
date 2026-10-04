/**
 * Общая подготовка книги для обоих вариантов: проверка формул нашим разбором, приведение листов к
 * написанию книги, имена. Неверная формула не роняет расчёт: ячейка получает ошибку, причина в `errors`.
 */
import { parse, FormulaError, sheetKey, validateFormula, type Ast, type CheckCode } from './parser'
import { WHITELIST } from './functions'
import { parseAddress, type BookInput, type CellInput, type ErrorCode } from './model'

export interface PreparedFormula {
  sheet: string
  addr: string
  row: number
  col: number
  ast: Ast
}

export interface InvalidFormula {
  sheet: string
  addr: string
  code: CheckCode
  message: string
  at: number
}

export interface PreparedName {
  name: string
  scope: string | null
  ast: Ast
}

export interface PreparedBook {
  sheets: string[]
  /** Имя листа в любом регистре → написание книги. */
  sheetOf: (s: string) => string
  formulas: PreparedFormula[]
  invalid: InvalidFormula[]
  names: PreparedName[]
}

/** Чем неверная формула видна в ячейке. */
export function invalidError(code: CheckCode): ErrorCode {
  if (code === 'unknown_function' || code === 'unknown_name') return 'NAME'
  if (code === 'unknown_sheet' || code === 'external_ref') return 'REF'
  return 'ERROR'
}

const NAME_RE = /^[A-Za-z_\\][A-Za-z0-9_.]{0,254}$/
const LOOKS_LIKE_CELL = /^[A-Za-z]{1,3}[0-9]+$|^[Rr][0-9]*[Cc][0-9]*$/

/** Допустимое имя диапазона: латиница, не адрес ячейки, не R1C1, не TRUE/FALSE. */
export function validateName(name: string): { ok: true } | { ok: false; message: string } {
  if (!NAME_RE.test(name)) return { ok: false, message: 'name must be latin letters, digits, _ and .' }
  if (LOOKS_LIKE_CELL.test(name)) return { ok: false, message: 'name looks like a cell address' }
  if (/^(TRUE|FALSE)$/i.test(name)) return { ok: false, message: 'reserved name' }
  return { ok: true }
}

export function prepareBook(book: BookInput, functions: ReadonlySet<string> = WHITELIST): PreparedBook {
  const sheets = book.sheets.map((s) => s.name)
  const byKey = new Map(sheets.map((s) => [sheetKey(s), s]))
  const sheetOf = (s: string) => byKey.get(sheetKey(s)) ?? s
  const names: PreparedName[] = []
  for (const n of book.names ?? []) {
    const ok = validateName(n.name)
    if (!ok.ok) throw new Error(`bad name ${n.name}: ${ok.message}`)
    let ast: Ast
    try {
      ast = parse(n.ref)
    } catch (e) {
      if (e instanceof FormulaError) throw new Error(`bad name ref ${n.name}: ${e.message}`)
      throw e
    }
    if (ast.k !== 'ref' || ast.sheet === null) throw new Error(`name ${n.name} must reference a range with a sheet`)
    names.push({ name: n.name, scope: n.scope ? sheetOf(n.scope) : null, ast })
  }
  const nameList = names.map((n) => n.name)
  const formulas: PreparedFormula[] = []
  const invalid: InvalidFormula[] = []
  for (const s of book.sheets) {
    for (const [addr, cell] of Object.entries(s.cells)) {
      if (cell.formula === undefined) continue
      const r = validateFormula(cell.formula, { sheets, names: nameList, functions })
      if (!r.ok) {
        invalid.push({ sheet: s.name, addr, code: r.code, message: r.message, at: r.at })
        continue
      }
      const { row, col } = parseAddress(addr)
      formulas.push({ sheet: s.name, addr, row, col, ast: r.ast })
    }
  }
  return { sheets, sheetOf, formulas, invalid, names }
}

export const isFormulaCell = (c: CellInput) => c.formula !== undefined
