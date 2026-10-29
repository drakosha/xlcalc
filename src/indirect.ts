/**
 * The reference INDIRECT's text names, as Excel reads it: in A1 style a cell or range (`B2`, `$A:$A`, `'Q1 data'!A1:B9`),
 * a spill reference (`B2#`), a defined name that refers to a reference, or a structured reference (`Sales[Qty]`); in
 * R1C1 style `R2C3`, `R[-1]C`, `R2`, `C3` and ranges of them, relative to the formula's cell. Anything else, an unknown
 * sheet or name, and another workbook's reference (Excel does not read a closed workbook) are #REF!.
 */
import type { EngineState } from './engine.js'
import { codegenContext, sheetIndexOf } from './engine.js'
import { nameRect } from './exec/walk.js'
import { MAX_COLS, MAX_ROWS } from './model.js'
import { nameFor } from './names.js'
import { validateIn } from './prepare.js'
import { validate, type Ast } from './syntax.js'
import { tableAt } from './tables/model.js'
import { E, Rng, type CErr } from './values.js'

/** How many texts a state keeps read: a book's INDIRECTs repeat a few texts many times. */
const KEPT = 4096

/** One line of an R1C1 reference: absolute from 1, or relative to the formula's cell (`[n]`, `R` alone 0). */
type Line = { abs: boolean; n: number }

/** `R…C…`, `R…` or `C…`: rows or columns left out are the whole column or row. */
const PART = /^(?:R(\[[+-]?\d+\]|\d+)?)?(?:C(\[[+-]?\d+\]|\d+)?)?$/i

function line(t: string | undefined): Line {
  if (t === undefined || t === '') return { abs: false, n: 0 }
  return t[0] === '[' ? { abs: false, n: Number(t.slice(1, -1)) } : { abs: true, n: Number(t) }
}

/** A line's index from the formula's `at`, or -1 outside the sheet. */
function index(l: Line, at: number, max: number): number {
  const i = l.abs ? l.n - 1 : at + l.n
  return i >= 0 && i < max && Number.isSafeInteger(l.n) ? i : -1
}

/** An R1C1 reference's parts at row `r`, column `c`: its rows and columns, or undefined where it is not one. */
function r1c1(text: string, r: number, c: number): { r1: number; r2: number; c1: number; c2: number } | undefined {
  const parts = text.split(':')
  if (parts.length > 2) return undefined
  const read = parts.map((p) => {
    const m = p.length ? PART.exec(p) : null
    if (!m) return undefined
    const hasR = /^R/i.test(p), hasC = /C/i.test(p)
    return { row: hasR ? line(m[1]) : undefined, col: hasC ? line(m[2]) : undefined }
  })
  if (read.some((x) => !x)) return undefined
  const [a, b = a] = read as { row?: Line; col?: Line }[]
  // the ends of a range are of one kind: cells, rows or columns
  if (!!a.row !== !!b.row || !!a.col !== !!b.col) return undefined
  const rows = a.row ? [index(a.row, r, MAX_ROWS), index(b.row!, r, MAX_ROWS)] : [0, MAX_ROWS - 1]
  const cols = a.col ? [index(a.col, c, MAX_COLS), index(b.col!, c, MAX_COLS)] : [0, MAX_COLS - 1]
  if ([...rows, ...cols].some((i) => i < 0)) return undefined
  return { r1: Math.min(...rows), r2: Math.max(...rows), c1: Math.min(...cols), c2: Math.max(...cols) }
}

/** A sheet name before `!` (quoted with doubled apostrophes, or plain) and the rest, or the text alone. */
function sheetPart(text: string): { sheet?: string; rest: string } | undefined {
  if (text.startsWith("'")) {
    const m = /^'((?:[^']|'')+)'!(.*)$/s.exec(text)
    return m ? { sheet: m[1].replaceAll("''", "'"), rest: m[2] } : undefined
  }
  const i = text.lastIndexOf('!')
  return i < 0 ? { rest: text } : /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(text.slice(0, i)) ? { sheet: text.slice(0, i), rest: text.slice(i + 1) } : undefined
}

/** INDIRECT over a state: `refer(text, a1, s, r, c)` from the formula at sheet `s`, row `r`, column `c`. */
export function referrer(state: EngineState) {
  const read = new Map<string, Ast | null>()
  // read as a formula of the cell's sheet and table, the table's columns seen without its name
  const parsed = (text: string, s: number, r: number, c: number): Ast | null => {
    const enclosing = tableAt(state.tables, s, r, c)?.name
    const key = `${s}\u0000${enclosing ?? ''}\u0000${text}`
    let ast = read.get(key)
    if (ast === undefined) {
      const v = text.startsWith('=') ? null : validate(text, validateIn(state, s, enclosing))
      ast = v?.ok ? v.ast : null
      if (read.size >= KEPT) read.clear()
      read.set(key, ast)
    }
    return ast
  }
  const sheetOf = (name: string | null, s: number) => (name === null ? s : sheetIndexOf(state, name))
  const table = (n: Extract<Ast, { k: 'table' }>, s: number, r: number): Rng | CErr => {
    const x = codegenContext(state, s).table!(n, r)
    return typeof x === 'string' ? E[x] : new Rng(x.s, x.r1, x.c1, x.r2, x.c2)
  }
  // the reference a name refers to, its relative bounds moved to the formula's cell; a name of a formula is #REF!
  const named = (n: Extract<Ast, { k: 'name' }>, s: number, r: number, c: number): Rng | CErr => {
    const scope = n.sheet === undefined ? s : sheetIndexOf(state, n.sheet)
    const d = scope === undefined ? undefined : nameFor(state.names, scope, n.name)
    const ast = d?.ast.k === 'paren' ? d.ast.a : d?.ast
    if (ast?.k === 'table') return table(ast, s, r)
    if (ast?.k !== 'ref' || ast.ref.spill) return E.REF
    const sheet = sheetOf(ast.sheet, s)
    if (sheet === undefined) return E.REF
    const x = nameRect(ast.ref, r, c)
    return new Rng(sheet, x.r1, x.c1, x.r2, x.c2)
  }
  const a1 = (text: string, s: number, r: number, c: number): Rng | CErr => {
    const ast = parsed(text, s, r, c)
    switch (ast?.k) {
      case 'ref': {
        const sheet = sheetOf(ast.sheet, s)
        if (sheet === undefined) return E.REF
        const x = ast.ref
        if (!x.spill) return new Rng(sheet, x.r1, x.c1, x.r2, x.c2)
        // how far the array spills is known once its formula is computed: the graph learns its cell, for its place alone too
        state.ctx.reads?.push(new Rng(sheet, x.r1, x.c1, x.r1, x.c1))
        return state.ctx.spillRef(sheet, x.r1, x.c1)
      }
      case 'table':
        return table(ast, s, r)
      case 'name':
        return named(ast, s, r, c)
      default:
        return E.REF
    }
  }
  return (text: string, isA1: boolean, s: number, r: number, c: number): Rng | CErr => {
    if (isA1) return a1(text, s, r, c)
    const p = sheetPart(text)
    const sheet = p && (p.sheet === undefined ? s : sheetIndexOf(state, p.sheet))
    const x = p && sheet !== undefined ? r1c1(p.rest, r, c) : undefined
    return x ? new Rng(sheet!, x.r1, x.c1, x.r2, x.c2) : E.REF
  }
}
