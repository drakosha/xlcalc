/**
 * Formula shapes: formulas whose tokens differ only in their references and numbers (a formula copied down a column,
 * `A2*1.1+1`, `A3*1.1+2`) read, check and print alike, so one of them is parsed, checked and printed as the shape's
 * template and each other is the template with its own references and numbers put in. No check reads where a reference
 * points or what a number is, and both print as written (a number as Excel prints it). The numbers of an array constant
 * stay in the key, as one literal with a sign; a LET has no shape, since a cell's word in it may be a name.
 */
import { MAX_FORMULA_LENGTH } from '../contract.js'
import { FormulaError, type Ast, type RangeRef } from './ast.js'
import { mapAst, printFormula, walk, type PrintOptions } from './print.js'
import { tokenize, type Tok } from './tokens.js'

/** What a formula of a shape puts in its template: its references, and its numbers outside array constants. */
export type Hole = RangeRef | number

/** A shape's template: the AST of its first formula and that formula's holes, which the AST holds. */
export interface Template {
  ast: Ast
  holes: Hole[]
  /** Set at the first formula that takes the template: where each hole is, by a reference or a number's node. */
  at?: Map<RangeRef | Ast, number>
  /** The text around the holes, and the holes in the order the text prints them. */
  pieces?: string[]
  order?: number[]
  /** How a hole is printed into the text (`ready`). */
  fill?: (h: Hole) => string
}

const sized = (x: string) => `${x.length}:${x}`
const bookText = (b: number | string) => (typeof b === 'number' ? `#${b}` : sized(b))

/** At most this many shapes are kept by a cache of shapes: a workbook of formulas that share none keeps no more. */
export const SHAPES = 4096

/** A shape's key within a table (`-` outside one), as its formulas read and print alike only in the same table. */
export const tableShapeKey = (table: string | undefined, key: string) => `${table === undefined ? '-' : sized(table)}:${key}`

/**
 * A formula's tokens and shape, for a cache of shapes: `null` for a text too long or blank, which fails before it is
 * read and is taken on its own; `'unread'` for one that does not tokenize; `shape` null for a LET.
 */
export function shaped(text: string): { toks: Tok[]; shape: ReturnType<typeof shapeOf> } | 'unread' | null {
  if (text.length > MAX_FORMULA_LENGTH || text.trim() === '') return null
  let toks: Tok[]
  try {
    toks = tokenize(text)
  } catch (e) {
    if (e instanceof FormulaError) return 'unread'
    throw e
  }
  return { toks, shape: shapeOf(toks) }
}

/** A formula's shape by its tokens: a key every token but a hole is spelled into, and the holes in the text's order; null for a LET. */
export function shapeOf(toks: Tok[]): { key: string; holes: Hole[] } | null {
  let key = ''
  const holes: Hole[] = []
  let braces = 0
  for (const t of toks) {
    switch (t.t) {
      case 'num':
        if (braces) key += `n${t.v};`
        else {
          holes.push(t.v)
          key += 'n'
        }
        break
      case 'str':
        key += `s${sized(t.v)}`
        break
      case 'err':
        key += `e${t.code};${t.sheet === undefined ? '-' : sized(t.sheet)}${t.ref === undefined ? '-' : sized(t.ref)}${t.book === undefined ? '-' : bookText(t.book)}`
        break
      case 'ref':
        holes.push(t.ref)
        key += `r${t.sheet === null ? '-' : sized(t.sheet)}${t.last === undefined ? '-' : sized(t.last)}`
        break
      case 'ident':
        key += `i${sized(t.v)}${t.sheet === undefined ? '-' : sized(t.sheet)}`
        break
      case 'xref':
        holes.push(t.ref)
        key += `x${bookText(t.book)}${sized(t.sheet)}`
        break
      case 'xname':
        key += `y${bookText(t.book)}${t.sheet === undefined ? '-' : sized(t.sheet)}${sized(t.v)}`
        break
      case 'table':
        key += `t${t.table === null ? '-' : sized(t.table)}${t.spec.areas.join(',')};${t.spec.cols ? t.spec.cols.map(sized).join('') : '-'}`
        break
      case 'fn':
        if (t.v === 'LET') return null
        key += `f${sized(t.v)}`
        break
      case 'op':
        key += `o${t.v};`
        break
      case 'eof':
        break
      default:
        if (t.t === '{') braces++
        else if (t.t === '}') braces--
        key += t.t
    }
  }
  return { key, holes }
}

/** A mark in the template's text where a hole goes; a text that holds it has no pieces. */
const MARK = '\u0000'

/** A hole as the canonical text prints it. */
export const filled = (h: Hole) => (typeof h === 'number' ? String(h) : h.text + (h.spill ? '#' : ''))

/**
 * Makes the template ready to be taken (`instance`), printing it with `print` (its `node` too), a hole as `fill` prints
 * it; false when its AST does not hold each hole once or its text cannot be cut at them.
 */
export function ready(t: Template, print: PrintOptions, fill = filled): boolean {
  const at = new Map<RangeRef | Ast, number>()
  const nums: number[] = []
  t.holes.forEach((h, i) => (typeof h === 'number' ? nums.push(i) : at.set(h, i)))
  let refs = 0, k = 0
  walk(t.ast, (n) => {
    if ((n.k === 'ref' || n.k === 'ref3d' || n.k === 'xref') && at.has(n.ref)) refs++
    else if (n.k === 'num') at.set(n, nums[k++] ?? -1)
  })
  if (refs !== t.holes.length - nums.length || k !== nums.length) return false
  const order: number[] = []
  const text = printFormula(t.ast, {
    ...print,
    node: (n, go, usual) => {
      const i = n.k === 'ref' || n.k === 'ref3d' || n.k === 'xref' ? at.get(n.ref) : n.k === 'num' ? at.get(n) : undefined
      if (i === undefined) return print.node?.(n, go, usual)
      order.push(i)
      const u = print.node?.(n, go, usual) ?? usual(n), f = fill(t.holes[i])
      return u.endsWith(f) ? u.slice(0, u.length - f.length) + MARK : u
    },
  })
  const pieces = text.split(MARK)
  if (pieces.length !== order.length + 1 || order.length !== t.holes.length) return false
  Object.assign(t, { at, pieces, order, fill })
  return true
}

/** The text of a formula of the template's shape (`ready`), whose holes are `holes`. */
export function instanceText(t: Template, holes: Hole[]): string {
  const pieces = t.pieces!, order = t.order!, fill = t.fill ?? filled
  // joined, not concatenated: a formula's text is kept, and a long one concatenated hole by hole keeps a node per piece
  const parts = [pieces[0]]
  for (let j = 0; j < order.length; j++) parts.push(fill(holes[order[j]]), pieces[j + 1])
  return parts.join('')
}

/** The AST and text of a formula of the template's shape (`ready`), whose holes are `holes`. */
export function instance(t: Template, holes: Hole[]): { ast: Ast; text: string } {
  const at = t.at!
  const ast = mapAst(t.ast, (n) => {
    if (n.k === 'ref' || n.k === 'ref3d' || n.k === 'xref') {
      const i = at.get(n.ref)
      return i === undefined ? undefined : { ...n, ref: holes[i] as RangeRef }
    }
    if (n.k !== 'num') return undefined
    const i = at.get(n)
    return i === undefined ? undefined : { k: 'num', v: holes[i] as number }
  })
  return { ast, text: instanceText(t, holes) }
}
