/**
 * Whether a file can hold a formula of the book, and its tree when it can: what the sheet, table, conditional format and
 * data validation writers share before a formula is printed in a file's spelling (`formulas/to-file.ts`). The formulas
 * of a book's cells of one shape (`syntax/shape.ts`), a formula copied down a column, are checked and printed once.
 */
import type { ExternalBook } from '../../contract.js'
import { FormulaError, type Ast } from '../../syntax/ast.js'
import { parseTokens } from '../../syntax/parse.js'
import { sheetKey, walk } from '../../syntax/print.js'
import { bindTables } from '../../syntax/check.js'
import { bindLinks } from '../../syntax/links.js'
import { filled, instanceText, ready, shaped, SHAPES, tableShapeKey, type Hole, type Template } from '../../syntax/shape.js'
import type { Tok } from '../../syntax/tokens.js'
import { FILE_PRINT, toFile } from '../formulas/to-file.js'

/** What a formula of the book may name in its file. */
export interface FileNames {
  /** Upper-case names (`sheetKey`) of the book's sheets. */
  sheets: Set<string>
  /** The book's tables with their columns. */
  tables: { name: string; columns: string[] }[]
  /** The book's links to other workbooks, which `[1]Sheet1!A1` names. */
  externalBooks?: ExternalBook[]
}

/**
 * The formula's tree when the file can hold it: it parses, names only the book's sheets, its structured references name
 * the book's tables and columns (a file names a reference's table always: `[@Qty]` is written `Sales[[#This Row],[Qty]]`),
 * and its references to other workbooks the book's links and their sheets. A defined name's ref (`opts.anySheet`) keeps
 * a reference to a sheet the book does not have, and binds its links without the walk that looks for them. `opts.toks`
 * are the formula's tokens where they were read already.
 */
export function fileTree(formula: string, names: FileNames, enclosing?: string, opts: { union?: boolean; anySheet?: boolean; toks?: Tok[] } = {}): Ast | null {
  let ast: Ast
  try {
    ast = parseTokens(formula, opts.toks ?? null, opts)
  } catch (e) {
    if (e instanceof FormulaError) return null
    throw e
  }
  const bound = bindTables(ast, formula, { tables: names.tables, ...(enclosing ? { table: enclosing } : {}) })
  if (!bound.ok) return null
  ast = bound.ast
  let ok = true
  // a defined name's ref may name a sheet the book does not have; its links are bound without the walk
  let linked = !!opts.anySheet
  const known = (sheet: string | null | undefined) => sheet == null || names.sheets.has(sheetKey(sheet))
  if (!opts.anySheet)
    walk(ast, (n) => {
      // a reference into another workbook, one Excel lost too (`[1]Sheet1!#REF!`, of any sheet of its link)
      if (n.k === 'xref' || n.k === 'xname' || (n.k === 'err' && n.book !== undefined)) linked = true
      else if (n.k === 'ref3d' ? !known(n.first) || !known(n.last) : (n.k === 'ref' || n.k === 'err' || n.k === 'name') && !known(n.sheet)) ok = false
    })
  if (!ok || !linked) return ok ? ast : null
  const links = bindLinks(ast, names.externalBooks, [...names.sheets])
  return links.ok ? links.ast : null
}

/** The shapes of a book's formulas written, by the table they are in and their key: a template, or null for a shape written formula by formula. */
export type FileShapes = Map<string, Template | null>

/** A hole as a file spells it: a number with an upper-case exponent (`toFile`). */
const fileHole = (h: Hole) => (typeof h === 'number' ? String(h).toUpperCase() : filled(h))

/**
 * A formula of a cell in a file's spelling (`toFile` of its `fileTree`), null when the file cannot hold it; through its
 * shape's template where a formula of the shape in the same table (`enclosing`) was written before (`shapes`, a book's).
 */
export function fileText(formula: string, names: FileNames, enclosing: string | undefined, shapes?: FileShapes): string | null {
  const own = (toks?: Tok[]) => {
    const tree = fileTree(formula, names, enclosing, toks ? { toks } : {})
    return tree && toFile(tree)
  }
  if (!shapes) return own()
  const read = shaped(formula)
  if (read === null) return own()
  if (read === 'unread') return null
  const { toks, shape } = read
  // a spill reference is written as ANCHORARRAY, which its shape does not tell
  if (!shape || shape.holes.some((h) => typeof h !== 'number' && h.spill)) return own(toks)
  const key = tableShapeKey(enclosing, shape.key)
  const t = shapes.get(key)
  if (t && (t.at || ready(t, FILE_PRINT, fileHole))) return instanceText(t, shape.holes)
  if (t !== undefined) {
    // a shape the file cannot hold, or a template whose tree or text cannot take holes in
    shapes.set(key, null)
    return own(toks)
  }
  const tree = fileTree(formula, names, enclosing, { toks })
  if (shapes.size < SHAPES) shapes.set(key, tree ? { ast: tree, holes: shape.holes } : null)
  return tree && toFile(tree)
}
