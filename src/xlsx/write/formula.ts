/**
 * Whether a file can hold a formula of the book, and its tree when it can: what the sheet, table, conditional format and
 * data validation writers share before a formula is printed in a file's spelling (`formulas/to-file.ts`).
 */
import type { ExternalBook } from '../../contract.js'
import { FormulaError, type Ast } from '../../syntax/ast.js'
import { parse } from '../../syntax/parse.js'
import { sheetKey, walk } from '../../syntax/print.js'
import { bindTables } from '../../syntax/check.js'
import { bindLinks } from '../../syntax/links.js'

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
 * and its references to other workbooks the book's links and their sheets.
 */
export function fileTree(formula: string, names: FileNames, enclosing?: string): Ast | null {
  let ast: Ast
  try {
    ast = parse(formula)
  } catch (e) {
    if (e instanceof FormulaError) return null
    throw e
  }
  const bound = bindTables(ast, formula, { tables: names.tables, ...(enclosing ? { table: enclosing } : {}) })
  if (!bound.ok) return null
  ast = bound.ast
  let ok = true
  let linked = false
  const known = (sheet: string | null | undefined) => sheet == null || names.sheets.has(sheetKey(sheet))
  walk(ast, (n) => {
    // a reference into another workbook, one Excel lost too (`[1]Sheet1!#REF!`, of any sheet of its link)
    if (n.k === 'xref' || n.k === 'xname' || (n.k === 'err' && n.book !== undefined)) linked = true
    else if (n.k === 'ref3d' ? !known(n.first) || !known(n.last) : (n.k === 'ref' || n.k === 'err' || n.k === 'name') && !known(n.sheet)) ok = false
  })
  if (!ok || !linked) return ok ? ast : null
  const links = bindLinks(ast, names.externalBooks, [...names.sheets])
  return links.ok ? links.ast : null
}
