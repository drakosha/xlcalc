/**
 * A formula of the book in a file's spelling, printed from its syntax tree: functions newer than Excel 2007 with
 * their prefix (`_xlfn.XLOOKUP`, `_xlfn._xlws.FILTER`), LET as `_xlfn.LET` with its names as `_xlpm.x`, a spill
 * reference `E2#` as `_xlfn.ANCHORARRAY(E2)`, `@x` as `_xlfn.SINGLE(x)`, sheets always quoted, numbers with an upper-case exponent,
 * another workbook's sheet by its link's number in quotes only where its name needs them, as Excel writes it
 * (`[1]Sheet1!A1`, `'[1]My Sheet'!A1`, `[1]Sheet1!Local`, `[1]Sheet1!#REF!`), a reference to a deleted sheet as `#REF!`
 * (the engine keeps `#REF!A1`, as Excel shows it). The rest as the engine's canonical text prints it.
 */
import type { Ast } from '../../syntax/ast.js'
import { ERROR_TEXT } from '../../syntax/errors.js'
import { plain } from '../../syntax/links.js'
import { linkedSheet, printFormula, quoteSheet, type PrintOptions } from '../../syntax/print.js'
import { specText } from '../../syntax/table.js'
import { XLFN, XLWS } from './prefixes.js'


/** Another workbook's sheet before `!`: quoted unless its name is a plain one that reads as no cell. */
const fileLinkSheet = (book: number | string, sheet: string) => (plain(sheet) ? linkedSheet(book, sheet) : quoteSheet(linkedSheet(book, sheet)))

const fn = (name: string) => (XLWS.has(name) ? `_xlfn._xlws.${name}` : XLFN.has(name) ? `_xlfn.${name}` : name)

/** The nodes a file writes otherwise than the canonical text; undefined for the rest. */
function fileNode(n: Ast, go: (n: Ast) => string, usual: (n: Ast) => string): string | undefined {
  switch (n.k) {
    case 'num':
      return String(n.v).toUpperCase()
    case 'err':
      // a reference to a deleted sheet (`#REF!A1`, `#REF!E2#`) is `#REF!` in a file, as Excel writes it: Excel opens no
      // file with the other
      if (n.ref !== undefined && n.sheet === undefined && n.book === undefined) return n.ref.endsWith('#') ? `_xlfn.ANCHORARRAY(${ERROR_TEXT[n.code]})` : ERROR_TEXT[n.code]
      // a lost spill reference (`'S'!#REF!#`) is the anchor of a lost cell
      if (n.ref?.endsWith('#')) return `_xlfn.ANCHORARRAY(${usual({ ...n, ref: n.ref.slice(0, -1) })})`
      return typeof n.book === 'number' && n.sheet !== undefined ? `${fileLinkSheet(n.book, n.sheet)}!${ERROR_TEXT[n.code]}${n.ref ?? ''}` : undefined
    case 'ref': {
      if (!n.ref.spill) return undefined
      const { spill: _spill, ...ref } = n.ref
      return `_xlfn.ANCHORARRAY(${usual({ ...n, ref })})`
    }
    case 'table':
      return `${n.table ?? ''}[${specText(n.spec, { file: true })}]`
    case 'at':
      return `_xlfn.SINGLE(${go(n.a)})`
    case 'fn':
      return usual({ ...n, name: fn(n.name) })
    case 'let':
      return `_xlfn.${usual({ ...n, binds: n.binds.map((b) => ({ ...b, name: `_xlpm.${b.name}` })) })}`
    case 'local':
      return `_xlpm.${n.name}`
    case 'xref':
      return typeof n.book === 'number' ? `${fileLinkSheet(n.book, n.sheet)}!${n.ref.text}` : undefined
    case 'xname':
      return typeof n.book === 'number' && n.sheet !== undefined ? `${fileLinkSheet(n.book, n.sheet)}!${n.name}` : undefined
  }
  return undefined
}

/** The printing of a file's spelling. */
export const FILE_PRINT: PrintOptions = { node: fileNode }

export function toFile(ast: Ast): string {
  return printFormula(ast, FILE_PRINT)
}
