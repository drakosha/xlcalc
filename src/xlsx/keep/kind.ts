/**
 * The oldest form of a formula written anew over a file: a formula of the old kind (`<f>`), which Excel before 365
 * reads, wherever Excel reads it as the workbook computes it, that is where reading it back as the old kind (`@` where
 * Excel 365 shows it, `legacy/convert.ts`) gives the workbook's formula; else none, and it stays a dynamic array.
 */
import type { Ast } from '../../syntax/ast.js'
import { FormulaError } from '../../syntax/ast.js'
import { parse } from '../../syntax/parse.js'
import { canonical, mapAst } from '../../syntax/print.js'
import { escapeText } from '../xml/text.js'

/**
 * The old kind's `<f>` element of a formula of a sheet, or undefined where only a dynamic array means it; `written`, the
 * formula's text as the dynamic array's element has it, which stays where the formula has no `@` to take out.
 */
export function oldKind(formula: string, sheet: string, written: string, legacy: (ast: Ast, sheet: string) => Ast | undefined, text: (formula: string) => string | null): string | undefined {
  let ast: Ast
  try {
    ast = parse(formula)
  } catch (e) {
    if (e instanceof FormulaError) return undefined
    throw e
  }
  // the old kind does the implicit intersection `@` asks for where it puts one itself
  const bare = mapAst(ast, (n) => (n.k === 'at' ? n.a : undefined))
  const back = legacy(bare, sheet)
  if (bare === ast) return back === undefined ? `<f>${written}</f>` : undefined
  if (canonical(back ?? bare) !== canonical(ast)) return undefined
  const file = text(canonical(bare))
  return file === null ? undefined : `<f>${escapeText(file)}</f>`
}
