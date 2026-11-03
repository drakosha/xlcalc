/**
 * The oldest form of a formula written: a formula of the old kind (`<f>`), which Excel before 365 reads, wherever Excel
 * reads it as the workbook computes it, that is where reading it back as the old kind (`@` where Excel 365 shows it,
 * `legacy/convert.ts`) gives the workbook's formula; else none, and it is written as a dynamic array. Excel 365 saves a
 * typed formula so too, though it makes a dynamic array of a few that the old kind computes alike (`SUMPRODUCT((A1:A5>2)*B1:B5)`).
 */
import type { NameInput } from '../../contract.js'
import type { Ast } from '../../syntax/ast.js'
import { FormulaError } from '../../syntax/ast.js'
import { legacyConverter, legacyKey } from '../../legacy/convert.js'
import { parse } from '../../syntax/parse.js'
import { canonical, mapAst, walk } from '../../syntax/print.js'
import { SHAPES, type Hole } from '../../syntax/shape.js'
import { escapeText } from '../xml/text.js'

/** A formula of the old kind as the dynamic one Excel 365 shows for it: undefined where the same (`legacyConverter`). */
type Legacy = (ast: Ast, sheet: string) => Ast | undefined

/**
 * The old kind's `<f>` element of a formula of a sheet, or undefined where only a dynamic array means it; `written`, the
 * formula's text as the dynamic array's element has it, which stays where the old kind reads it back as it is; `text`,
 * a formula's text in the file (`fileText`) for one whose `@` the old kind does itself.
 */
export function oldKind(formula: string, sheet: string, written: string, legacy: Legacy, text: (formula: string) => string | null): string | undefined {
  let ast: Ast
  try {
    ast = parse(formula)
  } catch (e) {
    if (e instanceof FormulaError) return undefined
    throw e
  }
  // a call of a LAMBDA or of a name (`Twice(B1:B3)`): the conversion does not know what its answer is, and Excel 365
  // saves one as a dynamic array; Excel before 365 has no LAMBDA
  let calls = false
  walk(ast, (n) => (calls ||= n.k === 'call'))
  if (calls) return undefined
  // the old kind does the implicit intersection `@` asks for where it puts one itself
  const bare = mapAst(ast, (n) => (n.k === 'at' ? n.a : undefined))
  if (bare !== ast && canonical(legacy(bare, sheet) ?? bare) === canonical(ast)) {
    const file = text(canonical(bare))
    return file === null ? undefined : `<f>${escapeText(file)}</f>`
  }
  // an `@` it does not put (`@A4`, of one cell) stays, as `_xlfn.SINGLE` the old kind reads back as `@`
  return legacy(ast, sheet) === undefined ? `<f>${written}</f>` : undefined
}

/** A formula's shape in a book's cache of shapes (`fileText`): its key and holes. */
export interface ShapeSeen {
  key: string
  holes: Hole[]
}

/** `oldKind` over a book's formulas, `shape` a formula's where it has one: formulas of a shape that convert alike (`legacyKey`) and have no `@` take one answer. */
export type OldKinds = (formula: string, sheet: string, written: string, text: (formula: string) => string | null, shape?: ShapeSeen) => string | undefined

/** The book's `oldKind`, its converter made over the book's names when a formula first asks. */
export function oldKinds(names: readonly NameInput[] | undefined): OldKinds {
  let legacy: Legacy | undefined
  const known = new Map<string, boolean>()
  return (formula, sheet, written, text, shape) => {
    legacy ??= legacyConverter(names)
    // a formula without `@` is of the old kind as written or not at all; the converter sees a name of the sheet's
    const key = shape && !formula.includes('@') ? `${sheet.length}:${sheet}${shape.key}|${legacyKey(shape.holes)}` : undefined
    const hit = key === undefined ? undefined : known.get(key)
    if (hit !== undefined) return hit ? `<f>${written}</f>` : undefined
    const f = oldKind(formula, sheet, written, legacy, text)
    if (key !== undefined && known.size < SHAPES) known.set(key, f !== undefined)
    return f
  }
}
