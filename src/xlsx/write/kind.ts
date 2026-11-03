/**
 * The oldest form of a formula written: a formula of the old kind (`<f>`), which Excel before 365 reads, wherever Excel
 * reads it as the workbook computes it, that is where reading it back as the old kind (`@` where Excel 365 shows it,
 * `legacy/convert.ts`) gives what reading it as a dynamic array gives; else none, and it is written as a dynamic array. Excel 365 saves a
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
import { fromFile } from '../formulas/from-file.js'
import { inputNameFor } from '../../names.js'

/** A formula of the old kind as the dynamic one Excel 365 shows for it: undefined where the same (`legacyConverter`). */
type Legacy = (ast: Ast, sheet: string) => Ast | undefined

/**
 * The old kind's `<f>` element of a formula of a sheet, or undefined where only a dynamic array means it: where the
 * reader reads a text of the old kind back (`fromFile`, then `@` as Excel 365 shows it) as it reads `file`, the
 * formula's text in a dynamic array. The text is `file` itself, or, where the formula has `@`, its text without the `@`
 * the old kind does itself (`text`, a formula's file text); an `@` of a cell it does not put (`@A4`) stays as `_xlfn.SINGLE`.
 */
export function oldKind(formula: string, sheet: string, file: string, legacy: Legacy, text: (formula: string) => string | null, computed: (name: string, sheet: string) => boolean = () => false): string | undefined {
  const ast = parsed(formula)
  if (!ast) return undefined
  // a call of a LAMBDA or of a name (`Twice(B1:B3)`): the conversion does not know what its answer is, and Excel 365
  // saves one as a dynamic array; Excel before 365 has no LAMBDA. A reference into another workbook: Excel computes some
  // of the old kind otherwise (`SUM([1]Sheet1!A:A)` #REF! as a dynamic array, 6 of the old kind) and refused a file of
  // them. A name of a formula that is no literal or reference (`computed`): the old kind of `SUMIF(sq,">10")` over
  // `sq` = `A1:A3*10` is #VALUE!, the dynamic one #SPILL!, which the text Excel shows does not tell
  let unknown = false
  walk(ast, (n) => (unknown ||= n.k === 'call' || n.k === 'xref' || n.k === 'xname' || (n.k === 'err' && n.book !== undefined) || (n.k === 'name' && computed(n.name, n.sheet ?? sheet))))
  if (unknown) return undefined
  const dynamic = parsed(fromFile(file).formula)
  if (!dynamic) return undefined
  const want = canonical(dynamic)
  const bare = mapAst(ast, (n) => (n.k === 'at' ? n.a : undefined))
  // `_xlfn.SINGLE` of the old kind only around a cell, as Excel saves a typed `@A4+1`: around more, Excel takes the ranges
  // inside as one value each too (`_xlfn.SINGLE((A1:A3*10))` comes back `@(@A1:A3*10)`), which the conversion does not see
  let single = true
  walk(ast, (n) => (single &&= n.k !== 'at' || (n.a.k === 'ref' && n.a.ref.kind === 'cell' && !n.a.ref.spill)))
  for (const t of [bare === ast ? null : text(canonical(bare)), single ? file : null]) {
    const old = t === null ? undefined : parsed(fromFile(t).formula)
    if (old && canonical(legacy(old, sheet) ?? old) === want) return `<f>${escapeText(t!)}</f>`
  }
  return undefined
}

/** A formula's tree, undefined where it does not parse; `union` for a name's formula. */
function parsed(formula: string, union = false): Ast | undefined {
  try {
    return parse(formula, { union })
  } catch (e) {
    if (e instanceof FormulaError) return undefined
    throw e
  }
}

/** A formula's shape in a book's cache of shapes (`fileText`): its key and holes. */
export interface ShapeSeen {
  key: string
  holes: Hole[]
}

/** `oldKind` over a book's formulas, `shape` a formula's where it has one: formulas of a shape that convert alike (`legacyKey`) and have no `@` take one answer. */
export type OldKinds = (formula: string, sheet: string, file: string, text: (formula: string) => string | null, shape?: ShapeSeen) => string | undefined

/** The book's `oldKind`, its converter made over the book's names when a formula first asks. */
export function oldKinds(names: readonly NameInput[] = []): OldKinds {
  let legacy: Legacy | undefined
  const known = new Map<string, boolean>()
  // a name whose formula is no literal and no reference
  const computed = (name: string, sheet: string) => {
    const found = inputNameFor(names, sheet, name)
    const def = found && parsed(found.ref, true)
    const lit = def?.k === 'un' && def.op === '-' ? def.a : def
    return !!def && !['num', 'str', 'bool', 'err', 'ref', 'ref3d'].includes(lit!.k)
  }
  return (formula, sheet, file, text, shape) => {
    legacy ??= legacyConverter(names)
    // a formula without `@` is of the old kind as written or not at all; the converter sees a name of the sheet's
    const key = shape && !formula.includes('@') ? `${sheet.length}:${sheet}${shape.key}|${legacyKey(shape.holes)}` : undefined
    const hit = key === undefined ? undefined : known.get(key)
    if (hit !== undefined) return hit ? `<f>${escapeText(file)}</f>` : undefined
    const f = oldKind(formula, sheet, file, legacy, text, computed)
    if (key !== undefined && known.size < SHAPES) known.set(key, f !== undefined)
    return f
  }
}
