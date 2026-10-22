/**
 * A file's sheet names that Excel would not take, repaired: `: \ / ? * [ ]` and an apostrophe at either end become `_`,
 * a name is cut to 31 characters, an empty one is `SheetN` by its place, and one reserved (`History`) or taken by
 * another sheet in any case gets a number, ` (2)`, as Excel numbers a copied sheet. Names Excel takes stay, and so does
 * the first of names that differ only in case.
 */
import { sheetNameProblem } from '../../input.js'
import { sheetKey } from '../../syntax.js'
import { rewriteFormula, renaming } from '../../structure/rewrite.js'

const MAX = 31

/** `s` cut to `n` UTF-16 units, without splitting a surrogate pair. */
function cut(s: string, n: number): string {
  if (s.length <= n) return s
  const last = s.charCodeAt(n - 1)
  return s.slice(0, last >= 0xd800 && last <= 0xdbff ? n - 1 : n)
}

export interface Repaired {
  /** The names, in order, as the book takes them. */
  names: string[]
  /** The key of a repaired name to its new name, for the first sheet with it: formulas refer to it so. */
  renamed: Map<string, string>
}

export function repairNames(names: string[]): Repaired {
  const keys = new Set<string>()
  // the first sheet with a name in any case, which formulas refer to by it
  const first = names.map((n) => keys.size < keys.add(sheetKey(n)).size)
  const kept = names.map((n, i) => first[i] && sheetNameProblem(n) === null)
  const taken = new Set(names.filter((_, i) => kept[i]).map(sheetKey))
  const renamed = new Map<string, string>()
  const out = names.map((n, i) => {
    if (kept[i]) return n
    const base = cut(n.replace(/[:\\/?*[\]]/g, '_'), MAX).replace(/^'|'$/g, '_') || `Sheet${i + 1}`
    let name = base
    for (let x = 2; taken.has(sheetKey(name)) || sheetNameProblem(name); x++) name = cut(base, MAX - ` (${x})`.length) + ` (${x})`
    taken.add(sheetKey(name))
    if (first[i]) renamed.set(sheetKey(n), name)
    return name
  })
  return { names: out, renamed }
}

/** A formula's references to the sheets `renamed` (by the key of their names in the file) rewritten to the new names. */
export function renamedIn(renamed: ReadonlyMap<string, string>): (formula: string) => string {
  const t = renaming(renamed)
  // the old names are no sheets of the book: references find them by key
  const sheets = new Map([...renamed.keys()].map((k) => [k, k]))
  return (formula) => rewriteFormula(formula, null, null, t, sheets)
}
