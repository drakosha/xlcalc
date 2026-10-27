/**
 * A book's style table (`BookInput.styles`), its styles found by their fields: a session's `patchCells` adds a style by
 * them, and `internStyle` gives a host a style's number without a session, by the same equality (the order of fields,
 * fields set to undefined and a protection's defaults, `{ locked: true }`, do not count).
 */
import type { CellStyle } from './contract.js'
import { styleTaken } from './input/protection.js'
import { checkStyle, checkStyles } from './input/styles.js'

/** A key of a value independent of the order of its fields. */
export function keyOf(v: unknown): string {
  if (v === null || typeof v !== 'object') return JSON.stringify(v)
  const o = v as Record<string, unknown>
  return `{${Object.keys(o)
    .filter((k) => o[k] !== undefined)
    .sort()
    .map((k) => `${JSON.stringify(k)}:${keyOf(o[k])}`)
    .join(',')}}`
}

/** A style's key: its fields as the workbook holds them (`styleTaken`), in any order. */
export const styleKey = (s: CellStyle): string => keyOf(styleTaken(s))

/** Each style's key to its first index in `styles` (read back to front, so the first stays). */
const firstKeys = (styles: readonly CellStyle[]): Map<string, number> => new Map(styles.map((s, i) => [styleKey(s), i] as [string, number]).reverse())

/** The book's styles, found by their fields: an equal style's number, or a new one added at the end. */
export class StyleTable {
  private keys: Map<string, number> | null = null
  private list: CellStyle[]
  added = false

  constructor(private readonly base: CellStyle[] | undefined) {
    this.list = base ?? []
  }

  /** The style of a number; the plain style for none. */
  at(i: number | undefined): CellStyle {
    return i === undefined ? {} : this.list[i]
  }

  isPlain(i: number | undefined): boolean {
    return i === undefined || styleKey(this.list[i]) === '{}'
  }

  indexOf(style: CellStyle): number {
    this.keys ??= firstKeys(this.list)
    const taken = styleTaken(style)
    const k = keyOf(taken)
    let i = this.keys.get(k)
    if (i === undefined) {
      if (!this.added) this.list = [...this.list]
      this.added = true
      i = this.list.push(taken) - 1
      this.keys.set(k, i)
    }
    return i
  }

  /** The new table, when styles were added. */
  get table(): CellStyle[] | undefined {
    return this.added ? this.list : undefined
  }
}

/** A style table `internStyle` read: the first number of each style's key, and its styles then, to see it changed in place. */
interface Indexed {
  keys: Map<string, number>
  items: readonly CellStyle[]
}

/**
 * The tables `internStyle` was given or answered, by the array: a table answered shares the keys of the one it grew from,
 * so a chain of calls, each given the last answer's table, reads no table twice.
 */
const indexes = new WeakMap<readonly CellStyle[], Indexed>()

/** A table read anew, its styles checked: its own keys. */
function indexed(styles: readonly CellStyle[]): Indexed {
  checkStyles(styles, 'styles')
  const x = { keys: firstKeys(styles), items: [...styles] }
  indexes.set(styles, x)
  return x
}

/** Whether a table holds the styles it held when it was read. */
const unchanged = (styles: readonly CellStyle[], items: readonly CellStyle[]) => styles.length === items.length && styles.every((s, i) => s === items[i])

/**
 * A style's index in a style table, as `patchCells` finds it: the index of the first equal style and the same array, or
 * a new array with a copy of the style appended and its index. Equal styles have the same fields with the same values,
 * in any order, a protection's defaults aside (`{ protection: { locked: true } }` is `{}`). A table is read once per array: the next call, given the returned table, reads only the added style. A
 * table changed in place (a style added, replaced or removed) is read again; a style changed in place goes unnoticed.
 */
export function internStyle(styles: readonly CellStyle[], style: CellStyle): { index: number; styles: CellStyle[] } {
  const taken = styleTaken(checkStyle(style, 'style'))
  const k = keyOf(taken)
  let x = indexes.get(styles)
  if (!x || !unchanged(styles, x.items)) x = indexed(styles)
  let i = x.keys.get(k)
  // a number past the table or another style there: a key another table grown from the same one added, or a style changed
  if (i !== undefined && (i >= styles.length || styleKey(styles[i]) !== k)) i = (x = indexed(styles)).keys.get(k)
  if (i !== undefined) return { index: i, styles: styles as CellStyle[] }
  const next = [...styles, JSON.parse(JSON.stringify(taken)) as CellStyle]
  x.keys.set(k, styles.length)
  indexes.set(next, { keys: x.keys, items: [...next] })
  return { index: styles.length, styles: next }
}
