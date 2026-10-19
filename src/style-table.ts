/**
 * A book's style table (`BookInput.styles`), its styles found by their fields: a session's `patchCells` adds a style by
 * them, and `internStyle` gives a host a style's number without a session, by the same equality (the order of fields and
 * fields set to undefined do not count).
 */
import type { CellStyle } from './contract.js'
import { checkStyle, checkStyles } from './input/styles.js'

/** A key of a value independent of the order of its fields. */
function keyOf(v: unknown): string {
  if (v === null || typeof v !== 'object') return JSON.stringify(v)
  const o = v as Record<string, unknown>
  return `{${Object.keys(o)
    .filter((k) => o[k] !== undefined)
    .sort()
    .map((k) => `${JSON.stringify(k)}:${keyOf(o[k])}`)
    .join(',')}}`
}

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
    return i === undefined || keyOf(this.list[i]) === '{}'
  }

  indexOf(style: CellStyle): number {
    if (!this.keys) {
      this.keys = new Map()
      this.list.forEach((s, i) => this.keys!.has(keyOf(s)) || this.keys!.set(keyOf(s), i))
    }
    const k = keyOf(style)
    let i = this.keys.get(k)
    if (i === undefined) {
      if (!this.added) this.list = [...this.list]
      this.added = true
      i = this.list.push(style) - 1
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
  const keys = new Map<string, number>()
  styles.forEach((s, i) => {
    const k = keyOf(s)
    if (!keys.has(k)) keys.set(k, i)
  })
  const x = { keys, items: [...styles] }
  indexes.set(styles, x)
  return x
}

/** Whether a table holds the styles it held when it was read. */
const unchanged = (styles: readonly CellStyle[], items: readonly CellStyle[]) => styles.length === items.length && styles.every((s, i) => s === items[i])

/**
 * A style's index in a style table, as `patchCells` finds it: the index of the first equal style and the same array, or
 * a new array with a copy of the style appended and its index. Equal styles have the same fields with the same values,
 * in any order. A table is read once per array: the next call, given the returned table, reads only the added style. A
 * table changed in place (a style added, replaced or removed) is read again; a style changed in place goes unnoticed.
 */
export function internStyle(styles: readonly CellStyle[], style: CellStyle): { index: number; styles: CellStyle[] } {
  const k = keyOf(checkStyle(style, 'style'))
  let x = indexes.get(styles)
  if (!x || !unchanged(styles, x.items)) x = indexed(styles)
  let i = x.keys.get(k)
  // a number past the table or another style there: a key another table grown from the same one added, or a style changed
  if (i !== undefined && (i >= styles.length || keyOf(styles[i]) !== k)) i = (x = indexed(styles)).keys.get(k)
  if (i !== undefined) return { index: i, styles: styles as CellStyle[] }
  const next = [...styles, JSON.parse(JSON.stringify(style)) as CellStyle]
  x.keys.set(k, styles.length)
  indexes.set(next, { keys: x.keys, items: [...next] })
  return { index: styles.length, styles: next }
}
