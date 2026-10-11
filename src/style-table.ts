/**
 * A book's style table (`BookInput.styles`), its styles found by their fields: a session's `patchCells` adds a style by
 * them, and `internStyle` gives a host a style's number without a session, by the same equality (the order of fields and
 * fields set to undefined do not count).
 */
import type { CellStyle } from './contract.js'
import { checkStyle, checkStyles } from './input.js'

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

/**
 * A style's number in a style table, as `patchCells` finds it: the first equal style's number and the same array, or a
 * new array with a copy of the style at the end and its number. Equal styles have the same fields with the same values,
 * in any order. One style at a time: the table's styles are compared anew on each call.
 */
export function internStyle(styles: readonly CellStyle[], style: CellStyle): { index: number; styles: CellStyle[] } {
  checkStyles(styles, 'styles')
  const t = new StyleTable(styles as CellStyle[])
  const index = t.indexOf(JSON.parse(JSON.stringify(checkStyle(style, 'style'))) as CellStyle)
  return { index, styles: t.table ?? (styles as CellStyle[]) }
}
