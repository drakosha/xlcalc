/** Cached ranges indexed by their referenced bounds, including their unoccupied cells. */
import { AreaIndex } from './area-index.js'
import type { Area } from './spill/state.js'

interface Entry extends Area { key: string }

export class RangeCache<T> extends Map<string, T> {
  private readonly ranges = new Map<string, Entry>()
  private readonly index = new AreaIndex<Entry>()

  constructor() { super() }

  override set(key: string, value: T): this {
    if (!this.has(key)) {
      const [s, r1, c1, r2, c2] = key.split(':').map(Number)
      const entry = { key, s, r1, c1, r2, c2 }
      this.ranges.set(key, entry)
      this.index.add(entry)
    }
    return super.set(key, value)
  }

  override delete(key: string): boolean {
    const entry = this.ranges.get(key)
    if (!entry) return false
    this.index.delete(entry)
    this.ranges.delete(key)
    return super.delete(key)
  }

  override clear(): void {
    super.clear()
    this.ranges.clear()
    this.index.clear()
  }

  invalidate(area: Area): void {
    this.index.forEach(area, (entry) => this.delete(entry.key))
  }
}
