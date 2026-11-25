import { quoteSheetIfNeeded } from '../a1.js'
import { formulaText, type Formula } from '../xlsb/formula.js'
import type { Context } from '../xlsb/parts.js'
import type { Payload } from '../xlsb/records.js'
import { parsedFormula } from './formula.js'

interface ItemRef {
  field: number
  item: number
}

interface CalculatedItem extends ItemRef {
  formula: string
}

interface Name {
  field?: number
  item?: ItemRef
  paired: boolean
}

interface SavedFormula {
  formula: Formula
  field?: number
  target?: ItemRef
  nameCount: number
  names: Name[]
}

interface FieldValues {
  name: string
  items: readonly { type: string; value?: number | string }[]
}

export class CacheFormulas {
  private readonly entries: SavedFormula[] = []
  private pending?: SavedFormula
  private reference?: { field: number; count: number }

  take(type: number, payload: Payload, field: number | undefined): boolean {
    if (type === 0x00f9) {
      const size = payload.u16()
      const nameCount = payload.u16()
      const entry = { formula: parsedFormula(payload, size), field, nameCount, names: [] }
      this.entries.push(entry)
      this.pending = entry
      this.reference = undefined
      return true
    }
    const entry = this.pending
    if (!entry) return false
    if (type === 0x00f6) {
      const flags = payload.u16()
      const index = payload.i16()
      payload.skip(2)
      const pairCount = payload.u16()
      entry.names.push({
        field: !(flags & 2) && index >= 0 ? index : undefined,
        paired: pairCount === 1,
      })
    } else if (type === 0x00f8) {
      const name = entry.names.at(-1)
      const item: ItemRef = { field: payload.u16(), item: payload.i16() }
      payload.skip(2)
      const flags = payload.u16()
      if (name?.paired && !(flags & 8) && item.item >= 0) name.item = item
    } else if (type === 0x00f2) {
      payload.skip(2)
      const index = payload.u16() & 1023
      payload.skip(2)
      this.reference = { field: index, count: payload.u16() }
    } else if (type === 0x00f5) {
      if (this.reference?.count === 1 && !entry.target) entry.target = { field: this.reference.field, item: payload.u16() }
      else entry.target = undefined
    } else if (type === 0x0103) {
      payload.skip(2)
      const appliesTo = payload.i16()
      if (entry.target && appliesTo >= 0 && appliesTo !== entry.target.field) entry.target = undefined
      this.pending = undefined
    } else if (type !== 0x00f0) return false // SxRule starts the target of a calculated item.
    return true
  }

  compile(fields: readonly FieldValues[], context: Context): { fields: Map<number, string>; items: CalculatedItem[] } | undefined {
    const formulas = new Map<number, string>()
    const items: CalculatedItem[] = []
    for (const entry of this.entries) {
      if (entry.nameCount !== entry.names.length) return undefined
      const names = entry.names.map(({ field, item }) => {
        if (field !== undefined && fields[field]) return quoteSheetIfNeeded(fields[field].name)
        if (!item) return '#REF!'
        const source = fields[item.field]
        const value = source?.items[item.item]
        if (!value) return '#REF!'
        const label = value.type === 'b' ? value.value ? 'TRUE' : 'FALSE' : String(value.value ?? '')
        return `${quoteSheetIfNeeded(source.name)}[${quoteSheetIfNeeded(label)}]`
      })
      const parsed = formulaText(entry.formula, { ...context, pivotNames: names })
      if (!parsed || !('text' in parsed)) return undefined
      if (entry.field !== undefined) {
        formulas.set(entry.field, parsed.text)
        continue
      }
      const target = entry.target
      if (!target || !fields[target.field]?.items[target.item]) return undefined
      items.push({ ...target, formula: parsed.text })
    }
    return { fields: formulas, items }
  }
}
