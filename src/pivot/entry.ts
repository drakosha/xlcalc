/**
 * An entry into a cell of a PivotTable's area as Excel takes it: text typed over an item's label names the item (an
 * item's name of the same field moves that item there), over a field's name names the field (another field's name on
 * the same axis swaps the two), over a value field's name names it, over a subtotal's label gives the field's subtotal
 * caption (`?` for the item's label), over `Grand Total`, `Row Labels` or `Column Labels` gives the PivotTable's caption,
 * over a Filters area's item shows that item. The entry becomes an edit of the PivotTable's definition
 * (`editPivotTable`); a value, an empty cell, a cleared label, a formula and a name already taken are refused.
 */
import type { BookInput, CellInput, PivotAxisField, PivotFieldPatch, PivotTable, PivotTablePatch, Refusal, StructureOp } from '../contract.js'
import { say, type EngineMessage } from '../locale/messages.js'
import { formatAddress, parseAddress } from '../model.js'
import { hasCell } from '../structure/area.js'
import { sheetIndex } from '../structure/rewrite.js'
import { pivotParts } from './area.js'
import { labeler, layoutPivot } from './layout.js'
import { savedResult } from './saved.js'

/** What an entry into a PivotTable's cell does: an edit of its definition, nothing (the same name), or Excel's refusal; null for a cell of no PivotTable. */
export type PivotEntry = { ok: true; op: StructureOp | null } | Refusal<'pivot'>

const same = (a: string, b: string) => a.toLocaleLowerCase() === b.toLocaleLowerCase()

/** The entry `input` into `cell` of `sheet`, as `PivotEntry`; null when the cell is in no PivotTable's area. */
export function pivotEntry(book: BookInput, sheet: string, cell: string, input: CellInput | null): PivotEntry | null {
  const s = sheetIndex(book, sheet)
  const sh = book.sheets[s]
  const { row, col } = parseAddress(cell)
  const table = sh?.pivotTables?.find((t) => pivotParts(t).some((x) => hasCell(x, row, col)))
  const cache = table && book.pivotCaches?.find((c) => c.id === table.cache)
  if (!table || !cache) return null
  const where = { sheet: sh.name, cell, pivot: table.name }
  const refuse = (m: EngineMessage): PivotEntry => ({ ok: false, code: 'pivot', message: say(book.locale, m) })
  const part = layoutPivot(savedResult(table, cache), book).cells.find((c) => formatAddress(c.row, c.col) === cell)?.part
  if (!part) return refuse({ key: 'pivotPart', ...where })
  if (input?.formula !== undefined) return refuse({ key: 'pivotFormula', ...where })
  const v = input?.value
  if (v === undefined || v === null || v === '') return refuse({ key: 'pivotClear', ...where })
  const { words, rules, fieldName, valueName, labelText } = labeler(table, cache, book)
  const text = typeof v === 'string' ? v : typeof v === 'object' ? v.date : rules.display(v, undefined, Infinity).text
  const edit = (set: PivotTablePatch): PivotEntry => ({ ok: true, op: { op: 'editPivotTable', sheet: sh.name, pivot: table.name, set } })
  const field = (f: number, patch: Omit<PivotFieldPatch, 'field'>): PivotTablePatch => ({ fields: [{ field: f, ...patch }] })
  // a value field's name: none of the cache's fields' names, nor another value field's
  const nameValue = (k: number): PivotEntry => {
    if (same(text, valueName(k))) return { ok: true, op: null }
    const taken = cache.fields.some((x) => same(x.name, text)) || (table.values ?? []).some((_, j) => j !== k && same(valueName(j), text))
    if (taken) return refuse({ key: 'pivotNameTaken', name: text })
    return edit({ values: (table.values ?? []).map((d, j) => (j === k ? { ...d, name: text } : d)) })
  }
  switch (part.kind) {
    case 'value':
      return nameValue(part.value)
    case 'grand':
      return edit({ captions: { ...table.captions, grandTotal: text } })
    case 'rowHeader':
    case 'columnHeader':
      return edit({ captions: { ...table.captions, [part.kind]: text } })
    case 'field':
      return part.field === 'values' ? edit({ captions: { ...table.captions, values: text } }) : nameField(table, part.field, text, fieldName, edit, refuse)
    case 'subtotal': {
      if (part.field === 'values') return refuse({ key: 'pivotPart', ...where })
      const label = labelText(part.field, part.item)
      const at = text.toLocaleLowerCase().indexOf(label.toLocaleLowerCase())
      return edit(field(part.field, { subtotalCaption: at < 0 ? text : `${text.slice(0, at)}?${text.slice(at + label.length)}` }))
    }
    case 'page': {
      const k = table.pages!.findIndex((p) => p.field === part.field)
      const items = table.fields[part.field].items ?? []
      const all = same(text, words.all)
      const j = items.findIndex((_, i) => same(labelText(part.field, i), text))
      if (!all && j < 0) return refuse({ key: 'pivotNotItem', item: text })
      return edit({ pages: table.pages!.map((p, i) => (i !== k ? p : all ? { field: p.field } : { field: p.field, item: j })) })
    }
    case 'item': {
      if (part.field === 'values') return nameValue(part.item)
      const f = part.field
      if (same(labelText(f, part.item), text)) return { ok: true, op: null }
      const items = table.fields[f].items ?? []
      const j = items.findIndex((_, i) => i !== part.item && same(labelText(f, i), text))
      // another item's name moves that item before this one; a new name is the item's caption
      if (j >= 0) {
        const moved = items.filter((_, i) => i !== j)
        moved.splice(moved.indexOf(items[part.item]), 0, items[j])
        return edit(field(f, { items: moved }))
      }
      return edit(field(f, { items: items.map((it, i) => (i === part.item ? { ...it, caption: text } : it)) }))
    }
  }
}

/** A field named anew: another field's name on its axis swaps the two, another field's name elsewhere is refused. */
function nameField(table: PivotTable, f: number, text: string, fieldName: (f: PivotAxisField) => string, edit: (set: PivotTablePatch) => PivotEntry, refuse: (m: EngineMessage) => PivotEntry): PivotEntry {
  if (same(fieldName(f), text)) return { ok: true, op: null }
  for (const axis of ['rows', 'columns'] as const) {
    const fields = table[axis] ?? []
    if (!fields.includes(f)) continue
    const other = fields.find((g) => g !== f && same(fieldName(g), text))
    if (other !== undefined) return edit({ [axis]: fields.map((g) => (g === f ? other : g === other ? f : g)) })
  }
  if (table.fields.some((_, g) => g !== f && same(fieldName(g), text))) return refuse({ key: 'pivotNameTaken', name: text })
  return edit({ fields: [{ field: f, name: text }] })
}
