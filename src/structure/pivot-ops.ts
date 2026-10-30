/**
 * Operations on a book's PivotTables (`editStructure`): one created over a source or a cache, its definition changed,
 * renamed, moved, deleted, refreshed; slicers and timelines added, changed and deleted. After a change of what a
 * PivotTable shows, `xlcalc/pivot` lays it out again when loaded (`pivot/hook.ts`); without it, the PivotTable keeps its
 * cells and its cache is marked `stale`, which a file asks Excel to refresh on opening.
 */
import type { BookInput, PivotCache, PivotCacheField, PivotRef, PivotSource, PivotTable, PivotTablePatch, StructureOp } from '../contract.js'
import { caseKey } from '../collate.js'
import { InputError } from '../input-error.js'
import { checkSheetOf } from '../input.js'
import { checkString, object, onlyKeys } from '../input/basic.js'
import { pivotRefs } from '../input/pivot-refs.js'
import { checkAddress } from '../input/range.js'
import { say, type EngineMessage } from '../locale/messages.js'
import { withField } from '../looks/sheet.js'
import { formatAddress, parseAddress } from '../model.js'
import { pivotRefresh } from '../pivot/hook.js'
import { valueFieldName } from '../pivot/names.js'
import { tableKey, tablesOf } from '../tables/model.js'
import { quoteSheet, sheetKey } from '../syntax.js'
import { areaText, meets, parseRange, rectText, type Rect } from './area.js'
import type { Applied, Geometry } from './geometry.js'
import { moveRange } from './move.js'
import { checkedPivots, movedPivot, patched, pivotArea, pivotsOn, relaid, withoutPivot } from './pivots.js'
import { sheetIndex } from './rewrite.js'
import { insertSheet } from './sheets.js'
import type { ValuesIn } from './values.js'

export type PivotOp = Extract<StructureOp, { op: `${string}Pivot${string}` | `${string}Slicer` | `${string}Timeline` }>

/** The fields of each operation, for `editStructure`'s check of its shape. */
export const PIVOT_FIELDS: Record<PivotOp['op'], string[]> = {
  addPivotTable: ['sheet', 'cell', 'newSheet', 'name', 'source', 'definition'],
  editPivotTable: ['sheet', 'pivot', 'set', 'source'],
  renamePivotTable: ['sheet', 'pivot', 'name'],
  movePivotTable: ['sheet', 'pivot', 'to'],
  deletePivotTable: ['sheet', 'pivot'],
  refreshPivots: ['pivots'],
  addSlicer: ['sheet', 'slicer', 'cache'],
  editSlicer: ['slicer', 'set', 'cache'],
  deleteSlicer: ['slicer'],
  addTimeline: ['sheet', 'timeline', 'cache'],
  editTimeline: ['timeline', 'set', 'cache'],
  deleteTimeline: ['timeline'],
}

/** The fields of a PivotTable a patch sets: all but its identity, place and saved layout. */
const PATCHED = ['rows', 'columns', 'pages', 'values', 'valuesOnRows', 'valuesPosition', 'valuesRow', 'filters', 'grandTotals', 'layout', 'indent', 'repeatLabels', 'mergeLabels', 'fieldHeaders', 'captions', 'errors', 'empty', 'pagesAcrossFirst', 'pagesWrap', 'subtotalHiddenItems', 'asteriskTotals', 'customListSort', 'multipleFieldFilters', 'showDrill', 'autofitColumns', 'preserveFormatting', 'classicLayout', 'style', 'enableDrill', 'altText', 'formats', 'fields']

/** A new PivotTable's place before its first layout: three columns and eighteen rows, as Excel draws an empty one. */
const EMPTY_ROWS = 18
const EMPTY_COLUMNS = 3

/** A new PivotTable's style and options, as Excel 365 makes one. */
const NEW_STYLE = { name: 'PivotStyleLight16', rowHeaders: true, columnHeaders: true, lastColumn: true }

const refuse = (book: BookInput, m: EngineMessage): Applied => ({ ok: false, code: 'pivot', message: say(book.locale, m) })

/** A PivotTable of the book by sheet and name: its sheet's index and its index there. */
function pivotAt(book: BookInput, sheet: unknown, name: unknown, path: string): { s: number; k: number; p: PivotTable } {
  const s = book.sheets.indexOf(checkSheetOf(book, sheet, `${path}.sheet`))
  const n = checkString(name, `${path}.pivot`)
  const k = (book.sheets[s].pivotTables ?? []).findIndex((p) => caseKey(p.name) === caseKey(n))
  if (k < 0) throw new InputError(`${path}.pivot`, `unknown PivotTable ${JSON.stringify(n)} on ${JSON.stringify(book.sheets[s].name)}`)
  return { s, k, p: book.sheets[s].pivotTables![k] }
}

/** The book with the PivotTable at `k` of sheet `s` replaced by `next` (removed when null). */
function withPivot(book: BookInput, s: number, k: number, next: PivotTable | null): BookInput {
  const list = [...(book.sheets[s].pivotTables ?? [])]
  if (next) list.splice(k, 1, next)
  else list.splice(k, 1)
  const sheets = [...book.sheets]
  sheets[s] = withField(sheets[s], 'pivotTables', list.length ? list : undefined)
  return { ...book, sheets }
}


/**
 * A PivotTable with a patch of its definition, checked: one axis a field, the value fields named as Excel names them
 * (`Sum of Amount`, a second `Sum of Amount2`).
 */
function withDefinition(p: PivotTable, c: PivotCache, set: unknown, path: string): PivotTable {
  const patch = object(set, path, 'a patch of a PivotTable such as { rows: [0], values: [{ field: 1 }] }')
  onlyKeys(patch, path, PATCHED)
  const { fields, ...rest } = patch as PivotTablePatch
  let out = patched(p, rest)
  if (fields !== undefined) {
    if (!Array.isArray(fields)) throw new InputError(`${path}.fields`, 'expected an array of field patches such as [{ field: 0, items: [...] }]')
    const list = [...out.fields]
    fields.forEach((f, i) => {
      const fp = object(f, `${path}.fields[${i}]`, 'a field patch such as { field: 0, subtotals: [] }')
      const n = fp.field
      if (typeof n !== 'number' || !Number.isInteger(n) || n < 0 || n >= list.length) throw new InputError(`${path}.fields[${i}].field`, `expected the index of a field of the cache, 0 to ${list.length - 1}, got ${JSON.stringify(n)}`)
      const { field: _, ...own } = fp
      list[n] = patched(list[n], own)
    })
    out = { ...out, fields: list }
  }
  // a field lies on one axis
  const seen = new Map<number | 'values', string>()
  for (const [axis, list] of [['rows', out.rows], ['columns', out.columns], ['pages', out.pages?.map((g) => g.field)]] as const)
    (list ?? []).forEach((f, i) => {
      if (seen.has(f)) throw new InputError(`${path}.${axis}[${i}]`, `the field ${JSON.stringify(f)} is on ${seen.get(f)} already: a field lies on one axis`)
      seen.set(f, axis)
    })
  if (out.values) {
    const names = new Set<string>()
    out = {
      ...out,
      values: out.values.map((d) => {
        let name = d.name ?? valueFieldName(d.function ?? 'sum', c.fields[d.field]?.name ?? '')
        if (d.name === undefined) for (let n = 2; names.has(caseKey(name)); n++) name = `${valueFieldName(d.function ?? 'sum', c.fields[d.field]?.name ?? '')}${n}`
        names.add(caseKey(name))
        return d.name === undefined ? { ...d, name } : d
      }),
    }
  }
  return out
}

/** The text of a header cell as a field's name: a number or a boolean as Excel shows it plainly; empty without a value. */
const headerText = (v: unknown) => (v === null || v === undefined || v === '' ? '' : typeof v === 'object' ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v))

/** Field names unique as Excel makes them: a repeated name with 2, 3… after it. */
function uniqueNames(names: string[]): string[] {
  const seen = new Set<string>()
  return names.map((n) => {
    let name = n
    for (let k = 2; seen.has(caseKey(name)); k++) name = `${n}${k}`
    seen.add(caseKey(name))
    return name
  })
}

/** The cells of a source in the book: its sheet and rectangle, its header row first; null for a source the book does not hold. */
function sourceCells(book: BookInput, src: PivotSource): { sheet: string; x: Rect } | null {
  if ('book' in src && src.book !== undefined) return null
  if (src.type === 'range') return sheetIndex(book, src.sheet) < 0 ? null : (x => x && { sheet: book.sheets[sheetIndex(book, src.sheet)].name, x })(parseRange(src.range))
  if (src.type === 'table') {
    const t = tablesOf(book).byName.get(tableKey(src.table))
    return t ? { sheet: book.sheets[t.s].name, x: { r1: t.r1, r2: t.totals ? t.r2 - 1 : t.r2, c1: t.c1, c2: t.c2 } } : null
  }
  if (src.type === 'name') {
    const n = book.names?.find((x) => x.scope === undefined && caseKey(x.name) === caseKey(src.name))
    const m = n && /^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/.exec(n.ref.replace(/^=/, ''))
    if (!m) return null
    const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]
    const x = parseRange(`${m[3]}${m[4]}:${m[5]}${m[6]}`)
    return x && sheetIndex(book, sheet) >= 0 ? { sheet: book.sheets[sheetIndex(book, sheet)].name, x } : null
  }
  return null
}

/** A source's fields by its header row's values; a refusal for an empty header or a source without a row of data. */
function sourceFields(book: BookInput, src: PivotSource, values: ValuesIn, pivot: string): PivotCacheField[] | Applied | null {
  const at = sourceCells(book, src)
  if (!at) return null
  const label = src.type === 'range' ? `${quoteSheet(at.sheet)}!${rectText(at.x)}` : src.type === 'table' ? src.table : src.type === 'name' ? src.name : ''
  if (at.x.r2 <= at.x.r1) return refuse(book, { key: 'pivotSourceRows', source: label })
  const head = new Map(values(at.sheet, { ...at.x, r2: at.x.r1 }).map(([, c, v]) => [c, headerText(v)]))
  const names: string[] = []
  for (let c = at.x.c1; c <= at.x.c2; c++) {
    const n = head.get(c) ?? ''
    if (!n) return refuse(book, { key: 'pivotFieldName', pivot })
    names.push(n)
  }
  return uniqueNames(names).map((name) => ({ name }))
}

/** A refusal of a PivotTable's place over cells that are not empty, a table or another PivotTable. */
function placeRefusal(book: BookInput, s: number, x: Rect, pivot: string, own?: PivotTable): Applied | null {
  const sheet = book.sheets[s]
  const range = areaText(x)
  const other = pivotsOn(book, s).find((t) => t.p !== own && meets(t.x, x))
  if (other) return refuse(book, { key: 'pivotOverPivot', pivot, other: other.p.name })
  const table = tablesOf(book).list.find((t) => t.s === s && meets(t, x))
  if (table) return refuse(book, { key: 'pivotOverTable', pivot, table: table.name })
  const ownArea = own && pivotArea(own)
  for (const [a, cell] of Object.entries(sheet.cells)) {
    if (cell.value === undefined && cell.formula === undefined) continue
    const { row, col } = parseAddress(a)
    if (row >= x.r1 && row <= x.r2 && col >= x.c1 && col <= x.c2 && !(ownArea && row >= ownArea.r1 && row <= ownArea.r2 && col >= ownArea.c1 && col <= ownArea.c2)) return refuse(book, { key: 'pivotOverData', sheet: sheet.name, range, pivot })
  }
  return null
}

/** The first name `prefix1`, `prefix2`… no PivotTable of the book has. */
function freeName(book: BookInput, prefix: string): string {
  const taken = new Set(book.sheets.flatMap((s) => (s.pivotTables ?? []).map((p) => caseKey(p.name))))
  let n = 1
  while (taken.has(caseKey(`${prefix}${n}`))) n++
  return `${prefix}${n}`
}



function addPivotTable(book: BookInput, o: Record<string, unknown>, values: ValuesIn): Applied {
  const name = o.name === undefined ? undefined : checkString(o.name, 'op.name')
  let start = book
  let geometry: Geometry = []
  if (o.newSheet !== undefined) {
    const n = object(o.newSheet, 'op.newSheet', 'a new sheet such as { before: "Data" }')
    onlyKeys(n, 'op.newSheet', ['before'])
    const r = insertSheet(book, checkString(o.sheet, 'op.sheet'), n.before === undefined ? undefined : checkSheetOf(book, n.before, 'op.newSheet.before').name)
    if (!r.ok) return r
    ;[start, geometry] = [r.book, r.geometry]
  }
  const s = start.sheets.indexOf(checkSheetOf(start, o.sheet, 'op.sheet'))
  const at = parseAddress(o.cell === undefined && o.newSheet !== undefined ? 'A3' : checkAddress(o.cell, 'op.cell'))
  const pivotName = name ?? freeName(start, 'PivotTable')
  if (!pivotName) return refuse(start, { key: 'pivotNameEmpty' })
  if (start.sheets[s].pivotTables?.some((p) => caseKey(p.name) === caseKey(pivotName))) return refuse(start, { key: 'pivotTableNameTaken', sheet: start.sheets[s].name, name: pivotName })
  const source = object(o.source, 'op.source', 'a source such as { type: "range", sheet: "Data", range: "A1:C21" } or { cache: 1 }')
  let cache: PivotCache
  let caches = start.pivotCaches ?? []
  if ('cache' in source) {
    onlyKeys(source, 'op.source', ['cache'])
    const c = caches.find((x) => x.id === source.cache)
    if (!c) throw new InputError('op.source.cache', `expected the id of a pivot cache of the workbook, got ${JSON.stringify(source.cache)}`)
    cache = c
  } else {
    const src = source as unknown as PivotSource
    const fields = sourceFields(start, src, values, pivotName)
    if (fields && !Array.isArray(fields)) return fields
    cache = { id: caches.reduce((m, c) => Math.max(m, c.id), 0) + 1, source: src, fields: fields ?? [] }
    caches = [...caches, cache]
  }
  const x = { r1: at.row, c1: at.col, r2: at.row + EMPTY_ROWS - 1, c2: at.col + EMPTY_COLUMNS - 1 }
  const over = placeRefusal(start, s, x, pivotName)
  if (over) return over
  let p: PivotTable = { name: pivotName, cache: cache.id, location: { range: areaText(x), firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 0 }, fields: cache.fields.map(() => ({})), indent: 0, multipleFieldFilters: false, style: NEW_STYLE }
  if (o.definition !== undefined) p = withDefinition(p, cache, o.definition, 'op.definition')
  const sheets = [...start.sheets]
  sheets[s] = withField(sheets[s], 'pivotTables', [...(sheets[s].pivotTables ?? []), p])
  const out = checkedPivots({ ...start, sheets, pivotCaches: caches }, 'op')
  return relaid(out, [{ sheet: sheets[s].name, name: pivotName }], values, true, geometry, 'refuse')
}

function editPivotTable(book: BookInput, o: Record<string, unknown>, values: ValuesIn): Applied {
  const { s, k, p } = pivotAt(book, o.sheet, o.pivot, 'op')
  let caches = book.pivotCaches ?? []
  let c = caches.find((x) => x.id === p.cache)!
  let next = p
  if (o.source !== undefined) {
    // Change Data Source: a cache of its own when another PivotTable or a slicer shares it
    const src = object(o.source, 'op.source', 'a source such as { type: "range", sheet: "Data", range: "A1:C21" }') as unknown as PivotSource
    const fields = sourceFields(book, src, values, p.name)
    if (fields && !Array.isArray(fields)) return fields
    const shared = book.sheets.some((x) => x.pivotTables?.some((q) => q !== p && q.cache === p.cache)) || book.slicerCaches?.some((x) => x.cache === p.cache) || book.timelineCaches?.some((x) => x.cache === p.cache)
    const id = shared ? caches.reduce((m, x) => Math.max(m, x.id), 0) + 1 : c.id
    // the PivotTable's fields kept by name, as Excel keeps those the new source still has
    const names = (fields ?? c.fields).map((f) => caseKey(f.name))
    const kept = new Map(c.fields.map((f, i) => [caseKey(f.name), i]))
    const fresh: PivotCache = { id, source: src, fields: fields ?? c.fields }
    caches = shared ? [...caches, fresh] : caches.map((x) => (x === c ? fresh : x))
    const map = (i: number) => names.indexOf(caseKey(c.fields[i]?.name ?? ''))
    next = {
      ...p,
      cache: id,
      fields: names.map((n) => (kept.has(n) ? { ...p.fields[kept.get(n)!], items: undefined } : {})),
      ...(p.rows ? { rows: p.rows.map((f) => (f === 'values' ? f : map(f))).filter((f) => f !== -1) } : {}),
      ...(p.columns ? { columns: p.columns.map((f) => (f === 'values' ? f : map(f))).filter((f) => f !== -1) } : {}),
      ...(p.pages ? { pages: p.pages.flatMap((g) => (map(g.field) < 0 ? [] : [{ field: map(g.field) }])) } : {}),
      ...(p.values ? { values: p.values.flatMap((d) => (map(d.field) < 0 ? [] : [{ ...d, field: map(d.field) }])) } : {}),
    }
    for (const f of next.fields) if (f.items === undefined) delete f.items
    c = fresh
  }
  if (o.set !== undefined) next = withDefinition(next, c, o.set, 'op.set')
  const out = checkedPivots({ ...withPivot(book, s, k, next), pivotCaches: caches }, 'op')
  return relaid(out, [{ sheet: book.sheets[s].name, name: p.name }], values, o.source !== undefined)
}

/** The slicers' and timelines' caches naming the PivotTable `from` by `to`. */
function renamedRefs(book: BookInput, from: PivotRef, to: PivotRef): BookInput {
  const is = (r: PivotRef) => sheetKey(r.sheet) === sheetKey(from.sheet) && caseKey(r.name) === caseKey(from.name)
  const refs = <C extends { pivots: PivotRef[] }>(list: C[]) => list.map((c) => (c.pivots.some(is) ? { ...c, pivots: c.pivots.map((r) => (is(r) ? to : r)) } : c))
  const out = { ...book }
  if (book.slicerCaches) out.slicerCaches = refs(book.slicerCaches)
  if (book.timelineCaches) out.timelineCaches = refs(book.timelineCaches)
  return out
}

function renamePivotTable(book: BookInput, o: Record<string, unknown>): Applied {
  const { s, k, p } = pivotAt(book, o.sheet, o.pivot, 'op')
  const name = checkString(o.name, 'op.name')
  if (!name) return refuse(book, { key: 'pivotNameEmpty' })
  if (book.sheets[s].pivotTables!.some((q, i) => i !== k && caseKey(q.name) === caseKey(name))) return refuse(book, { key: 'pivotTableNameTaken', sheet: book.sheets[s].name, name })
  let out = withPivot(book, s, k, { ...p, name })
  // the PivotTable's conditional formats name it
  const rules = out.sheets[s].conditionalFormats
  if (rules?.some((f) => f.pivot && caseKey(f.pivot.pivot) === caseKey(p.name))) {
    const sheets = [...out.sheets]
    sheets[s] = { ...sheets[s], conditionalFormats: rules.map((f) => (f.pivot && caseKey(f.pivot.pivot) === caseKey(p.name) ? { ...f, pivot: { ...f.pivot, pivot: name } } : f)) }
    out = { ...out, sheets }
  }
  return { ok: true, book: renamedRefs(out, { sheet: book.sheets[s].name, name: p.name }, { sheet: book.sheets[s].name, name }), geometry: [] }
}

function movePivotTable(book: BookInput, o: Record<string, unknown>, values: ValuesIn): Applied {
  const { s, p } = pivotAt(book, o.sheet, o.pivot, 'op')
  const to = object(o.to, 'op.to', 'a cell such as { sheet: "P", cell: "E3" }')
  onlyKeys(to, 'op.to', ['sheet', 'cell'])
  const t = book.sheets.indexOf(checkSheetOf(book, to.sheet, 'op.to.sheet'))
  const at = parseAddress(checkAddress(to.cell, 'op.to.cell'))
  const x = pivotArea(p)
  const dest = { r1: at.row, c1: at.col, r2: at.row + x.r2 - x.r1, c2: at.col + x.c2 - x.c1 }
  const over = placeRefusal(book, t, dest, p.name, t === s ? p : undefined)
  if (over) return over
  // the cells move as cut and paste moves them, the PivotTable with them, to its sheet's list
  const r = moveRange(book, book.sheets[s].name, x, book.sheets[t].name, at.row, at.col)
  if (!r.ok) return r
  const moved = movedPivot(p, dest.r1 - x.r1, dest.c1 - x.c1)
  let out = r.book
  if (s === t) out = withPivot(out, s, out.sheets[s].pivotTables!.findIndex((q) => q.name === p.name), moved)
  else {
    out = withPivot(out, s, out.sheets[s].pivotTables!.findIndex((q) => q.name === p.name), null)
    const sheets = [...out.sheets]
    sheets[t] = withField(sheets[t], 'pivotTables', [...(sheets[t].pivotTables ?? []), moved])
    out = renamedRefs({ ...out, sheets }, { sheet: book.sheets[s].name, name: p.name }, { sheet: book.sheets[t].name, name: p.name })
  }
  return relaid(out, [{ sheet: book.sheets[t].name, name: p.name }], values, false, r.geometry, 'refuse')
}

function deletePivotTable(book: BookInput, o: Record<string, unknown>): Applied {
  const { s, p } = pivotAt(book, o.sheet, o.pivot, 'op')
  const x = pivotArea(p)
  const cells = { ...book.sheets[s].cells }
  for (const a of Object.keys(cells)) {
    const { row, col } = parseAddress(a)
    if (row >= x.r1 && row <= x.r2 && col >= x.c1 && col <= x.c2) delete cells[a]
  }
  const sheets = [...book.sheets]
  sheets[s] = { ...sheets[s], cells }
  return { ok: true, book: withoutPivot({ ...book, sheets }, { sheet: book.sheets[s].name, name: p.name }), geometry: [{ kind: 'clear', s, area: x }] }
}

function refreshPivots(book: BookInput, o: Record<string, unknown>, values: ValuesIn): Applied {
  const refs = o.pivots === undefined ? undefined : pivotRefs(o.pivots, 'op.pivots')
  const missing = refs?.find((r) => !book.sheets[sheetIndex(book, r.sheet)]?.pivotTables?.some((p) => caseKey(p.name) === caseKey(r.name)))
  if (missing) return refuse(book, { key: 'pivotUnknown', sheet: missing.sheet, pivot: missing.name })
  const refresh = pivotRefresh()
  return refresh ? refresh(book, refs, values, values.today?.() ?? 0, true) : refuse(book, { key: 'pivotNotLoaded' })
}

export function pivotOp(book: BookInput, op: PivotOp, values: ValuesIn): Applied {
  const o = op as unknown as Record<string, unknown>
  switch (op.op) {
    case 'addPivotTable':
      return addPivotTable(book, o, values)
    case 'editPivotTable':
      return editPivotTable(book, o, values)
    case 'renamePivotTable':
      return renamePivotTable(book, o)
    case 'movePivotTable':
      return movePivotTable(book, o, values)
    case 'deletePivotTable':
      return deletePivotTable(book, o)
    case 'refreshPivots':
      return refreshPivots(book, o, values)
    default:
      throw new InputError('op.op', `${op.op} is not supported yet`)
  }
}

