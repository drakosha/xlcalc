/**
 * A workbook's PivotTables as a file Excel opens can hold them, as the sets `bench/coverage/links/file-refusals*.ts` measured:
 * Excel refuses the whole file over any of these, so the writer changes them and says so. A cache with two fields of one name or a
 * field of no name is left out with its PivotTables, slicers and timelines, their cells kept; a name or a caption over 255 characters
 * is cut, a name cut into another of its sheet taking a number at its end; a source sheet's name over 31 is cut; errors
 * Excel does not store in a cache are written `#VALUE!`.
 */
import type { BookInput, PivotCache, PivotTable, PivotValue, SheetInput } from '../../contract.js'
import { caseKey } from '../../collate.js'
import type { Issues } from '../issues.js'
import { clip } from './xml.js'

const MAX_TEXT = 255
const MAX_SHEET = 31
const STORED: ReadonlySet<string> = new Set(['DIV0', 'NA', 'NAME', 'NUM', 'REF', 'VALUE', 'NULL'])

const badFields = (c: PivotCache) => {
  const seen = new Set<string>()
  return c.fields.some((f) => {
    const k = caseKey(f.name)
    const bad = f.name === '' || seen.has(k)
    seen.add(k)
    return bad
  })
}

const stored = (v: PivotValue): boolean => v === null || typeof v !== 'object' || !('error' in v) || STORED.has(v.error)

function fitCache(c: PivotCache, issues: Issues): PivotCache {
  const value = (v: PivotValue): PivotValue => (stored(v) ? v : { error: 'VALUE' })
  const clean = c.fields.every((f) => !f.items || f.items.every(stored)) && (c.records ?? []).every((r) => r.every(stored))
  if (!clean) issues.add('errorValue', { key: 'pivotErrorValue' })
  const source = c.source.type === 'range' ? { ...c.source, sheet: clip(c.source.sheet, MAX_SHEET, issues, { key: 'pivotSourceCut' }, {}) } : c.source
  return clean ? { ...c, source } : { ...c, source, fields: c.fields.map((f) => (f.items ? { ...f, items: f.items.map(value) } : f)), records: c.records?.map((r) => r.map(value)) }
}

/** A sheet's PivotTables with their names and captions cut, and their new names by the old. */
function fitTables(s: SheetInput, issues: Issues): { tables?: PivotTable[]; renamed: Map<string, string> } {
  const renamed = new Map<string, string>()
  const taken = new Set((s.pivotTables ?? []).filter((p) => p.name.length <= MAX_TEXT).map((p) => caseKey(p.name)))
  const where = { sheet: s.name }
  const cut = (t: string | undefined) => (t === undefined ? t : clip(t, MAX_TEXT, issues, { key: 'pivotTextCut' }, where))
  const tables = s.pivotTables?.map((p) => {
    let name = p.name
    if (name.length > MAX_TEXT) {
      name = cut(name)!
      for (let n = 2; taken.has(caseKey(name)); n++) name = `${name.slice(0, MAX_TEXT - String(n).length)}${n}`
      taken.add(caseKey(name))
      renamed.set(p.name, name)
    }
    const cap = p.captions
    return {
      ...p,
      name,
      ...(cap ? { captions: { ...cap, ...(cap.values === undefined ? {} : { values: cut(cap.values) }), ...(cap.grandTotal === undefined ? {} : { grandTotal: cut(cap.grandTotal) }) } } : {}),
      ...(p.errors?.text === undefined ? {} : { errors: { ...p.errors, text: cut(p.errors.text) } }),
      ...(p.empty?.text === undefined ? {} : { empty: { ...p.empty, text: cut(p.empty.text) } }),
    }
  })
  return { tables, renamed }
}

/** The workbook with its PivotTables as a file holds them; the workbook itself where it has none. */
export function fitPivots(book: BookInput, issues: Issues): BookInput {
  if (!book.pivotCaches?.length) return book
  const out = book.pivotCaches.filter(badFields).map((c) => c.id)
  if (out.length) issues.add('pivotTable', { key: 'pivotFieldsOut' })
  const gone = new Set(out)
  const lost = new Set((book.slicerCaches ?? []).filter((c) => c.cache !== undefined && gone.has(c.cache)).map((c) => caseKey(c.name)))
  const fits = new Map(book.sheets.map((s) => [s.name, fitTables({ ...s, pivotTables: s.pivotTables?.filter((p) => !gone.has(p.cache)) }, issues)]))
  const rename = <C extends { pivots: { sheet: string; name: string }[] }>(c: C): C => ({ ...c, pivots: c.pivots.map((r) => ({ ...r, name: fits.get(r.sheet)?.renamed.get(r.name) ?? r.name })) })
  const lostTables = new Map(book.sheets.map((s) => [s.name, new Set((s.pivotTables ?? []).filter((p) => gone.has(p.cache)).map((p) => caseKey(p.name)))]))
  return {
    ...book,
    pivotCaches: book.pivotCaches.filter((c) => !gone.has(c.id)).map((c) => fitCache(c, issues)),
    ...(book.slicerCaches ? { slicerCaches: book.slicerCaches.filter((c) => c.cache === undefined || !gone.has(c.cache)).map(rename) } : {}),
    ...(book.timelineCaches ? { timelineCaches: book.timelineCaches.filter((c) => !gone.has(c.cache)).map(rename) } : {}),
    sheets: book.sheets.map((s) => {
      const fit = fits.get(s.name)!
      const dropped = lostTables.get(s.name)!
      return {
        ...s,
        ...(fit.tables ? { pivotTables: fit.tables } : {}),
        ...(s.slicers ? { slicers: s.slicers.filter((x) => !lost.has(caseKey(x.cache))) } : {}),
        ...(s.timelines && book.timelineCaches ? { timelines: s.timelines.filter((t) => book.timelineCaches!.some((c) => caseKey(c.name) === caseKey(t.cache) && !gone.has(c.cache))) } : {}),
        ...(s.conditionalFormats ? { conditionalFormats: s.conditionalFormats.filter((r) => !r.pivot || !dropped.has(caseKey(r.pivot.pivot))).map((r) => (r.pivot && fit.renamed.has(r.pivot.pivot) ? { ...r, pivot: { ...r.pivot, pivot: fit.renamed.get(r.pivot.pivot)! } } : r)) } : {}),
      }
    }),
  }
}
