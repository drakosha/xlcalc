/**
 * Canonical text of values and the hashes: one line per cell, ordered by sheet and address
 * (`"Sheet"!E12 n 1234.5`, `s "…"`, `b TRUE`, `e DIV0`); `-0` is 0, non-finite is `#NUM!`, numbers are
 * rounded to 15 significant digits; then one line per range of a large spilled array's unstored cells
 * (`"Sheet"!B4:B1048576 n 0`). The engine configuration version is part of the values hash, and so is a
 * book's locale other than en-US.
 */
import type { BookInput, CellValue } from './contract.js'
import { lines, spanKey } from './layout/spans.js'
import { DEFAULT_LOCALE } from './locale/registry.js'
import { compareAddress, sortAddresses } from './model.js'
import type { Axis } from './structure/area.js'
import { Sha256 } from './sha256.js'

/** A number for output and hashing: 15 significant digits, no `-0`. */
export function outNumber(v: number): number {
  if (v === 0) return 0
  return Number(v.toPrecision(15))
}

export function canonValue(v: CellValue): string {
  if (v === null) return '_'
  if (typeof v === 'number') return Number.isFinite(v) ? `n ${String(outNumber(v))}` : 'e NUM'
  if (typeof v === 'string') return `s ${JSON.stringify(v)}`
  if (typeof v === 'boolean') return v ? 'b TRUE' : 'b FALSE'
  return `e ${v.error}`
}

const byText = (a: string, b: string) => (a < b ? -1 : a > b ? 1 : 0)

/** A range of a sheet whose cells all hold one value: the unstored cells of a large spilled array. */
export interface FillLine {
  sheet: string
  range: string
  value: CellValue
}

/** Values: sheet → address → value; empty cells are skipped. Then the fill ranges, by sheet and their top-left cell. */
export function canonText(values: Map<string, Map<string, CellValue>>, fills: readonly FillLine[] = []): string {
  const lines: string[] = []
  for (const sheet of [...values.keys()].sort(byText)) {
    const cells = values.get(sheet)!
    for (const addr of sortAddresses(cells.keys())) {
      const v = cells.get(addr)!
      if (v !== null) lines.push(`${JSON.stringify(sheet)}!${addr} ${canonValue(v)}`)
    }
  }
  const top = (f: FillLine) => f.range.split(':')[0]
  for (const f of [...fills].sort((a, b) => byText(a.sheet, b.sheet) || compareAddress(top(a), top(b)))) lines.push(`${JSON.stringify(f.sheet)}!${f.range} ${canonValue(f.value)}`)
  return lines.join('\n')
}

/** A locale that takes part in the hashes: one given and not en-US, so books computed as before hash as before. */
const hashedLocale = (locale: string | undefined): string | undefined => (locale === DEFAULT_LOCALE ? undefined : locale)

/** The values hash: the configuration version, `;locale=<tag>` for a book in a locale other than en-US, the values and fills. */
export function valuesHash(values: Map<string, Map<string, CellValue>>, configVersion: string, fills: readonly FillLine[] = [], locale?: string): string {
  const tag = hashedLocale(locale)
  return new Sha256().update(configVersion).update(tag === undefined ? '' : `;locale=${tag}`).update('\n').update(canonText(values, fills)).hex()
}

/** JSON with sorted object keys: the order of fields in the input does not affect the hash. */
function stableJson(v: unknown): string {
  if (v === null || typeof v !== 'object') return JSON.stringify(v) ?? 'null'
  if (Array.isArray(v)) return `[${v.map(stableJson).join(',')}]`
  const o = v as Record<string, unknown>
  return `{${Object.keys(o)
    .filter((k) => o[k] !== undefined)
    .sort(byText)
    .map((k) => `${JSON.stringify(k)}:${stableJson(o[k])}`)
    .join(',')}}`
}

/** Lines of a sheet's columns or rows in order (`C 2 {...}`, `R 5 {...}`); none without them. */
function lineEntries(tag: string, map: Record<string, unknown> | undefined, axis: Axis): string[] {
  return lines(map, axis).map(([lo, hi, v]) => `${tag} ${spanKey(lo, hi, axis)} ${stableJson(v)}\n`)
}

/**
 * Input hash: sheets by name, cells by address, names by (scope, name). Columns, rows, merged areas (in
 * order), the filter mode, the style table and a locale other than en-US come after their sheet and the names
 * when there are any, so a book without them hashes as before they existed.
 */
export function inputHash(book: BookInput): string {
  const h = new Sha256()
  for (const s of [...book.sheets].sort((a, b) => byText(a.name, b.name))) {
    h.update(`S ${JSON.stringify(s.name)}\n`)
    for (const addr of sortAddresses(Object.keys(s.cells))) h.update(`${addr} ${stableJson(s.cells[addr])}\n`)
    for (const line of [...lineEntries('C', s.columns, 'c'), ...lineEntries('R', s.rows, 'r')]) h.update(line)
    for (const m of [...(s.merges ?? [])].sort(compareMerges)) h.update(`M ${m}\n`)
    if (s.filtered) h.update('F\n')
  }
  const names = [...(book.names ?? [])].sort((a, b) => byText(`${a.scope ?? ''}\u0000${a.name}`, `${b.scope ?? ''}\u0000${b.name}`))
  for (const n of names) h.update(`N ${stableJson([n.scope ?? null, n.name, n.ref])}\n`)
  if (book.styles?.length) h.update(`Y ${stableJson(book.styles)}\n`)
  const tag = hashedLocale(book.locale)
  if (tag !== undefined) h.update(`L ${JSON.stringify(tag)}\n`)
  return h.hex()
}

/** Merged areas by their top-left cell. */
const compareMerges = (a: string, b: string) => compareAddress(a.split(':')[0], b.split(':')[0])
