/**
 * Canonical text of values and the hashes: one line per cell, ordered by sheet and address
 * (`"Sheet"!E12 n 1234.5`, `s "…"`, `b TRUE`, `e DIV0`); `-0` is 0, non-finite is `#NUM!`, numbers are
 * rounded to 15 significant digits; then one line per range of a large spilled array's unstored cells
 * (`"Sheet"!B4:B1048576 n 0`). The engine configuration version is part of the values hash, and so are a
 * book's locale other than en-US and the 1904 date system.
 */
import type { BookInput, CellValue } from './contract.js'
import { paced, run, type Steps } from './exec/budget.js'
import { lines, spanKey } from './layout/spans.js'
import { DEFAULT_LOCALE } from './locale/registry.js'
import { addressOrder, compareAddress, sortAddresses } from './model.js'
import type { Axis } from './structure/area.js'
import { Sha256 } from './sha256.js'
import { to15 } from './values.js'

/** A number for output and hashing: 15 significant digits, no `-0`. */
export function outNumber(v: number): number {
  if (v === 0) return 0
  return to15(v)
}

export function canonValue(v: CellValue): string {
  if (v === null) return '_'
  if (typeof v === 'number') return Number.isFinite(v) ? `n ${String(outNumber(v))}` : 'e NUM'
  if (typeof v === 'string') return `s ${JSON.stringify(v)}`
  if (typeof v === 'boolean') return v ? 'b TRUE' : 'b FALSE'
  return `e ${v.error}`
}

const byText = (a: string, b: string) => (a < b ? -1 : a > b ? 1 : 0)

/** `sortAddresses` in steps: reading the addresses yields when the computation's slice is used up; the sort itself is one step. */
export function* sortedSteps(addrs: string[]): Steps<string[]> {
  const keyed: (readonly [number, string])[] = []
  for (const a of addrs) {
    if (paced(1)) yield
    keyed.push([addressOrder(a), a])
  }
  return keyed.sort((x, y) => x[0] - y[0]).map((x) => x[1])
}

/** A range of a sheet whose cells all hold one value: the unstored cells of a large spilled array. */
export interface FillLine {
  sheet: string
  range: string
  value: CellValue
}

/** Sheets in the order of the canonical text: by name, code unit by code unit. */
export const sheetOrder = (names: readonly string[]): number[] => names.map((_, i) => i).sort((a, b) => byText(names[a], names[b]))

/** A locale that takes part in the hashes: one given and not en-US, so books computed as before hash as before. */
const hashedLocale = (locale: string | undefined): string | undefined => (locale === DEFAULT_LOCALE ? undefined : locale)

/** Lines hashed at once. */
const PORTION = 1024

/** The values hash written as the lines come, without the canonical text or a map of the values: lines joined by `\n`, none after the last. */
export class ValuesHasher {
  private readonly h = new Sha256()
  /** Lines not hashed yet: a portion is hashed as one joined string, faster than a string per line. */
  private pending: string[] = []
  private started = false

  /** The header: the configuration version, `;locale=<tag>` for a locale other than en-US, `;date1904` for the 1904 system, `\n`. */
  constructor(configVersion: string, locale?: string, dateSystem?: 1900 | 1904) {
    const tag = hashedLocale(locale)
    this.h.update(`${configVersion}${tag === undefined ? '' : `;locale=${tag}`}${dateSystem === 1904 ? ';date1904' : ''}\n`)
  }

  /** A cell's line, none for an empty cell; cells come by sheet (`sheetOrder`), then address, each once. `prepared`: rounded by outNumber already. */
  cell(quoted: string, addr: string, v: CellValue, prepared: boolean): void {
    if (v !== null) this.line(`${quoted}!${addr} ${prepared && typeof v === 'number' && Number.isFinite(v) ? `n ${String(v)}` : canonValue(v)}`)
  }

  /** The digest after the fills' lines, by sheet, then top-left cell. The hasher must not be used afterwards. */
  hex(fills: readonly FillLine[]): string {
    const top = (f: FillLine) => f.range.split(':')[0]
    for (const f of [...fills].sort((a, b) => byText(a.sheet, b.sheet) || compareAddress(top(a), top(b)))) this.line(`${JSON.stringify(f.sheet)}!${f.range} ${canonValue(f.value)}`)
    this.flush()
    return this.h.hex()
  }

  private line(text: string): void {
    if (this.pending.push(text) === PORTION) this.flush()
  }

  /** Whole lines per update, so a character is never split between two. */
  private flush(): void {
    if (!this.pending.length) return
    if (this.started) this.h.update('\n')
    this.h.update(this.pending.join('\n'))
    this.pending = []
    this.started = true
  }
}

/** The values hash of values by sheet name and address; `prepared`: a result's, by address already and rounded by outNumber. */
export function valuesHash(values: Map<string, Map<string, CellValue>>, configVersion: string, fills: readonly FillLine[] = [], locale?: string, dateSystem?: 1900 | 1904, prepared = false): string {
  const w = new ValuesHasher(configVersion, locale, dateSystem)
  const names = [...values.keys()]
  for (const i of sheetOrder(names)) {
    const cells = values.get(names[i])!
    const quoted = JSON.stringify(names[i])
    for (const addr of prepared ? cells.keys() : sortAddresses(cells.keys())) w.cell(quoted, addr, cells.get(addr)!, prepared)
  }
  return w.hex(fills)
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
 * Input hash: sheets by name, cells by address (their notes and links with them), names by (scope, name). Columns,
 * rows, merged areas (in order), the filter mode, tables (in order), frozen panes, rules of data validation, outline
 * settings, the style table, a locale other than en-US, the 1904 date system and iterative calculation come after their
 * sheet and the names when there are any, so a book without them hashes as before.
 */
export const inputHash = (book: BookInput): string => run(inputHashSteps(book))

/** `inputHash` in steps: the cells yield when the computation's slice is used up. */
export function* inputHashSteps(book: BookInput): Steps<string> {
  const h = new Sha256()
  for (const s of [...book.sheets].sort((a, b) => byText(a.name, b.name))) {
    h.update(`S ${JSON.stringify(s.name)}\n`)
    for (const addr of yield* sortedSteps(Object.keys(s.cells))) {
      if (paced(1)) yield
      h.update(`${addr} ${stableJson(s.cells[addr])}\n`)
    }
    for (const line of [...lineEntries('C', s.columns, 'c'), ...lineEntries('R', s.rows, 'r')]) h.update(line)
    for (const m of [...(s.merges ?? [])].sort(compareMerges)) h.update(`M ${m}\n`)
    if (s.filtered) h.update('F\n')
    for (const t of s.tables ?? []) h.update(`T ${stableJson(t)}\n`)
    if (s.frozen) h.update(`P ${stableJson(s.frozen)}\n`)
    // rules of data validation in any order: no two meet
    for (const v of (s.validations ?? []).map(stableJson).sort(byText)) h.update(`V ${v}\n`)
    if (s.outline) h.update(`O ${stableJson(s.outline)}\n`)
  }
  const names = [...(book.names ?? [])].sort((a, b) => byText(`${a.scope ?? ''}\u0000${a.name}`, `${b.scope ?? ''}\u0000${b.name}`))
  for (const n of names) h.update(`N ${stableJson([n.scope ?? null, n.name, n.ref])}\n`)
  if (book.styles?.length) h.update(`Y ${stableJson(book.styles)}\n`)
  const tag = hashedLocale(book.locale)
  if (tag !== undefined) h.update(`L ${JSON.stringify(tag)}\n`)
  if (book.dateSystem === 1904) h.update('D 1904\n')
  if (book.iteration) h.update(`I ${stableJson(book.iteration)}\n`)
  return h.hex()
}

/** Merged areas by their top-left cell. */
const compareMerges = (a: string, b: string) => compareAddress(a.split(':')[0], b.split(':')[0])
