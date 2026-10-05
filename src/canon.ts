/**
 * Canonical text of values and the hashes: one line per cell, ordered by sheet and address
 * (`"Sheet"!E12 n 1234.5`, `s "…"`, `b TRUE`, `e DIV0`); `-0` is 0, non-finite is `#NUM!`, numbers are
 * rounded to 15 significant digits. The engine configuration version is part of the values hash.
 */
import type { BookInput, CellValue } from './contract.js'
import { sortAddresses } from './model.js'
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

/** Values: sheet → address → value; empty cells are skipped. */
export function canonText(values: Map<string, Map<string, CellValue>>): string {
  const lines: string[] = []
  for (const sheet of [...values.keys()].sort(byText)) {
    const cells = values.get(sheet)!
    for (const addr of sortAddresses(cells.keys())) {
      const v = cells.get(addr)!
      if (v !== null) lines.push(`${JSON.stringify(sheet)}!${addr} ${canonValue(v)}`)
    }
  }
  return lines.join('\n')
}

export function valuesHash(values: Map<string, Map<string, CellValue>>, configVersion: string): string {
  return new Sha256().update(configVersion).update('\n').update(canonText(values)).hex()
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

/** Input hash: sheets by name, cells by address, names by (scope, name). */
export function inputHash(book: BookInput): string {
  const h = new Sha256()
  for (const s of [...book.sheets].sort((a, b) => byText(a.name, b.name))) {
    h.update(`S ${JSON.stringify(s.name)}\n`)
    for (const addr of sortAddresses(Object.keys(s.cells))) h.update(`${addr} ${stableJson(s.cells[addr])}\n`)
  }
  const names = [...(book.names ?? [])].sort((a, b) => byText(`${a.scope ?? ''}\u0000${a.name}`, `${b.scope ?? ''}\u0000${b.name}`))
  for (const n of names) h.update(`N ${stableJson([n.scope ?? null, n.name, n.ref])}\n`)
  return h.hex()
}
