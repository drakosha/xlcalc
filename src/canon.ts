/**
 * Канонический текст значений и хэши: построчно по листу и адресу (`"Лист"!E12 n 1234.5`, `s "…"`,
 * `b TRUE`, `e DIV0`), `-0` это 0, нефинитное `#NUM!`, числа округлены до 15 значащих цифр. Версия
 * конфигурации движка входит в хэш значений.
 */
import { createHash } from 'node:crypto'
import type { BookInput, CellValue } from './contract'
import { compareAddress } from './model'

/** Число для выдачи и хэша: 15 значащих цифр, без `-0`. */
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

/** Значения: лист → адрес → значение; пустые не входят. */
export function canonText(values: Map<string, Map<string, CellValue>>): string {
  const lines: string[] = []
  for (const sheet of [...values.keys()].sort(byText)) {
    const cells = values.get(sheet)!
    for (const addr of [...cells.keys()].sort(compareAddress)) {
      const v = cells.get(addr)!
      if (v !== null) lines.push(`${JSON.stringify(sheet)}!${addr} ${canonValue(v)}`)
    }
  }
  return lines.join('\n')
}

export function valuesHash(values: Map<string, Map<string, CellValue>>, configVersion: string): string {
  return createHash('sha256').update(configVersion).update('\n').update(canonText(values)).digest('hex')
}

/** JSON с ключами объектов по порядку: порядок полей во входе не влияет на хэш. */
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

/** Хэш входа: листы по имени, ячейки по адресу, имена по (scope, name). */
export function inputHash(book: BookInput): string {
  const h = createHash('sha256')
  for (const s of [...book.sheets].sort((a, b) => byText(a.name, b.name))) {
    h.update(`S ${JSON.stringify(s.name)}\n`)
    for (const addr of Object.keys(s.cells).sort(compareAddress)) h.update(`${addr} ${stableJson(s.cells[addr])}\n`)
  }
  const names = [...(book.names ?? [])].sort((a, b) => byText(`${a.scope ?? ''}\u0000${a.name}`, `${b.scope ?? ''}\u0000${b.name}`))
  for (const n of names) h.update(`N ${stableJson([n.scope ?? null, n.name, n.ref])}\n`)
  return h.digest('hex')
}
