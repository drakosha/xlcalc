/** New item counts and equal-value groups for the held-out sort cases. */
import type { SourceColumn } from './data.js'

export const ORDER_COUNTS = [2, 3, 7, 8, 9, 16, 17, 31, 32, 33, 127, 511, 512, 513, 1024]

function columns(n: number, mixed: boolean): SourceColumn[] {
  const rows = Array.from({ length: n }, (_, k) => (k + Math.floor(n / 3)) % n).flatMap((i) => [
    [`Item ${String(i).padStart(4, '0')}`, 'First', mixed ? (i * 17) % 23 - 11 : n % 2 ? i % 7 : 4],
    [`Item ${String(i).padStart(4, '0')}`, 'Second', mixed ? (i * 7) % 13 : n % 2 ? 7 - i % 7 : 6],
  ])
  return ['Region', 'Product', 'Qty'].map((name, j) => ({ name, values: rows.map((r) => ({ value: r[j] })) }))
}

export const ORDER_DATA: Record<string, () => SourceColumn[]> = Object.fromEntries([
  ...ORDER_COUNTS.map((n) => [`order-${n}`, () => columns(n, false)]),
  ...[600, 777].map((n) => [`order-mixed-${n}`, () => columns(n, true)]),
])
