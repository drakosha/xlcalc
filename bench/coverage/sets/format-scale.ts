/**
 * Excel cases for TEXT rounding a number that a percent sign or a thousands scale multiplies: the product is
 * rounded as Excel shows it, to 15 significant digits, before the format's own rounding (`TEXT(0.285,"0%")`).
 * Computed as dynamic array formulas (Excel 365), in Excel en-US (`test/excel/cases/format-scale.json`).
 */
import type { Probe } from '../probes.js'

export const DYNAMIC = true

const CASES: [value: string, format: string][] = [
  ['0.285', '0%'], ['0.145', '0%'], ['0.0285', '0.0%'], ['1.0005', '0.0%'], ['0.575', '0%'], ['0.285', '0.%'], ['1234500', '#,##0,'], ['2500', '0,'],
  ['1234567500', '0.0,,'], ['0.565', '0%'], ['1.005', '0.00'], ['2.675', '0.00'], ['-0.285', '0%'], ['0.29', '0%'],
]

export function buildProbes(): Probe[] {
  return CASES.map(([v, f], i) => ({ id: `TEXT#${i + 1}`, group: 'TEXT', fn: 'TEXT', formula: `TEXT(${v},"${f}")` }))
}
