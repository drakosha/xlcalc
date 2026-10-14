/** Builders of the sets of TEXT: a value shown in a format code. */
import type { Probe } from '../probes.js'

/** A format code as a formula's text literal. */
const lit = (s: string) => `"${s.replace(/"/g, '""')}"`

/** `TEXT(value,"format")` for each `[value, format]`, ids `TEXT#<n>` (a code written as it is, its quotes doubled by the set). */
export const textProbes = (cases: readonly (readonly [string, string])[]): Probe[] => cases.map(([v, f], i) => ({ id: `TEXT#${i + 1}`, group: 'TEXT', fn: 'TEXT', formula: `TEXT(${v},"${f}")` }))

/** Each number code over each number (group `number`) and each date code over each date (group `date`), as TEXT. */
export function textFormatProbes(numberCodes: readonly string[], numbers: readonly string[], dateCodes: readonly string[], dates: readonly string[]): Probe[] {
  const out: Probe[] = []
  const add = (group: string, value: string, code: string) => out.push({ id: `${group}#${out.filter((p) => p.group === group).length + 1}`, group, fn: 'TEXT', formula: `TEXT(${value},${lit(code)})` })
  for (const c of numberCodes) for (const v of numbers) add('number', v, c)
  for (const c of dateCodes) for (const v of dates) add('date', v, c)
  return out
}
