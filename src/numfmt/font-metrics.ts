import type { BookInput, CellStyle, FontStyle } from '../contract.js'
import { columnPadding, fontFamily } from '../column-width.js'
import { LIBERATION_SANS, LIBERATION_SERIF } from './liberation.js'

/** Measurements for a supported face and point size. A letter outside the table takes a digit's width. */
export interface FontMeasure {
  size: number
  digit: number
  padding: number
  rowHeight: number
  advance(code: number): number
  width(code: number): number
}

function advancesOf(groups: readonly (readonly [number, string])[]): Map<number, number> {
  const out = new Map<number, number>()
  for (const [width, letters] of groups) for (const letter of letters) out.set(letter.codePointAt(0)!, width)
  return out
}

const faces: Record<string, Map<number, number>> = {}
const measures = new Map<string, FontMeasure>()

/** The regular metrics for the cell's style over the Normal font, for the sizes measured in Excel. */
export function measureFont(normal: BookInput['font'], font?: FontStyle): FontMeasure | undefined {
  const family = fontFamily(font?.name || normal?.name)
  const size = font?.size ?? normal?.size ?? 11
  if (family === undefined || !Number.isInteger(size) || size < 8 || size > 14) return undefined

  const key = `${family}:${size}`
  const known = measures.get(key)
  if (known) return known

  const steps = faces[family] ??= advancesOf(family === 'sans' ? LIBERATION_SANS : LIBERATION_SERIF)
  const zero = steps.get(48)!
  const advance = (code: number) => steps.get(code) ?? zero
  const width = (code: number) => Math.round(advance(code) * size / 2048)
  const digit = width(48)
  const rowHeight = [11, 12, 13, 14, 16, 17, 18][size - 8]
  const metric = { size, digit, padding: columnPadding(digit), rowHeight, advance, width }
  measures.set(key, metric)
  return metric
}

/** A style's measurements by its index. Undefined when every cell uses the legacy Calibri metrics. */
export function styleMeasures(normal: BookInput['font'], styles: readonly CellStyle[] | undefined): ((style?: number) => FontMeasure | undefined) | undefined {
  const base = measureFont(normal)
  const measured = styles?.map((style) => measureFont(normal, style.font))
  if (!base && !measured?.some(Boolean)) return undefined
  return (style) => style === undefined ? base : measured?.[style]
}
