/** VALUETOTEXT and ARRAYTOTEXT, this engine's own (Excel 2021): format 0 concise, 1 strict (text quoted). */
import { gridAt, type Cx, type FnDef, type Grid } from '../../args.js'
import type { Rules } from '../../locale/rules.js'
import { ERROR_TEXT } from '../../syntax.js'
import { E, isErr, type V } from '../../values.js'
import { capped, MAX_TEXT } from './common.js'

const asText = (v: V, strict: boolean, rules: Rules): string => {
  if (isErr(v)) return ERROR_TEXT[v.code]
  if (typeof v === 'string') return strict ? `"${v.replace(/"/g, '""')}"` : v
  return rules.toText(v) as string
}

function valuetotext(cx: Cx, x: V, format = 0): V {
  if (format !== 0 && format !== 1) return E.VALUE
  return asText(x, format === 1, cx.rules)
}

function arraytotext(cx: Cx, g: Grid, format = 0): V {
  if (format !== 0 && format !== 1) return E.VALUE
  const strict = format === 1
  if (g.rows * g.cols > MAX_TEXT) return E.VALUE
  let out = strict ? '{' : ''
  for (let r = 0; r < g.rows; r++) {
    for (let c = 0; c < g.cols; c++) {
      if (r || c) out += strict ? (c ? ',' : ';') : ', '
      out += asText(gridAt(g, r * g.cols + c), strict, cx.rules)
    }
    if (out.length > MAX_TEXT) return E.VALUE
  }
  return capped(strict ? out + '}' : out)
}

export const functions: Record<string, FnDef> = {
  ARRAYTOTEXT: { args: 'G w?', cx: true, fn: arraytotext },
  VALUETOTEXT: { args: 'a w?', cx: true, fn: valuetotext },
}
