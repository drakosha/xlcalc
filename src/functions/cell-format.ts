/**
 * CELL's codes of a number format (`CELL("format")`, `"color"`, `"parentheses"`), by Microsoft's table of them and the
 * same rules for the formats it does not list: `G` General, text or a fraction; `F`, `,` (thousands), `C` (currency),
 * `P` (percent) and `S` (scientific) with the decimals of the positive section; `D1` to `D9` the dates and times; `-`
 * after it when negatives show in a color, `()` when positives show in parentheses.
 */

/** The sections of a format code (`;` outside quotes and brackets). */
function sections(code: string): string[] {
  const out: string[] = []
  let cur = ''
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]
    if (ch === '"') {
      const end = code.indexOf('"', i + 1)
      const to = end < 0 ? code.length : end
      cur += code.slice(i, to + 1)
      i = to
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      cur += code.slice(i, i + 2)
      i++
    } else if (ch === '[') {
      const end = code.indexOf(']', i)
      const to = end < 0 ? code.length : end
      cur += code.slice(i, to + 1)
      i = to
    } else if (ch === ';') {
      out.push(cur)
      cur = ''
    } else cur += ch
  }
  out.push(cur)
  return out
}

/** A section without its literal text, padding, fills and brackets, a currency in brackets (`[$€-407]`) kept as `$`. */
const plain = (s: string) =>
  s
    .replace(/"[^"]*"?/g, '')
    .replace(/\\./g, '')
    .replace(/[_*]./g, '')
    .replace(/\[\$[^\]]*\]/g, '$')
    .replace(/\[[^\]]*\]/g, '')

const COLORED = /\[(black|blue|cyan|green|magenta|red|white|yellow|color\s*\d+)\]/i

/** The date or time code of a section's date and time tokens, or undefined for none. */
function dateCode(p: string): string | undefined {
  const t = p.toLowerCase()
  const ampm = /am\/pm|a\/p/.test(t)
  const s = t.replace(/am\/pm|a\/p/g, '')
  const [d, m, y, h, sec] = ['d', 'm', 'y', 'h', 's'].map((x) => s.includes(x))
  if (!d && !m && !y && !h && !sec) return undefined
  // a time: h with minutes or seconds; m after h or before s is minutes
  if (h && !d && !y) return sec ? (ampm ? 'D6' : 'D8') : ampm ? 'D7' : 'D9'
  const textMonth = /mmm/.test(s)
  if (y && d) return textMonth ? 'D1' : 'D4'
  if (d && m) return textMonth ? 'D2' : 'D5'
  if (m && y) return 'D3'
  return 'D4'
}

/** CELL's `format`, `color` and `parentheses` of a format code (en-US syntax). */
export function formatCodes(code: string): { format: string; color: number; parentheses: number } {
  const secs = sections(code)
  const color = secs.length > 1 && COLORED.test(secs[1]) ? 1 : 0
  const p = plain(secs[0])
  const parentheses = p.includes('(') ? 1 : 0
  const suffix = (color ? '-' : '') + (parentheses ? '()' : '')
  const general = /^\s*general\s*$/i.test(p) || /^[\s@]*$/.test(p) || /\?\s*\/|#\s*\/|0\s*\//.test(p)
  if (general) return { format: 'G' + suffix, color, parentheses }
  const date = dateCode(p.replace(/e[+-]/gi, ''))
  if (date && !/[0#?]/.test(p)) return { format: date + suffix, color, parentheses }
  const dot = p.indexOf('.')
  const decimals = dot < 0 ? 0 : (p.slice(dot + 1).match(/^[0#?]*/)?.[0].length ?? 0)
  const kind = p.includes('%') ? 'P' : /e[+-]/i.test(p) ? 'S' : p.includes('$') ? 'C' : /[0#?],[0#?]/.test(p) ? ',' : 'F'
  return { format: kind + decimals + suffix, color, parentheses }
}
