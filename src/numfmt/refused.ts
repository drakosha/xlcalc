/**
 * Number format codes Excel does not open a file with, as `numFmts` is read on opening: measured in Excel for Windows
 * (`bench/coverage/display/digits-date-letters.ts`, `bare-literals.ts`). Literals need no quotes while no letter reads as a
 * code: `0 шт`, `#,##0 ₽` and `0 x` open, `0 m`, `0 kg` and `0 USD` do not.
 */
import { EN_CODES, parseFormat } from './parse.js'

/** Excel's limit of a code's length. */
const MAX_CODE = 255
/** The letters of dates, times and eras, in either case: beside a digit placeholder they refuse a section. */
const DATE_LETTERS = /[bdeghmnsy]/i

/**
 * A section with its quoted and escaped text, paddings and fills out: what Excel reads as codes; an elapsed time's bracket is
 * left as its letter, the others go. Null when something does not close or means nothing.
 */
function bareOf(section: string): string | null {
  let out = ''
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]
    if (ch === '"') {
      i = section.indexOf('"', i + 1)
      if (i < 0) return null
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      if (++i >= section.length) return null
    } else if (ch === '[') {
      const j = section.indexOf(']', i)
      if (j < 0 || parseFormat(section.slice(i, j + 1), EN_CODES) === null) return null
      if (/^\[[hms]+\]$/i.test(section.slice(i, j + 1))) out += 'h'
      i = j
    } else out += ch
  }
  return out
}

/** Whether Excel reads a section's codes, by its place among the code's sections. */
function sectionReads(section: string, last: boolean): boolean {
  const found = bareOf(section)
  if (found === null) return false
  const exp = /E[+-]/.exec(found)
  if (exp && !(/[0#?]/.test(found.slice(0, exp.index)) && /[0#?]/.test(found.slice(exp.index + 2)))) return false
  // a word of the code's own is no date's letter, and an exponent no era
  const bare = found.replace(/General/gi, ' ').replace(/AM\/PM|A\/P/gi, 'h').replace(/E[+-]/g, ' ')
  const date = bare.search(DATE_LETTERS)
  // zeros after a point past a date's first code are the second's decimals (`ss.00`)
  const rest = date < 0 ? bare : `${bare.slice(0, date)}${bare.slice(date).replace(/\.0{1,3}(?!\d)/g, ' ')}`
  const digits = /[0#?]/.test(rest)
  if (date >= 0 && digits) return false
  if (bare.includes('@') && (digits || date >= 0 || !last)) return false
  if (date < 0 && rest.includes('/')) {
    const fraction = rest.split('/')
    if (fraction.length > 2 || !/[0-9#?]\s*$/.test(fraction[0]) || !/^\s*[0-9#?]/.test(fraction[1] ?? '')) return false
  }
  return true
}

/** Whether Excel refuses a file with this number format code. */
export function formatRefused(code: string): boolean {
  if (code.length > MAX_CODE) return true
  const sections: string[] = []
  let from = 0
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]
    if (ch === '"') {
      const j = code.indexOf('"', i + 1)
      if (j < 0) return true
      i = j
    } else if (ch === '\\' || ch === '_' || ch === '*') i++
    else if (ch === '[') {
      const j = code.indexOf(']', i)
      if (j < 0) return true
      i = j
    } else if (ch === ';') {
      sections.push(code.slice(from, i))
      from = i + 1
    }
  }
  sections.push(code.slice(from))
  return sections.length > 4 || !sections.every((s, i) => sectionReads(s, i === sections.length - 1))
}
