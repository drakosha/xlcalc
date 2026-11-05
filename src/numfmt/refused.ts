/**
 * Number format codes Excel does not open a file with, as `numFmts` is read on opening, a code no cell uses too: measured in
 * Excel for Windows in en-US and ru-RU and in Excel for Mac (`bench/coverage/display/`: `digits-date-letters`, `bare-literals`,
 * `bare-held`, `refused-held`, `written-held`). Literals need no quotes while no letter reads as a code: `0 шт`, `#,##0 ₽` and `0 x` open;
 * `0 m`, `0 kg`, `0 USD` and `0 шт/ч` do not, `0 "kg"` and `0 "шт/ч"` do.
 */
import type { NumberFormat } from '../contract.js'
import { checkFormat } from '../input/styles.js'
import { EN_CODES, parseFormat, splitSections } from './parse.js'

/** A longer code Excel cuts to this or refuses: never what was asked for. */
const MAX_CODE = 255
/** The letters of dates, times and eras, in either case: beside a digit placeholder they refuse a section. */
const DATE_LETTERS = /[bdeghmnsy]/i
/** Excel's numerals of East Asian scripts, which the parser does not read. */
const DBNUM = /^\[DBNum[1-4]\]$/i
/** Conditions a code may have (`[<1]0;[<2]0;0`). */
const MAX_CONDITIONS = 2
/**
 * What a whole code compiles to at most, in Excel's units, with a margin: Excel opens up to 252 of a lone run of text, and the
 * cost of other tokens is an estimate. A code within the margin is refused though Excel may open it: a code that long is rare, a
 * file Excel does not open is not worth the risk.
 */
const MAX_COST = 240
/** A bare literal, each; quoted text, 2 and 2 for each UTF-16 unit; a bracket; digit placeholders in a run; another code's character. */
const CHAR_COST = 3
const QUOTED_COST = 2
const BRACKET_COST = 14
const DIGITS_COST = 5
const CODE_COST = 2

interface Read {
  /** What Excel reads as codes: quoted and escaped text, paddings and fills out; an elapsed time's bracket left as its letter, the others gone. */
  bare: string
  conditions: number
  /** Colors and conditions: Excel takes one of each to a section. */
  twice: boolean
  /** What quoted and escaped text and brackets cost of the section's size. */
  cost: number
  /** Whether it ends in quoted text. */
  quotedLast: boolean
}

/** A section as Excel reads it; null when something does not close or means nothing. */
function readSection(section: string): Read | null {
  const out: Read = { bare: '', conditions: 0, twice: false, cost: 0, quotedLast: false }
  let colors = 0
  let conditions = 0
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]
    if (ch === '"') {
      const j = section.indexOf('"', i + 1)
      if (j < 0) return null
      out.cost += QUOTED_COST + QUOTED_COST * (j - i - 1)
      i = j
      out.quotedLast = i === section.length - 1
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      if (++i >= section.length) return null
      out.cost += CHAR_COST
    } else if (ch === '[') {
      const j = section.indexOf(']', i)
      const bracket = section.slice(i, j + 1)
      if (j < 0 || (parseFormat(bracket, EN_CODES) === null && !DBNUM.test(bracket))) return null
      if (/^\[[hms]+\]$/i.test(bracket)) out.bare += 'h'
      if (/^\[[<>=]/.test(bracket)) conditions++
      else if (/^\[[a-z]/i.test(bracket) && !/^\[[hms]+\]$/i.test(bracket)) colors++
      out.cost += BRACKET_COST
      i = j
    } else out.bare += ch
  }
  out.conditions = conditions
  out.twice = colors > 1 || conditions > 1
  return out
}

/** Whether Excel reads the codes of a section, by its place among the code's sections. */
function sectionReads(found: string, last: boolean): boolean {
  // lower-case, this is a year
  if (/^E+$/.test(found)) return false
  const exp = /E[+-]/.exec(found)
  if (exp && !(/[0#?]/.test(found.slice(0, exp.index)) && /[0#?]/.test(found.slice(exp.index + 2)))) return false
  const general = /General/i.test(found)
  const ampm = /AM\/PM|A\/P/i.exec(found)
  if (ampm && /[0#?]/.test(found.slice(ampm.index + ampm[0].length))) return false
  // a word of the code's own is no date's letter, and an exponent no era; `AM/PM` shows beside digits too (`0 AM/PM`)
  const bare = found.replace(/General/gi, ' ').replace(/AM\/PM|A\/P/gi, ' ').replace(/E[+-]/g, ' ')
  const date = bare.search(DATE_LETTERS)
  // zeros after a point past a date's first code are the second's decimals (`ss.00`)
  const rest = date < 0 ? bare : `${bare.slice(0, date)}${bare.slice(date).replace(/\.0{1,3}(?!\d)/g, ' ')}`
  const digits = /[0#?]/.test(rest)
  if (date >= 0 && digits) return false
  // a time's separator: a date's own, beside digits it refuses a section as a date's letter does
  if (digits && rest.includes(':')) return false
  if (general && (digits || bare.includes('@'))) return false
  if (bare.includes('@') && (digits || date >= 0 || !last)) return false
  if (date < 0 && rest.includes('/')) {
    const fraction = rest.split('/')
    if (fraction.length > 2 || !/[0-9#?]\s*$/.test(fraction[0]) || !/^\s*[0-9#?]/.test(fraction[1] ?? '')) return false
  }
  return true
}

/** A section's size in Excel's units: its text by `readSection`, a bare literal or a code's character of the rest. */
function sectionCost(read: Read): number {
  const rest = read.bare.replace(/General/gi, ' ')
  const codes = rest.replace(/[^0#?.,%@/]/g, '').replace(/[0#?]/g, '').length
  const literals = rest.replace(/[0#?.,%@/]/g, '').length
  return read.cost + literals * CHAR_COST + codes * CODE_COST + (/[0#?]/.test(rest) ? DIGITS_COST : 0) + (literals && !read.quotedLast ? 1 : 0)
}

/**
 * Whether Excel does not open a file with this number format, which `writeXlsx` and `writeKept` then refuse: for a Format Cells
 * window that says so at once. Text goes in quotes where its letters read as codes (`0 kg` is refused, `0 "kg"` is not). A
 * built-in format by its number is not refused. The check comes from measurements and is strict where it does not know.
 */
export function formatRefused(format: NumberFormat): boolean {
  checkFormat(format, 'format')
  if (typeof format === 'number') return false
  const code = format
  if (code.length > MAX_CODE) return true
  const sections = splitSections(code)
  if (sections.length > 4) return true
  const read = sections.map(readSection)
  if (read.some((r) => r === null)) return true
  const reads = read as Read[]
  const conditions = reads.reduce((n, r) => n + r.conditions, 0)
  if (conditions > MAX_CONDITIONS || (conditions && reads.length > 3) || reads.some((r) => r.twice)) return true
  if (reads.reduce((n, r) => n + sectionCost(r), 0) > MAX_COST) return true
  // a fourth section is for text alone (`0;-0;0;@`)
  if (reads.length === 4 && /[0#?]/.test(reads[3].bare)) return true
  return !reads.every((r, i) => sectionReads(r.bare, i === reads.length - 1))
}
