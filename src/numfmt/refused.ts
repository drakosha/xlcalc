/**
 * Number format codes Excel does not open a file with, as `numFmts` is read on opening, a code no cell uses too: measured in
 * Excel for Windows in en-US and ru-RU and in Excel for Mac (`bench/coverage/display/`: `digits-date-letters`, `bare-literals`,
 * `bare-held`, `refused-held`). Literals need no quotes while no letter reads as a code: `0 шт`, `#,##0 ₽` and `0 x` open;
 * `0 m`, `0 kg`, `0 USD` and `0 шт/ч` do not, `0 "kg"` and `0 "шт/ч"` do.
 */
import { EN_CODES, parseFormat } from './parse.js'

/** A longer code Excel cuts to this or refuses: never what was asked for. */
const MAX_CODE = 255
/** The letters of dates, times and eras, in either case: beside a digit placeholder they refuse a section. */
const DATE_LETTERS = /[bdeghmnsy]/i
/** Conditions a code may have (`[<1]0;[<2]0;0`). */
const MAX_CONDITIONS = 2
/** What a section's text compiles to at most, in Excel's units: a run of quoted text 2 and 2 for each UTF-16 unit, a bare literal 3 for each. */
const MAX_SECTION = 252
const CHAR_COST = 3
const QUOTED_COST = 2

interface Read {
  /** What Excel reads as codes: quoted and escaped text, paddings and fills out; an elapsed time's bracket left as its letter, the others gone. */
  bare: string
  conditions: number
  /** What quoted and escaped text and brackets cost of the section's size. */
  cost: number
  /** Whether it ends in quoted text. */
  quotedLast: boolean
}

/** A section as Excel reads it; null when something does not close or means nothing. */
function readSection(section: string): Read | null {
  const out: Read = { bare: '', conditions: 0, cost: 0, quotedLast: false }
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
      if (j < 0 || parseFormat(bracket, EN_CODES) === null) return null
      if (/^\[[hms]+\]$/i.test(bracket)) out.bare += 'h'
      if (/^\[[<>=]/.test(bracket)) out.conditions++
      i = j
    } else out.bare += ch
  }
  return out
}

/** Whether Excel reads the codes of a section, by its place among the code's sections. */
function sectionReads(found: string, last: boolean): boolean {
  const exp = /E[+-]/.exec(found)
  if (exp && !(/[0#?]/.test(found.slice(0, exp.index)) && /[0#?]/.test(found.slice(exp.index + 2)))) return false
  const general = /General/i.test(found)
  // a word of the code's own is no date's letter, and an exponent no era; `AM/PM` shows beside digits too (`0 AM/PM`)
  const bare = found.replace(/General/gi, ' ').replace(/AM\/PM|A\/P/gi, ' ').replace(/E[+-]/g, ' ')
  const date = bare.search(DATE_LETTERS)
  // zeros after a point past a date's first code are the second's decimals (`ss.00`)
  const rest = date < 0 ? bare : `${bare.slice(0, date)}${bare.slice(date).replace(/\.0{1,3}(?!\d)/g, ' ')}`
  const digits = /[0#?]/.test(rest)
  if (date >= 0 && digits) return false
  if (general && (digits || bare.includes('@'))) return false
  if (bare.includes('@') && (digits || date >= 0 || !last)) return false
  if (date < 0 && rest.includes('/')) {
    const fraction = rest.split('/')
    if (fraction.length > 2 || !/[0-9#?]\s*$/.test(fraction[0]) || !/^\s*[0-9#?]/.test(fraction[1] ?? '')) return false
  }
  return true
}

/** A section's size in Excel's units: its text by `readSection`, a bare literal of the rest, digit placeholders in a run but its first. */
function sectionCost(read: Read): number {
  const literals = read.bare.replace(/General|[0#?.,%@/]|E[+-]/gi, '').length
  const placeholders = /[0#?]/.test(read.bare) ? 5 : 0
  // a bare literal at the end takes one past its own (`$` × 83 opens, × 84 does not)
  return read.cost + literals * CHAR_COST + placeholders + (literals && !read.quotedLast ? 1 : 0)
}

/** A code's sections, split where `;` is no literal; null when a quote or a bracket does not close. */
function sectionsOf(code: string): string[] | null {
  const sections: string[] = []
  let from = 0
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]
    if (ch === '"' || ch === '[') {
      i = code.indexOf(ch === '"' ? '"' : ']', i + 1)
      if (i < 0) return null
    } else if (ch === '\\' || ch === '_' || ch === '*') i++
    else if (ch === ';') {
      sections.push(code.slice(from, i))
      from = i + 1
    }
  }
  sections.push(code.slice(from))
  return sections
}

/** Whether Excel refuses a file with this number format code. */
export function formatRefused(code: string): boolean {
  if (code.length > MAX_CODE) return true
  const sections = sectionsOf(code)
  if (!sections || sections.length > 4) return true
  const read = sections.map(readSection)
  if (read.some((r) => r === null)) return true
  const reads = read as Read[]
  if (reads.reduce((n, r) => n + r.conditions, 0) > MAX_CONDITIONS) return true
  return !reads.every((r, i) => sectionReads(r.bare, i === reads.length - 1) && sectionCost(r) <= MAX_SECTION)
}
