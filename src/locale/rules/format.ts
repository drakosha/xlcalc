/**
 * TEXT's format codes in a locale other than en-US, measured in ru-RU: the decimal separator (`0,00`), the
 * group separator for grouping and scaling (`# ##0`, `0 ` thousands; a space before the currency symbol is a
 * literal, `# ##0 ₽`), the locale's date and time letters (`ДД.ММ.ГГГГ ч:мм:сс`: day, year, hours and seconds in
 * either case, `М` the month or, after hours or before seconds, minutes as en-US `m`, `м` minutes always), the
 * General format by its name, among literals too (`Основной "шт."`), `AM/PM`. Latin letters are literals.
 * Excel refuses a point before a digit placeholder in a number format and a space in a fraction format.
 */
import { generalFormat, type Dialect, type Tok } from '../../format.js'
import type { RegionalData } from './data.js'

/** A section without its quoted text and escaped characters. */
const bare = (sec: string) => sec.replace(/"[^"]*"?|\\[\s\S]/g, '')

export function formatDialect(data: RegionalData, number: (text: string) => number | null, numberText: (n: number) => string): Dialect {
  const { codes, decimal, group, locale } = data
  const general = codes.general.toLowerCase()
  // letters read in either case; the month and minute letters differ by case only
  const letters = new Map([codes.day, codes.year, codes.hour, codes.second].map((c, i) => [c.toLowerCase(), 'dyhs'[i]]))
  const dateLetter = (ch: string) => ch === codes.month || ch === codes.minute || letters.has(ch.toLowerCase())
  const sameLetter = (a: string, b: string) => (a === codes.month || a === codes.minute ? a === b : a.toLowerCase() === b.toLowerCase())

  function tokenize(sec: string): Tok[] | null {
    const out: Tok[] = []
    let point = false
    let i = 0
    while (i < sec.length) {
      const ch = sec[i]
      const rest = sec.slice(i)
      if (ch === '"') {
        const j = sec.indexOf('"', i + 1)
        if (j < 0) return null
        out.push({ t: 'lit', s: sec.slice(i + 1, j) })
        i = j + 1
        continue
      }
      if (ch === '\\') {
        out.push({ t: 'lit', s: sec[i + 1] ?? '' })
        i += 2
        continue
      }
      if (rest.toLowerCase().startsWith(general)) {
        out.push({ t: 'general' })
        i += general.length
        continue
      }
      const ampm = /^(AM\/PM|A\/P)/i.exec(rest)
      if (ampm) {
        out.push({ t: 'ampm', s: ampm[0] })
        i += ampm[0].length
        continue
      }
      if (dateLetter(ch)) {
        let j = i
        while (j < sec.length && sameLetter(ch, sec[j])) j++
        const en = ch === codes.month ? 'm' : ch === codes.minute ? 'n' : letters.get(ch.toLowerCase())!
        out.push({ t: 'date', s: en.repeat(j - i) })
        i = j
        continue
      }
      if (ch === '[' || ch === 'E' || ch === 'e' || ch === '*' || ch === '_') return null
      if (ch === '0' || ch === '#' || ch === '?') out.push({ t: 'ph', c: ch })
      // a second decimal separator is a literal (`0,0,`)
      else if (ch === decimal) out.push(point ? { t: 'lit', s: ch } : { t: 'dot' })
      else if (ch === group) out.push(sec.startsWith(data.currency.symbol, i + 1) ? { t: 'lit', s: ch } : { t: 'comma' })
      else if (ch === '%') out.push({ t: 'pct' })
      else out.push({ t: 'lit', s: ch })
      point ||= ch === decimal
      i++
    }
    return out
  }

  /** A point before a digit placeholder in a number section, or a space in a fraction section. */
  function refused(sec: string): boolean {
    const b = bare(sec)
    const date = [...b].some(dateLetter) || /AM\/PM|A\/P/i.test(b)
    return (!date && /\.[^]*[0#?]/.test(b)) || (/[0#?]\//.test(b) && b.includes(group))
  }

  return {
    style: {
      decimal,
      group,
      months: data.months,
      monthsShort: data.monthsShort,
      days: data.days,
      daysShort: data.daysShort,
      general: (n) => generalFormat(n).replace('.', decimal),
      text: numberText,
    },
    tokenize,
    refuses: (sections) => sections.some(refused),
    number,
    bool: (b) => locale.booleans[b ? 'TRUE' : 'FALSE'],
  }
}
