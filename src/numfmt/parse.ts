/**
 * Number format codes read into sections of tokens, by a dialect's letters and separators (`Codes`): an xlsx
 * file's codes and TEXT in en-US by `EN_CODES`, TEXT in another locale by its own (`locale/rules/format.ts`).
 * Brackets give a section its colour, its condition, a currency symbol and locale (`[$₽-419]`) or elapsed time
 * (`[h]`); `"…"` and `\x` are literals, `_x` a space as wide as `x`, `*x` a fill, `@` the text.
 */
import { remembered } from '../memo.js'
import { EN_COLORS, PALETTE } from './palette.js'

/**
 * A token of a section: `dot` the decimal point and `comma` the group separator (grouping or scaling), as the code writes them, `slash` a `/`
 * written bare (a date separator in a date), `exp` an exponent (`E+` shows the sign of the exponent always, `E-`
 * a minus only), a date part or an elapsed one by its en-US letters (`n` minutes whatever stands around them).
 */
export type Tok =
  | { t: 'lit'; s: string; point?: true }
  | { t: 'ph'; c: '0' | '#' | '?' }
  | { t: 'dot'; s: string }
  | { t: 'comma'; s: string }
  | { t: 'pct' }
  | { t: 'slash' }
  | { t: 'exp'; plus: boolean }
  | { t: 'date'; s: string }
  | { t: 'elapsed'; s: string }
  | { t: 'ampm'; s: string }
  | { t: 'general' }
  | { t: 'text' }
  | { t: 'pad'; s: string }
  | { t: 'fill'; s: string }

export type Ph = Extract<Tok, { t: 'ph' }>

export interface Cond {
  op: '<' | '<=' | '>' | '>=' | '=' | '<>'
  limit: number
}

export interface Section {
  toks: Tok[]
  /** `#RRGGBB`. */
  color?: string
  cond?: Cond
  /** The locale of `[$-409]`, upper-case hex (`409`, `F800`; `[$-x-sysdate]` is `F800`), or the tag after the hyphen in upper case. */
  lcid?: string
  /** `B2` before the code: the Hijri calendar (`numfmt/hijri.ts`). */
  hijri?: true
}

/** How a dialect writes format codes. */
export interface Codes {
  decimal: string
  group: string
  /** The name of the General format, lower-case. */
  general: string
  /** Colour names, lower-case, and the word of `[Color10]`. */
  colors: ReadonlyMap<string, string>
  colorWord: string
  /** The en-US letter (`y m d h s`, `n` for minutes always) of a date code's letter, or undefined. */
  letter(ch: string): string | undefined
  /** Whether `b` continues the date code that `a` begins. */
  same(a: string, b: string): boolean
  /** The currency symbol: a group separator right before it is a literal (`# ##0 ₽` in ru-RU). */
  currency?: string
  /** Whether the dialect refuses a section whatever the value (TEXT answers #VALUE!). */
  refuses?(section: string): boolean
}

/** en-US codes, as xlsx files store them: date letters in either case. */
export const EN_CODES: Codes = {
  decimal: '.',
  group: ',',
  general: 'general',
  colors: EN_COLORS,
  colorWord: 'color',
  letter: (ch) => {
    const c = ch.toLowerCase()
    return 'ymdhs'.includes(c) ? c : undefined
  },
  same: (a, b) => a.toLowerCase() === b.toLowerCase(),
}

/** A code split at `;` outside quotes and escapes. */
export function splitSections(fmt: string): string[] {
  const out: string[] = []
  let cur = ''
  let quoted = false
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i]
    if (ch === '"') quoted = !quoted
    if (ch === '\\' && !quoted && i + 1 < fmt.length) {
      cur += ch + fmt[++i]
      continue
    }
    if (ch === ';' && !quoted) {
      out.push(cur)
      cur = ''
    } else cur += ch
  }
  out.push(cur)
  return out
}

const COND = /^(<>|<=|>=|<|>|=)\s*(-?\d*[.,]?\d+(?:[eE][+-]?\d+)?)$/

/** The culture tags of the system's long date and time ([MS-XLSX] formatCode16), by the locale codes they stand for. */
const SYSTEM_TAGS: Record<string, string> = { 'X-SYSDATE': 'F800', 'X-SYSTIME': 'F400' }

/** A bracket's meaning for the section, or false when it means nothing (the code is invalid). */
function bracket(inner: string, sec: Section, out: Tok[], codes: Codes): boolean {
  if (inner[0] === '$') {
    // `[$<symbol>-<culture>]`: the culture after the first hyphen is a locale code (`409`) or a tag (`x-sysdate`, `ru-RU`)
    const dash = inner.indexOf('-')
    const symbol = dash < 0 ? inner.slice(1) : inner.slice(1, dash)
    if (symbol) out.push({ t: 'lit', s: symbol })
    if (dash >= 0) {
      const culture = inner.slice(dash + 1).toUpperCase()
      sec.lcid = SYSTEM_TAGS[culture] ?? culture
    }
    return true
  }
  const unit = inner ? codes.letter(inner[0]) : undefined
  if (unit && 'hmns'.includes(unit) && [...inner].every((c) => codes.same(inner[0], c))) {
    out.push({ t: 'elapsed', s: (unit === 'm' ? 'n' : unit).repeat(inner.length) })
    return true
  }
  const name = inner.toLowerCase()
  const color = codes.colors.get(name) ?? (name.startsWith(codes.colorWord) ? PALETTE[Number(name.slice(codes.colorWord.length).trim()) - 1] : undefined)
  if (color) sec.color = color
  else {
    const c = COND.exec(inner)
    if (!c) return false
    sec.cond = { op: c[1] as Cond['op'], limit: Number(c[2].replace(codes.decimal, '.')) }
  }
  return true
}

/** A section's tokens and bracketed parts; null for a code the dialect does not read. */
function parseSection(text: string, codes: Codes): Section | null {
  if (codes.refuses?.(text)) return null
  const toks: Tok[] = []
  const sec: Section = { toks }
  let point = false
  let i = 0
  while (i < text.length) {
    const ch = text[i]
    const rest = text.slice(i)
    if (ch === '"') {
      const j = text.indexOf('"', i + 1)
      if (j < 0) return null
      toks.push({ t: 'lit', s: text.slice(i + 1, j) })
      i = j + 1
      continue
    }
    if (ch === '\\' || ch === '_' || ch === '*') {
      const next = text[i + 1] ?? ''
      toks.push(ch === '\\' ? { t: 'lit', s: next } : ch === '_' ? { t: 'pad', s: next } : { t: 'fill', s: next })
      i += 2
      continue
    }
    if (ch === '[') {
      const j = text.indexOf(']', i)
      if (j < 0 || !bracket(text.slice(i + 1, j), sec, toks, codes)) return null
      i = j + 1
      continue
    }
    // the calendar first in a section (ECMA-376 numFmts): B1 the Gregorian, which names months and days in English as
    // `[$-409]` does, B2 the Hijri (Excel for Mac in en-US and ru-RU)
    if (!toks.length && /^b[12]/i.test(rest)) {
      if (rest[1] === '1') sec.lcid = '409'
      else sec.hijri = true
      i += 2
      continue
    }
    if (rest.toLowerCase().startsWith(codes.general)) {
      toks.push({ t: 'general' })
      i += codes.general.length
      continue
    }
    const ampm = /^(AM\/PM|A\/P)/i.exec(rest)
    if (ampm) {
      toks.push({ t: 'ampm', s: ampm[0] })
      i += ampm[0].length
      continue
    }
    if ((ch === 'E' || ch === 'e') && (text[i + 1] === '+' || text[i + 1] === '-')) {
      // Excel takes an upper-case exponent only
      if (ch === 'e') return null
      toks.push({ t: 'exp', plus: text[i + 1] === '+' })
      i += 2
      continue
    }
    const letter = codes.letter(ch)
    if (letter) {
      let j = i
      while (j < text.length && codes.same(ch, text[j])) j++
      toks.push({ t: 'date', s: letter.repeat(j - i) })
      i = j
      continue
    }
    if (ch === 'E' || ch === 'e') return null
    if (ch === '0' || ch === '#' || ch === '?') toks.push({ t: 'ph', c: ch })
    // a second decimal separator is a literal (`0.0.`), marked: in a date it may start the second's decimals (`yy.00.00`)
    else if (ch === codes.decimal) toks.push(point ? { t: 'lit', s: ch, point: true } : { t: 'dot', s: ch })
    else if (ch === codes.group) toks.push(codes.currency && text.startsWith(codes.currency, i + 1) ? { t: 'lit', s: ch } : { t: 'comma', s: ch })
    else if (ch === '%') toks.push({ t: 'pct' })
    else if (ch === '/') toks.push({ t: 'slash' })
    else if (ch === '@') toks.push({ t: 'text' })
    else toks.push({ t: 'lit', s: ch })
    point ||= ch === codes.decimal
    i++
  }
  return sec
}

/** Parsed codes by dialect: books use few, and TEXT over a column repeats one. */
const CACHE = new WeakMap<Codes, Map<string, Section[] | null>>()

/** A code's sections; null when one of them is not read. */
export function parseFormat(code: string, codes: Codes): Section[] | null {
  let cache = CACHE.get(codes)
  if (!cache) CACHE.set(codes, (cache = new Map()))
  return remembered(cache, code, 4096, () => {
    const secs = splitSections(code).map((s) => parseSection(s, codes))
    return secs.every((s): s is Section => s !== null) ? secs : null
  })
}
