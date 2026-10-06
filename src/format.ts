/**
 * TEXT with Excel number formats (en-US): `;` sections (positive, negative, zero, text), quoted and
 * `\`-escaped literals, numbers (`0`, `#`, `?`, `.`, `,` for thousands grouping and scaling by 1000,
 * `%`), fractions (`# ?/?`, `?/??`, `# ?/8`), dates and times (`yyyy`, `yy`, `m`…`mmmm`, `d`…`dddd`,
 * `h`, `hh`, `mm` as minutes, `ss`, `AM/PM`), `General` and `@`. Exponents (`E+`), colours and bracketed
 * conditions are not supported: null (the caller turns it into #VALUE!).
 */
import { roundDigits } from './functions/common.js'
import { ymd } from './dates.js'
import { NUMBER_TEXT, readNumber } from './model.js'

type Tok = { t: 'lit'; s: string } | { t: 'ph'; c: '0' | '#' | '?' } | { t: 'dot' } | { t: 'comma' } | { t: 'pct' } | { t: 'date'; s: string } | { t: 'ampm'; s: string }

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

function splitSections(fmt: string): string[] {
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

function tokenize(sec: string): Tok[] | null {
  const out: Tok[] = []
  let i = 0
  while (i < sec.length) {
    const ch = sec[i]
    const rest = sec.slice(i)
    if (ch === '"') {
      const j = sec.indexOf('"', i + 1)
      if (j < 0) return null
      out.push({ t: 'lit', s: sec.slice(i + 1, j) })
      i = j + 1
    } else if (ch === '\\') {
      out.push({ t: 'lit', s: sec[i + 1] ?? '' })
      i += 2
    } else if (ch === '[' || ch === 'E' || ch === 'e' || ch === '*' || ch === '_') {
      return null
    } else if (ch === '0' || ch === '#' || ch === '?') {
      out.push({ t: 'ph', c: ch })
      i++
    } else if (ch === '.') {
      out.push({ t: 'dot' })
      i++
    } else if (ch === ',') {
      out.push({ t: 'comma' })
      i++
    } else if (ch === '%') {
      out.push({ t: 'pct' })
      i++
    } else if (/^(AM\/PM|A\/P)/i.test(rest)) {
      const m = /^(AM\/PM|A\/P)/i.exec(rest)![0]
      out.push({ t: 'ampm', s: m })
      i += m.length
    } else if (/[ymdhs]/i.test(ch)) {
      let j = i
      while (j < sec.length && sec[j].toLowerCase() === ch.toLowerCase()) j++
      out.push({ t: 'date', s: sec.slice(i, j).toLowerCase() })
      i = j
    } else {
      out.push({ t: 'lit', s: ch })
      i++
    }
  }
  return out
}

/** Round half away from zero on the decimal representation (1.005 → 1.01). */
export function roundHalfUp(abs: number, k: number): string {
  return roundDigits(abs, k, Math.round).toFixed(k)
}

function formatNumber(v: number, toks: Tok[], signed: boolean): string | null {
  const dot = toks.findIndex((t) => t.t === 'dot')
  const intEnd = dot < 0 ? toks.length : dot
  let scale = 0
  let grouping = false
  for (let i = 0; i < intEnd; i++) {
    if (toks[i].t !== 'comma') continue
    const before = toks.slice(0, i).some((t) => t.t === 'ph')
    const after = toks.slice(i + 1, intEnd).some((t) => t.t === 'ph')
    if (before && after) grouping = true
    else if (before) scale++
  }
  // commas after the decimal digits scale too (`0.0,,` shows millions)
  for (let i = intEnd + 1; i < toks.length; i++) if (toks[i].t === 'comma' && !toks.slice(i + 1).some((t) => t.t === 'ph')) scale++
  const pct = toks.filter((t) => t.t === 'pct').length
  let x = Math.abs(v) * Math.pow(100, pct) / Math.pow(1000, scale)
  const fracPh = dot < 0 ? [] : toks.slice(dot + 1).filter((t): t is Extract<Tok, { t: 'ph' }> => t.t === 'ph')
  const intPh = toks.slice(0, intEnd).filter((t): t is Extract<Tok, { t: 'ph' }> => t.t === 'ph')
  const fixed = roundHalfUp(x, fracPh.length)
  x = Number(fixed)
  let [ip, fp = ''] = fixed.split('.')
  const zeros = intPh.filter((p) => p.c === '0').length
  if (ip === '0' && zeros === 0) ip = ''
  ip = ip.padStart(zeros, '0')
  if (grouping) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  // integer part: right to left over digit placeholders, leftover digits go before the first one
  let out = ''
  let rest = ip
  const firstPh = toks.findIndex((t, i) => i < intEnd && t.t === 'ph')
  for (let i = intEnd - 1; i >= 0; i--) {
    const t = toks[i]
    if (t.t === 'ph') {
      const d = rest.length ? rest[rest.length - 1] : t.c === '0' ? '0' : t.c === '?' ? ' ' : ''
      rest = rest.slice(0, -1)
      out = (i === firstPh ? rest : '') + d + out
      if (i === firstPh) rest = ''
    } else if (t.t === 'lit') out = t.s + out
    else if (t.t === 'pct') out = '%' + out
  }
  if (firstPh < 0) out = ip + out
  // fractional part: trailing `#` drops zeros, `?` gives a space
  if (dot >= 0) {
    const digits = fp.split('')
    for (let k = fracPh.length - 1; k >= 0 && digits[k] === '0' && fracPh[k].c !== '0'; k--) digits[k] = fracPh[k].c === '?' ? ' ' : ''
    let k = 0
    out += '.'
    for (const t of toks.slice(dot + 1)) {
      if (t.t === 'ph') out += digits[k++] ?? ''
      else if (t.t === 'lit') out += t.s
      else if (t.t === 'pct') out += '%'
    }
  }
  const negative = signed && v < 0 && x !== 0
  return (negative ? '-' : '') + out
}

const isDigitLit = (t: Tok | undefined) => t?.t === 'lit' && /^[0-9]$/.test(t.s)

/**
 * A fraction: placeholders right before `/` are the numerator, placeholders before them (if any) the
 * integer part; after `/` either placeholders (the denominator has at most that many digits, the closest
 * fraction is taken) or a fixed denominator (`?/8`).
 */
function formatFraction(v: number, toks: Tok[], slash: number, signed: boolean): string | null {
  let i = slash - 1
  while (i >= 0 && toks[i].t === 'ph') i--
  const numPh = toks.slice(i + 1, slash) as Extract<Tok, { t: 'ph' }>[]
  const intToks = toks.slice(0, i + 1)
  const hasInt = intToks.some((t) => t.t === 'ph')
  let j = slash + 1
  let fixed = 0
  const denPh: Extract<Tok, { t: 'ph' }>[] = []
  if (isDigitLit(toks[j])) {
    let digits = ''
    for (; j < toks.length && (isDigitLit(toks[j]) || (toks[j].t === 'ph' && (toks[j] as Extract<Tok, { t: 'ph' }>).c === '0')); j++) digits += toks[j].t === 'lit' ? (toks[j] as Extract<Tok, { t: 'lit' }>).s : '0'
    fixed = Number(digits)
  } else for (; j < toks.length && toks[j].t === 'ph'; j++) denPh.push(toks[j] as Extract<Tok, { t: 'ph' }>)
  if (!numPh.length || (!fixed && !denPh.length)) return null
  const x = Math.abs(v)
  let ip = hasInt ? Math.floor(x) : 0
  const frac = x - ip
  let num = 0
  let den = fixed
  if (fixed) num = Math.round(frac * fixed)
  else {
    let best = Infinity
    for (let d = 1, max = 10 ** denPh.length - 1; d <= max; d++) {
      const n = Math.round(frac * d)
      const err = Math.abs(frac - n / d)
      if (err < best - 1e-12) [best, num, den] = [err, n, d]
    }
  }
  if (hasInt && num === den) [ip, num] = [ip + 1, 0]
  // unused digit places: `0` gives a zero before the digits, `?` a space (after the denominator's digits)
  const pad = (s: string, ph: Extract<Tok, { t: 'ph' }>[], left: boolean) => {
    const free = ph.slice(0, Math.max(0, ph.length - s.length))
    const zeros = '0'.repeat(free.filter((t) => t.c === '0').length)
    const spaces = ' '.repeat(free.filter((t) => t.c === '?').length)
    return left ? spaces + zeros + s : zeros + s + spaces
  }
  const intPart = hasInt ? formatNumber(ip, intToks, false)! : intToks.map((t) => (t.t === 'lit' ? t.s : '')).join('')
  const denText = fixed ? String(fixed) : pad(String(den), denPh, false)
  const fracText = pad(String(num), numPh, true) + '/' + denText
  const tail = toks.slice(j).map((t) => (t.t === 'lit' ? t.s : '')).join('')
  // a whole number: the fraction's place is blank
  const body = hasInt && num === 0 ? intPart + ' '.repeat(fracText.length) : intPart + fracText
  return (signed && v < 0 && (ip !== 0 || num !== 0) ? '-' : '') + body + tail
}

function formatDate(serial: number, toks: Tok[]): string | null {
  if (serial < 0) return null
  let day = Math.floor(serial)
  let secs = Math.round((serial - day) * 86_400)
  if (secs >= 86_400) [day, secs] = [day + 1, secs - 86_400]
  const ampm = toks.some((t) => t.t === 'ampm')
  // Excel's calendar: day 0 is 1900-01-00, day 60 the phantom 1900-02-29, day 1 a Sunday
  const { y: Y, m, d: D } = ymd(day)
  const M = m - 1
  const W = (day + 6) % 7
  const h = Math.floor(secs / 3600)
  const mi = Math.floor(secs / 60) % 60
  const s = secs % 60
  const two = (n: number) => String(n).padStart(2, '0')
  let out = ''
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i]
    if (t.t === 'lit') out += t.s
    else if (t.t === 'ampm') out += t.s.length > 3 ? (h < 12 ? 'AM' : 'PM') : h < 12 ? 'A' : 'P'
    else if (t.t === 'date') {
      const c = t.s[0]
      if (c === 'y') out += t.s.length <= 2 ? two(Y % 100) : String(Y)
      else if (c === 'd') out += t.s.length === 1 ? String(D) : t.s.length === 2 ? two(D) : t.s.length === 3 ? DAYS[W].slice(0, 3) : DAYS[W]
      else if (c === 'h') {
        const hh = ampm ? h % 12 || 12 : h
        out += t.s.length === 1 ? String(hh) : two(hh)
      } else if (c === 's') out += t.s.length === 1 ? String(s) : two(s)
      else if (c === 'm') {
        // minutes: after hours or before seconds (across separators)
        const prev = toks.slice(0, i).reverse().find((x) => x.t === 'date')
        const next = toks.slice(i + 1).find((x) => x.t === 'date')
        const minutes = t.s.length <= 2 && ((prev?.t === 'date' && prev.s[0] === 'h') || (next?.t === 'date' && next.s[0] === 's'))
        if (minutes) out += t.s.length === 1 ? String(mi) : two(mi)
        else out += t.s.length === 1 ? String(M + 1) : t.s.length === 2 ? two(M + 1) : t.s.length === 3 ? MONTHS[M].slice(0, 3) : t.s.length === 4 ? MONTHS[M] : MONTHS[M][0]
      }
    } else if (t.t === 'ph') out += t.c === '0' ? '0' : ''
    else if (t.t === 'dot') out += '.'
    else if (t.t === 'comma') out += ','
    else if (t.t === 'pct') out += '%'
  }
  return out
}

/**
 * A number as text the way Excel converts one (`&`, CONCAT, text arguments): rounded to 15 significant
 * digits, in positional notation unless that takes more than 20 characters (the sign aside), else in
 * scientific notation with an exponent of two digits at least (`1E+20`, `7.4500104300146E-06`).
 */
export function generalText(n: number): string {
  if (n === 0) return '0'
  const [m, e] = Math.abs(n).toExponential(14).split('e')
  const exp = Number(e)
  const digits = m.replace('.', '').replace(/0+$/, '')
  const fixed =
    exp < 0 ? `0.${'0'.repeat(-exp - 1)}${digits}` : digits.length <= exp + 1 ? digits.padEnd(exp + 1, '0') : `${digits.slice(0, exp + 1)}.${digits.slice(exp + 1)}`
  const text = fixed.length <= 20 ? fixed : `${digits[0]}${digits.length > 1 ? '.' + digits.slice(1) : ''}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`
  return n < 0 ? '-' + text : text
}

/**
 * A number in the General format of TEXT: at most 11 characters, the sign aside. Positional when that
 * shows all of the number's 15 significant digits or as many as scientific notation would (a fraction
 * keeps up to 9 decimals, `0.000123457`), else scientific with the mantissa rounded to fit
 * (`1.23457E-05`, `1E+11`).
 */
export function generalFormat(n: number): string {
  if (n === 0) return '0'
  const a = Math.abs(n)
  const [m, ex] = a.toExponential(14).split('e')
  const e = Number(ex)
  const need = m.replace('.', '').replace(/0+$/, '').length
  const sciCap = 11 - 3 - Math.max(2, String(Math.abs(e)).length)
  const fixedCap = e >= 0 ? (e < 11 ? Math.max(10, e + 1) : 0) : 10 + e
  let text: string | null = null
  if (fixedCap > 0 && (need <= fixedCap || fixedCap >= sciCap)) {
    const fixed = roundHalfUp(a, e >= 0 ? Math.max(0, 9 - e) : 9).replace(/(\.\d*?)0+$/, '$1').replace(/\.$/, '')
    if (fixed.length <= 11) text = fixed
  }
  if (text === null) {
    const [sm, se] = a.toExponential(sciCap - 1).split('e')
    text = `${sm.replace(/(\.\d*?)0+$/, '$1').replace(/\.$/, '')}E${se[0]}${se.slice(1).padStart(2, '0')}`
  }
  return n < 0 ? '-' + text : text
}

/** TEXT(value, format); null for an unsupported or invalid format. */
export function formatText(value: number | string | boolean, fmt: string): string | null {
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  if (fmt === '') return ''
  const sections = splitSections(fmt)
  if (typeof value === 'string') {
    const t = value.trim()
    const n = NUMBER_TEXT.test(t) ? readNumber(t) : NaN
    if (Number.isNaN(n)) {
      const textSec = sections.length >= 4 ? sections[3] : sections.find((s) => s.includes('@'))
      return textSec !== undefined ? textSec.replace(/@/g, value).replace(/"/g, '') : value
    }
    value = n
  }
  const v = value
  let sec = sections[0]
  let signed = true
  if (v < 0 && sections.length >= 2) {
    sec = sections[1]
    signed = false
  } else if (v === 0 && sections.length >= 3) sec = sections[2]
  // an empty section shows nothing (`0;` hides negative numbers)
  if (sec === '') return ''
  if (/^\s*general\s*$/i.test(sec)) return generalFormat(signed ? v : Math.abs(v))
  if (sec.trim() === '@') return generalText(v)
  const toks = tokenize(sec)
  if (!toks) return null
  if (toks.some((t) => t.t === 'date' || t.t === 'ampm')) return formatDate(v, toks)
  if (!toks.some((t) => t.t === 'ph')) return toks.map((t) => (t.t === 'lit' ? t.s : t.t === 'dot' ? '.' : t.t === 'comma' ? ',' : t.t === 'pct' ? '%' : '')).join('')
  const slash = toks.findIndex((t, i) => t.t === 'lit' && t.s === '/' && toks[i - 1]?.t === 'ph')
  if (slash >= 0) return formatFraction(v, toks, slash, signed)
  return formatNumber(v, toks, signed)
}
