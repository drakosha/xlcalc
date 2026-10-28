/**
 * PCRE2's commonest constructs in JavaScript's regular expressions, shared by match mode 3 without a registered level
 * (`rewrite`) and by the full translation of `xlcalc/regex` (`regex/syntax.ts`): `\d \s \w` and their opposites by
 * Unicode, as PCRE2 with UCP takes them (`\w` is letters, digits, nonspacing marks and connectors), `\b \B` by them,
 * `\A \z \Z`, `{,m}`, and a leading `(?i)`.
 */

/** The insides of `\d`, `\s` and `\w` as classes take them (`\s` is `\p{Z}`, `\h` and `\v`). */
export const D_IN = '\\p{Nd}'
export const S_IN = '\\p{Z}\\t-\\r\\x85\\u180e'
export const W_IN = '\\p{L}\\p{N}\\p{Mn}\\p{Pc}'
const ASCII = { d: '0-9', s: '\\t-\\r ', w: '0-9A-Za-z_' }

/** `\d \D \s \S \w \W` as a JavaScript class, or null for another letter; `ascii` restricts one to ASCII (`(?aD)`…). */
export function typeClass(c: string, ascii: Partial<Record<'d' | 's' | 'w', boolean>> = {}): string | null {
  const k = c.toLowerCase() as 'd' | 's' | 'w'
  const inner = k === 'd' ? D_IN : k === 's' ? S_IN : k === 'w' ? W_IN : null
  if (inner === null) return null
  return `[${c === k ? '' : '^'}${ascii[k] ? ASCII[k] : inner}]`
}

/** `\b` or `\B` by the word characters `w` (a class). */
export const boundary = (w: string, not: boolean) => (not ? `(?:(?<=${w})(?=${w})|(?<!${w})(?!${w}))` : `(?:(?<=${w})(?!${w})|(?<!${w})(?=${w}))`)

/** The end of the text, or a newline that ends it (`\Z`, `$`), or with `multi` any newline (`(?m)$`); `nl` the newline characters of a class. */
export const textEnd = (nl: string, multi = false) => (multi ? `(?=\\r\\n|[${nl}]|(?![\\s\\S]))` : `(?=(?:\\r\\n|[${nl}])?(?![\\s\\S]))`)

/** `\A` (and `\G`, the start of the subject) and `\z`. */
export const START = '(?<![\\s\\S])'
export const END = '(?![\\s\\S])'

/** A quantifier in braces at `i` of `p` (`{2}`, `{2,}`, `{,3}`, spaces inside) as JavaScript writes it, and its length; null for a `{` that is a literal. */
export function braces(p: string, i: number): [string, number] | null {
  const m = /^\{\s*(\d*)\s*(?:(,)\s*(\d*)\s*)?\}/.exec(p.slice(i))
  if (!m || (m[1] === '' && (m[2] === undefined || m[3] === ''))) return null
  return [`{${m[1] || 0}${m[2] ?? ''}${m[3] ?? ''}}`, m[0].length]
}

/**
 * A pattern of match mode 3 without a registered level: JavaScript's syntax (flag `u`) with the constructs above
 * rewritten as PCRE2 takes them; null if it is no pattern. In a class `\d`, `\D`, `\s` and `\w` are Unicode's, `\S`
 * and `\W` stay JavaScript's.
 */
export function rewrite(p: string): RegExp | null {
  const caseless = p.startsWith('(?i)')
  let out = ''
  let inClass = false
  for (let i = caseless ? 4 : 0; i < p.length; i++) {
    const c = p[i]
    if (c === '\\' && i + 1 < p.length) {
      const e = p[++i]
      const t = typeClass(e)
      if (t && inClass) out += e === 'D' ? '\\P{Nd}' : e === e.toLowerCase() ? t.slice(1, -1) : `\\${e}`
      else if (t) out += t
      else if (inClass) out += `\\${e}`
      else if (e === 'b' || e === 'B') out += boundary(`[${W_IN}]`, e === 'B')
      else if (e === 'A') out += START
      else if (e === 'z') out += END
      else if (e === 'Z') out += textEnd('\\n\\r')
      else out += `\\${e}`
      continue
    }
    const q = c === '{' && !inClass ? braces(p, i) : null
    if (q) {
      out += q[0]
      i += q[1] - 1
      continue
    }
    if (c === '[' && !inClass) {
      inClass = true
      out += c
      // a leading ] or ^] stays as JavaScript reads it
      if (p[i + 1] === '^') out += p[++i]
      continue
    }
    if (c === ']') inClass = false
    out += c
  }
  try {
    return new RegExp(out, caseless ? 'iu' : 'u')
  } catch {
    return null
  }
}
