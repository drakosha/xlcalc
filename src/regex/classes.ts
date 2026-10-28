/**
 * The classes of PCRE2 with UCP as JavaScript's classes (flag `v`), shared by the translation (`syntax.ts`) and the
 * engine, whose character tests are made from them (`chars.ts`): `\\h`, `\\v`, the POSIX classes, and `\\p{…}` by
 * PCRE2's loose names.
 */
import { S_IN, W_IN } from '../functions/lookup/regex-syntax.js'

export const W = `[${W_IN}]`
export const H = '[\\t \\xa0\\u1680\\u180e\\u2000-\\u200a\\u202f\\u205f\\u3000]'
export const V = '[\\n-\\r\\x85\\u2028\\u2029]'
export const S = `[${S_IN}]`
export const POSIX: Record<string, string> = {
  alpha: '\\p{L}', lower: '\\p{Ll}', upper: '\\p{Lu}', alnum: '[\\p{L}\\p{N}]', digit: '\\p{Nd}', xdigit: '[0-9A-Fa-f]', space: S,
  blank: H, cntrl: '\\p{Cc}', graph: '[\\p{L}\\p{M}\\p{N}\\p{P}\\p{S}\\p{Cf}]', print: '[\\p{L}\\p{M}\\p{N}\\p{P}\\p{S}\\p{Cf}\\p{Zs}]',
  punct: '[\\p{P}\\x24\\x2b\\x3c-\\x3e\\x5e\\x60\\x7c\\x7e]', word: W, ascii: '[\\0-\\x7f]',
}
export const ASCII_POSIX: Record<string, string> = { alpha: '[A-Za-z]', lower: '[a-z]', upper: '[A-Z]', alnum: '[0-9A-Za-z]', digit: '[0-9]', space: '[\\t-\\r ]', word: '[0-9A-Za-z_]', blank: '[\\t ]', cntrl: '[\\0-\\x1f\\x7f]', punct: '[\\x21-\\x2f\\x3a-\\x40\\x5b-\\x60\\x7b-\\x7e]', graph: '[\\x21-\\x7e]', print: '[\\x20-\\x7e]' }

const accepts = (p: string) => {
  try {
    return new RegExp(p, 'v') && true
  } catch {
    return false
  }
}
const title = (s: string) => s.trim().split(/[\s_-]+/).map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join('_')

/** `\p{name}` in JavaScript, or null: categories, scripts (a bare name is a script's extensions), PCRE2's `L&` and `X…` properties. */
export function property(name: string, neg: boolean, caseless: boolean): string | null {
  const special: Record<string, string> = { 'l&': 'LC', lc: 'LC', any: '[\\s\\S]', xan: '[\\p{L}\\p{N}]', xps: S, xsp: S, xwd: '[\\p{L}\\p{N}_]', xuc: '[$@`\\xa0-\\u{10ffff}]' }
  let k = name.replace(/[\s_-]/g, '').toLowerCase()
  // caseless, an upper, lower or title case letter is any cased letter (PCRE2 10.43)
  if (caseless && /^(lu|ll|lt|uppercaseletter|lowercaseletter|titlecaseletter)$/.test(k)) k = 'lc'
  const at = name.search(/[:=]/)
  const key = at < 0 ? '' : name.slice(0, at).replace(/[\s_-]/g, '').toLowerCase()
  const value = at < 0 ? name : name.slice(at + 1)
  const prefix = key === 'sc' || key === 'script' ? 'sc=' : key === 'scx' || key === 'scriptextensions' ? 'scx=' : key === 'gc' || key === 'generalcategory' ? '' : null
  const tries = at >= 0 ? (prefix === null ? [] : [prefix + value.trim(), prefix + title(value)]) : k in special ? [special[k]] : [name.trim(), title(name), name.trim().toUpperCase(), 'scx=' + name.trim(), 'scx=' + title(name)]
  for (const t of tries) {
    const p = t.startsWith('[') || t === 'LC' ? (t === 'LC' ? '\\p{LC}' : t) : `\\p{${t}}`
    if (accepts(p)) return neg ? `[^${p}]` : p
  }
  return null
}
