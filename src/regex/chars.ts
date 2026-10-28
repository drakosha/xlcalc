/**
 * Characters for the regular expressions of match mode 3, as PCRE2 reads them with UTF and UCP on (Excel for Mac):
 * code points of a UTF-16 text, Unicode properties, `\d \w \s \h \v` and the POSIX classes by Unicode properties,
 * caseless matching by case folding, and the characters that end a line.
 */

/** A test of one code point. */
export type CharTest = (c: number) => boolean

/** The code point at `i` of `s` (a surrogate pair is one). */
export function cpAt(s: string, i: number): number {
  const c = s.charCodeAt(i)
  if ((c & 0xfc00) === 0xd800 && i + 1 < s.length) {
    const d = s.charCodeAt(i + 1)
    if ((d & 0xfc00) === 0xdc00) return ((c - 0xd800) << 10) + (d - 0xdc00) + 0x10000
  }
  return c
}

/** The code point that ends at `i` of `s`. */
export function cpBefore(s: string, i: number): number {
  const d = s.charCodeAt(i - 1)
  if ((d & 0xfc00) === 0xdc00 && i >= 2) {
    const c = s.charCodeAt(i - 2)
    if ((c & 0xfc00) === 0xd800) return ((c - 0xd800) << 10) + (d - 0xdc00) + 0x10000
  }
  return d
}

/** UTF-16 units of code point `c`. */
export const width = (c: number) => (c > 0xffff ? 2 : 1)

/** `t` remembered for the characters of the BMP it was asked about: a property's regular expression runs once a character. */
export function remembered(t: CharTest): CharTest {
  let seen: Uint8Array | null = null
  return (c) => {
    if (c > 0xffff) return t(c)
    seen ??= new Uint8Array(0x10000)
    const v = seen[c]
    if (v !== 0) return v === 2
    const r = t(c)
    seen[c] = r ? 2 : 1
    return r
  }
}

/** A Unicode property as JavaScript writes it (`L`, `Script_Extensions=Greek`), a test of a code point, or null if unknown. */
function jsProperty(name: string): CharTest | null {
  let re: RegExp
  try {
    re = new RegExp(`\\p{${name}}`, 'u')
  } catch {
    return null
  }
  return remembered((c) => re.test(String.fromCodePoint(c)))
}

const memo = new Map<string, CharTest | null>()
const property = (name: string): CharTest | null => {
  if (!memo.has(name)) memo.set(name, jsProperty(name))
  return memo.get(name)!
}

const L = (c: number) => property('L')!(c)
const N = (c: number) => property('N')!(c)
const H = new Set([0x09, 0x20, 0xa0, 0x1680, 0x180e, 0x2000, 0x2001, 0x2002, 0x2003, 0x2004, 0x2005, 0x2006, 0x2007, 0x2008, 0x2009, 0x200a, 0x202f, 0x205f, 0x3000])
const V = new Set([0x0a, 0x0b, 0x0c, 0x0d, 0x85, 0x2028, 0x2029])

/** `\h` and `\v`. */
export const horizontal: CharTest = (c) => H.has(c)
export const vertical: CharTest = (c) => V.has(c)
const space: CharTest = /* @__PURE__ */ remembered((c) => (c < 128 ? c === 0x20 || (c >= 9 && c <= 13) : V.has(c) || property('Z')!(c)))
const word: CharTest = /* @__PURE__ */ remembered((c) => (c < 128 ? (c >= 48 && c <= 57) || ((c | 32) >= 97 && (c | 32) <= 122) || c === 95 : L(c) || N(c) || property('Mn')!(c) || property('Pc')!(c)))
const asciiDigit: CharTest = (c) => c >= 48 && c <= 57
const asciiWord: CharTest = (c) => asciiDigit(c) || ((c | 32) >= 97 && (c | 32) <= 122) || c === 95
const asciiSpace: CharTest = (c) => c === 0x20 || (c >= 9 && c <= 13)

/** ASCII restrictions of `(?aD)`, `(?aS)`, `(?aW)`, `(?aP)`, `(?aT)`. */
export interface Ascii {
  aD: boolean
  aS: boolean
  aW: boolean
  aP: boolean
  aT: boolean
}

/** `\d`, `\s` and `\w` (Unicode under UCP, ASCII under their restriction). */
export function typeTest(k: 'd' | 's' | 'w', a: Ascii): CharTest {
  if (k === 'd') return a.aD ? asciiDigit : (c) => (c < 128 ? asciiDigit(c) : property('Nd')!(c))
  if (k === 's') return a.aS ? asciiSpace : space
  return a.aW ? asciiWord : word
}

/** A word character for `\b` and `\B`. */
export const wordTest = (a: Ascii): CharTest => (a.aW ? asciiWord : word)

const POSIX_ASCII: Record<string, CharTest> = {
  alpha: (c) => (c | 32) >= 97 && (c | 32) <= 122,
  lower: (c) => c >= 97 && c <= 122,
  upper: (c) => c >= 65 && c <= 90,
  alnum: (c) => asciiDigit(c) || ((c | 32) >= 97 && (c | 32) <= 122),
  digit: asciiDigit,
  xdigit: (c) => asciiDigit(c) || ((c | 32) >= 97 && (c | 32) <= 102),
  space: asciiSpace,
  blank: (c) => c === 0x20 || c === 9,
  cntrl: (c) => c < 32 || c === 127,
  graph: (c) => c > 32 && c < 127,
  print: (c) => c >= 32 && c < 127,
  punct: (c) => c > 32 && c < 127 && !asciiDigit(c) && !((c | 32) >= 97 && (c | 32) <= 122),
  word: asciiWord,
  ascii: (c) => c < 128,
}

const graph: CharTest = (c) => c !== 0x61c && c !== 0x180e && !(c >= 0x2066 && c <= 0x2069) && (property('L')!(c) || property('M')!(c) || N(c) || property('P')!(c) || property('S')!(c) || property('Cf')!(c))

/** The POSIX classes under UCP (`pcre2pattern`, "POSIX character classes"): `[:digit:]` follows `\d`, `[:xdigit:]` stays ASCII. */
const POSIX_UCP: Record<string, CharTest> = {
  alpha: L,
  lower: (c) => property('Ll')!(c),
  upper: (c) => property('Lu')!(c),
  alnum: (c) => L(c) || N(c),
  space,
  blank: horizontal,
  cntrl: (c) => property('Cc')!(c),
  graph,
  print: (c) => graph(c) || (property('Zs')!(c) && c !== 0x180e),
  punct: (c) => property('P')!(c) || (c < 256 && property('S')!(c)),
  word,
}

/** A POSIX class by name (`alpha`), or null if there is none; `caseless` makes `lower` and `upper` any cased letter. */
export function posixTest(name: string, a: Ascii, caseless: boolean): CharTest | null {
  if (!(name in POSIX_ASCII)) return null
  if (caseless && (name === 'lower' || name === 'upper')) return a.aP ? POSIX_ASCII.alpha : (c) => property('LC')!(c)
  if (name === 'digit') return a.aP || a.aT ? asciiDigit : typeTest('d', { ...a, aD: false })
  return a.aP || !(name in POSIX_UCP) ? POSIX_ASCII[name] : POSIX_UCP[name]
}

const loose = (s: string) => s.replace(/[\s_-]/g, '').toLowerCase()

/** General categories by their long names, loosely (`uppercaseletter` is `Lu`). */
const CATEGORY_NAMES =
  'C Other Cc Control Cf Format Cn Unassigned Co Private_Use Cs Surrogate L Letter Ll Lowercase_Letter Lm Modifier_Letter Lo Other_Letter Lt Titlecase_Letter Lu Uppercase_Letter LC Cased_Letter M Mark Mc Spacing_Mark Me Enclosing_Mark Mn Nonspacing_Mark N Number Nd Decimal_Number Nl Letter_Number No Other_Number P Punctuation Pc Connector_Punctuation Pd Dash_Punctuation Pe Close_Punctuation Pf Final_Punctuation Pi Initial_Punctuation Po Other_Punctuation Ps Open_Punctuation S Symbol Sc Currency_Symbol Sk Modifier_Symbol Sm Math_Symbol So Other_Symbol Z Separator Zl Line_Separator Zp Paragraph_Separator Zs Space_Separator'.split(' ')
const CATEGORIES = new Map<string, string>()
for (let i = 0; i < CATEGORY_NAMES.length; i += 2) {
  CATEGORIES.set(loose(CATEGORY_NAMES[i]), CATEGORY_NAMES[i])
  CATEGORIES.set(loose(CATEGORY_NAMES[i + 1]), CATEGORY_NAMES[i])
}
CATEGORIES.set('l&', 'LC')

/** PCRE2's special properties (`Xan`, `Xwd`…). */
const SPECIAL: Record<string, CharTest> = {
  any: () => true,
  xan: (c) => L(c) || N(c),
  xps: space,
  xsp: space,
  xwd: (c) => L(c) || N(c) || c === 95,
  xuc: (c) => c === 0x24 || c === 0x40 || c === 0x60 || (c >= 0xa0 && !(c >= 0xd800 && c <= 0xdfff)),
}

/** Words of a loose name in Title_Case, as JavaScript names scripts and binary properties (`old italic` is `Old_Italic`). */
const titled = (s: string) => s.trim().split(/[\s_-]+/).map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join('_')

/** A value of a script or binary property as JavaScript knows it: as written, in Title_Case, in capitals. */
function known(prefix: string, value: string): CharTest | null {
  for (const v of [value.trim(), titled(value), value.trim().toUpperCase()]) {
    const t = property(prefix + v)
    if (t) return t
  }
  return null
}

/**
 * `\p{name}` (`name` as written between the braces or the one letter of `\pL`): a general category, a script (its
 * extensions, as a bare name), `sc:`/`scx:` with a script, a binary property or a special one; null if unknown.
 * `caseless` makes `Lu`, `Ll` and `Lt` any cased letter, as PCRE2 10.43 does.
 */
export function propertyTest(name: string, caseless: boolean): CharTest | null {
  const at = name.search(/[:=]/)
  if (at >= 0) {
    const key = loose(name.slice(0, at))
    const value = name.slice(at + 1)
    if (key === 'sc' || key === 'script') return known('Script=', value)
    if (key === 'scx' || key === 'scriptextensions') return known('Script_Extensions=', value)
    if (key === 'gc' || key === 'generalcategory') return propertyTest(value, caseless)
    return null
  }
  const k = loose(name)
  const gc = CATEGORIES.get(k)
  if (gc) return property(caseless && (gc === 'Lu' || gc === 'Ll' || gc === 'Lt') ? 'LC' : gc)
  if (k in SPECIAL) return SPECIAL[k]
  return known('Script_Extensions=', name) ?? known('', name)
}

const ONE = /^[\s\S]$/u
const single = (s: string) => (ONE.test(s) ? s.codePointAt(0)! : -1)
const folds = new Map<number, number>()

/** The case fold of `c`: its lower case after its upper case (`ς`, `σ` and `Σ` are `σ`; the Kelvin sign is `k`). */
export function fold(c: number): number {
  if (c < 128) return c >= 65 && c <= 90 ? c + 32 : c
  let f = folds.get(c)
  if (f === undefined) {
    const s = String.fromCodePoint(c)
    const up = single(s.toUpperCase())
    const low = single(String.fromCodePoint(up < 0 ? c : up).toLowerCase())
    const lower = single(s.toLowerCase())
    f = low >= 0 ? low : lower >= 0 ? lower : c
    folds.set(c, f)
  }
  return f
}

/** Caseless equality; `restrict` (`(?r)`) keeps ASCII and other characters apart. */
export const sameCase = (a: number, b: number, restrict: boolean) => a === b || (fold(a) === fold(b) && (!restrict || a < 128 === b < 128))

/** `c` and the characters of its case: its fold, lower and upper case, the upper case of its fold. */
export function caseVariants(c: number, restrict: boolean): number[] {
  const s = String.fromCodePoint(c)
  const f = fold(c)
  const out = [c, f, single(s.toLowerCase()), single(s.toUpperCase()), single(String.fromCodePoint(f).toUpperCase())]
  return out.filter((v, i) => v >= 0 && out.indexOf(v) === i && (!restrict || v < 128 === c < 128))
}

/** A character class: ranges of code points (caseless: any character of their case) and tests, or the opposite. */
export function classTest(ranges: readonly number[], tests: readonly CharTest[], negated: boolean, caseless: boolean, restrict: boolean): CharTest {
  const inRanges = (c: number) => {
    for (let i = 0; i < ranges.length; i += 2) if (c >= ranges[i] && c <= ranges[i + 1]) return true
    return false
  }
  const has = (c: number) => inRanges(c) || (caseless && caseVariants(c, restrict).some(inRanges)) || tests.some((t) => t(c))
  return remembered((c) => has(c) !== negated)
}

/** Newline conventions: which characters end a line (`(*CR)`…, ANYCRLF by default as in Excel). */
export const Newline = { CR: 0, LF: 1, CRLF: 2, ANYCRLF: 3, ANY: 4, NUL: 5 } as const
export type Newline = (typeof Newline)[keyof typeof Newline]

const anyNewline = (c: number) => c === 10 || c === 11 || c === 12 || c === 13 || c === 0x85 || c === 0x2028 || c === 0x2029

/** The length of the newline that starts at `i` of `s`, 0 if none. */
export function newlineAt(s: string, i: number, nl: Newline): number {
  if (i >= s.length) return 0
  const c = s.charCodeAt(i)
  const crlf = c === 13 && s.charCodeAt(i + 1) === 10
  switch (nl) {
    case Newline.LF:
      return c === 10 ? 1 : 0
    case Newline.CR:
      return c === 13 ? 1 : 0
    case Newline.CRLF:
      return crlf ? 2 : 0
    case Newline.NUL:
      return c === 0 ? 1 : 0
    case Newline.ANYCRLF:
      return crlf ? 2 : c === 10 || c === 13 ? 1 : 0
    default:
      return crlf ? 2 : anyNewline(c) ? 1 : 0
  }
}

/** Whether a newline ends at `i` of `s` (not between the CR and LF of a CRLF that is one newline). */
export function newlineBefore(s: string, i: number, nl: Newline): boolean {
  if (i === 0) return false
  if (nl === Newline.CRLF) return i >= 2 && s.charCodeAt(i - 2) === 13 && s.charCodeAt(i - 1) === 10
  // the CR of a CRLF starts a newline of two units: none ends between them
  return newlineAt(s, i - 1, nl) === 1
}
