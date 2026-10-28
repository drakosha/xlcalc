/**
 * Characters for the regular expressions of match mode 3, as PCRE2 reads them with UTF and UCP on (Excel for Mac):
 * code points of a UTF-16 text, tests of `\d \w \s \h \v`, the POSIX classes and `\p{…}` made from JavaScript's classes
 * (`classes.ts`, shared with the translation), caseless matching by case folding, and the characters that end a line.
 */
import { typeClass } from '../functions/lookup/regex-syntax.js'
import { ASCII_POSIX, H, POSIX, property, V } from './classes.js'

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
  // ASCII in a small table, the rest of the BMP in one made at its first character
  const ascii = new Uint8Array(128)
  let bmp: Uint8Array | null = null
  return (c) => {
    if (c > 0xffff) return t(c)
    const seen = c < 128 ? ascii : (bmp ??= new Uint8Array(0x10000))
    const v = seen[c]
    if (v !== 0) return v === 2
    const r = t(c)
    seen[c] = r ? 2 : 1
    return r
  }
}

/** A test of a JavaScript class (`classes.ts`), remembered for the BMP; memoized by the class. */
const tests = new Map<string, CharTest>()
function classOf(js: string): CharTest {
  let t = tests.get(js)
  if (!t) {
    const re = new RegExp(js, 'v')
    t = remembered((c) => re.test(String.fromCodePoint(c)))
    tests.set(js, t)
  }
  return t
}

/** `\h` and `\v`. */
export const horizontal: CharTest = (c) => classOf(H)(c)
export const vertical: CharTest = (c) => classOf(V)(c)

/** ASCII restrictions of `(?aD)`, `(?aS)`, `(?aW)`, `(?aP)`, `(?aT)`. */
export interface Ascii {
  aD: boolean
  aS: boolean
  aW: boolean
  aP: boolean
  aT: boolean
}

/** `\d`, `\s` and `\w` (Unicode under UCP, ASCII under their restriction). */
export const typeTest = (k: 'd' | 's' | 'w', a: Ascii): CharTest => classOf(typeClass(k, { d: a.aD, s: a.aS, w: a.aW })!)

/** A word character for `\b` and `\B`. */
export const wordTest = (a: Ascii): CharTest => typeTest('w', a)

/** A POSIX class by name (`alpha`), or null if there is none; `caseless` makes `lower` and `upper` any cased letter. */
export function posixTest(name: string, a: Ascii, caseless: boolean): CharTest | null {
  const ascii = a.aP || (a.aT && name === 'digit')
  const js = caseless && !ascii && (name === 'lower' || name === 'upper') ? '\\p{LC}' : (ascii && ASCII_POSIX[name]) || POSIX[name]
  return js ? classOf(js) : null
}

/** `\p{name}` (`name` as written between the braces or the one letter of `\pL`), or null if unknown; `caseless` makes `Lu`, `Ll` and `Lt` any cased letter. */
export function propertyTest(name: string, caseless: boolean): CharTest | null {
  const js = property(name, false, caseless)
  return js ? classOf(js) : null
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
    // the dotless i is its own: its upper case I folds to i (Excel: `(?i)ı` finds neither I nor i)
    f = c === 0x131 ? c : low >= 0 ? low : lower >= 0 ? lower : c
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
  return out.filter((v, i) => v >= 0 && out.indexOf(v) === i && fold(v) === f && (!restrict || v < 128 === c < 128))
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
