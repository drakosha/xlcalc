/** The order of text and the case of names. No imports: `values.ts` and `criteria.ts` both need it, and a cycle through them broke loading a subpath first. */

/**
 * String order for `<`/`>` by an `Intl.Collator` locale tag: case-insensitive, accent-sensitive. Without `Intl`
 * (QuickJS, for example) the order falls back to code units of the lower-cased strings.
 */
export const textOrder = (tag: string): ((a: string, b: string) => number) =>
  typeof Intl !== 'undefined' ? new Intl.Collator(tag, { sensitivity: 'accent' }).compare : (a, b) => ((a = a.toLowerCase()) < (b = b.toLowerCase()) ? -1 : a > b ? 1 : 0)

/** en-US's order of text: Latin before Cyrillic, as Excel en-US sorts. */
export const collator = { compare: /* @__PURE__ */ textOrder('en-US') }

/** The two letters whose upper case is an ASCII letter, which Excel for Windows keeps apart from it: the dotless ı (I) and the long ſ (S). */
const APART = /[ıſ]/

/**
 * The key a name compares by in any case (defined names, tables, sheets): Unicode's default upper case, the same in
 * every locale, a letter's whole (`straße` is `STRASSE`, `ﬁ` is `FI`), but the dotless ı and the long ſ stay apart from
 * I and S, as Excel for Windows compares defined names and tables (the table sets `name-case-held` and
 * `name-case-sheets-held`; Excel for Mac takes `ſ` for `S`). Excel compares sheets letter by letter (`ß` beside `SS`),
 * a known difference. Turkish rules pair `i` with `İ` instead, and how Excel tr-TR compares such names is not measured.
 */
export const caseKey = (s: string): string => (APART.test(s) ? s.replace(/[^ıſ]+/g, (run) => run.toUpperCase()) : s.toUpperCase())

/**
 * The key a word of formula text compares by (a function's name, TRUE and FALSE, LET's names, a locale's words): each
 * letter's own upper case where it is one letter, so `gauß` is no GAUSS and `ﬁnd` no FIND, the dotless ı and the long ſ
 * kept (`falſe` is no FALSE), as Excel for Windows reads them (the sets `name-case-held`).
 */
export function wordKey(s: string): string {
  const up = s.toUpperCase()
  if (up.length === s.length && !APART.test(s)) return up
  let out = ''
  for (const ch of s) {
    const u = APART.test(ch) ? ch : ch.toUpperCase()
    out += u.length === ch.length ? u : ch
  }
  return out
}
