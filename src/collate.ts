/** The order of text. No imports: `values.ts` and `criteria.ts` both need it, and a cycle through them broke loading a subpath first. */

/**
 * String order for `<`/`>` by an `Intl.Collator` locale tag: case-insensitive, accent-sensitive. Without `Intl`
 * (QuickJS, for example) the order falls back to code units of the lower-cased strings.
 */
export const textOrder = (tag: string): ((a: string, b: string) => number) =>
  typeof Intl !== 'undefined' ? new Intl.Collator(tag, { sensitivity: 'accent' }).compare : (a, b) => ((a = a.toLowerCase()) < (b = b.toLowerCase()) ? -1 : a > b ? 1 : 0)

/** en-US's order of text: Latin before Cyrillic, as Excel en-US sorts. */
export const collator = { compare: textOrder('en-US') }
