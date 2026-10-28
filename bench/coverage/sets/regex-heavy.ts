/**
 * The sets of match mode 3 whose heavy patterns JavaScript's regular expressions backtrack over for hours where
 * Excel's match limit stops them (`docs/regex.md`): only the PCRE2 engine of `xlcalc/regex` calculates them all, so
 * `test/regex-levels.test.ts` alone does, and the harness's `xlsx saved` step does not send them.
 */
export const REGEX_HEAVY_SETS: ReadonlySet<string> = new Set(['xmatch-regex-own-held', 'xmatch-regex-own-held-2'])
