/**
 * `xlcalc/regex`: the levels of match mode 3 (`XLOOKUP`, `XMATCH`) beyond the one built in, which
 * `registerRegex` loads for every calculation, a pool's threads too: `pcreSyntax`, Excel's syntax (PCRE2 10.43 and
 * later, with UTF and UCP) translated into JavaScript's regular expressions, and `pcreEngine`, a PCRE2 engine of
 * the package's own with Excel's match limit and the calculation's `timeoutMs` (`docs/regex.md`).
 */
import type { RegexLevel } from '../contract.js'
import { Matcher } from './match.js'
import { parse, RegexSyntaxError } from './parse.js'
import { compile } from './program.js'
import { pcreToRegExp } from './syntax.js'

/** Excel's syntax by JavaScript's regular expressions: fast, without Excel's match limit, a heavy pattern not stopped. */
export const pcreSyntax: RegexLevel = {
  kind: 'xlcalc-pcre-syntax',
  compile(pattern: string) {
    const re = pcreToRegExp(pattern)
    return re && { test: (s: string) => re.test(s) }
  },
}

/** The PCRE2 engine: Excel's syntax and match limit, its steps counted on the calculation's budget. */
export const pcreEngine: RegexLevel = {
  kind: 'xlcalc-pcre-engine',
  compile(pattern: string) {
    try {
      return new Matcher(compile(parse(pattern)))
    } catch (e) {
      if (e instanceof RegexSyntaxError) return null
      throw e
    }
  },
}
