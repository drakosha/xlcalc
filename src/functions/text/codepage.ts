/**
 * CHAR and CODE by the Windows code page of the book's locale (`Rules.codePage`): Windows-1252 in en-US, Windows-1251 in
 * ru-RU (`src/locale/rules/ru-RU.ts`), as Excel for Windows computes them. Excel for Mac answers codes 128–255 by Mac Roman
 * and Mac Cyrillic instead, and 95 (`_`) for a character outside its page where Windows answers 63 (`?`); codes 1–127 and
 * the argument edges are as both compute them (`bench/coverage/sets/char-code.ts`). Undefined bytes of a page are the C1
 * controls of the same value, as Windows converts them (`test/codepage.test.ts` checks the pages against the published ones).
 */
import type { Cx, FnDef } from '../../args.js'
import { E } from '../../values.js'

/** Excel for Windows's code of a character outside its code page: `?`, as Windows converts it (Excel for Mac answers 95, `_`). */
const OUTSIDE = 63

export const functions: Record<string, FnDef> = {
  CHAR: { args: 'i', cx: true, fn: (cx: Cx, n: number) => (n < 1 || n > 255 ? E.VALUE : n < 128 ? String.fromCharCode(n) : cx.rules.codePage[n - 128]) },
  CODE: {
    args: 's',
    cx: true,
    fn: (cx: Cx, s: string) => {
      if (s === '') return E.VALUE
      const c = s.charCodeAt(0)
      if (c < 128) return c
      const at = cx.rules.codePage.indexOf(s[0])
      return at < 0 ? OUTSIDE : at + 128
    },
  },
}
