/**
 * CHAR and CODE by the Windows code page of the book's locale: Windows-1252 in en-US, Windows-1251 in ru-RU, as
 * Excel for Windows computes them. Excel for Mac answers codes 128–255 by Mac Roman and Mac Cyrillic instead;
 * codes 1–127, the argument edges and the code of a character outside the page are as Excel computes them
 * (`bench/coverage/sets/char-code.ts`). Undefined bytes of a page are the C1 controls of the same value, as
 * Windows converts them (`test/codepage.test.ts` checks the tables against the published ones).
 */
import type { Cx, FnDef } from '../../args.js'
import { E } from '../../values.js'

const range = (from: number, count: number) => Array.from({ length: count }, (_, i) => from + i)

/** Characters of codes 128–255 by locale tag. */
export const CODE_PAGES: Readonly<Record<string, string>> = {
  'en-US': String.fromCharCode(
    0x20ac, 0x0081, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008d, 0x017d, 0x008f,
    0x0090, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x009d, 0x017e, 0x0178,
    ...range(0xa0, 96),
  ),
  'ru-RU': String.fromCharCode(
    0x0402, 0x0403, 0x201a, 0x0453, 0x201e, 0x2026, 0x2020, 0x2021, 0x20ac, 0x2030, 0x0409, 0x2039, 0x040a, 0x040c, 0x040b, 0x040f,
    0x0452, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x0098, 0x2122, 0x0459, 0x203a, 0x045a, 0x045c, 0x045b, 0x045f,
    0x00a0, 0x040e, 0x045e, 0x0408, 0x00a4, 0x0490, 0x00a6, 0x00a7, 0x0401, 0x00a9, 0x0404, 0x00ab, 0x00ac, 0x00ad, 0x00ae, 0x0407,
    0x00b0, 0x00b1, 0x0406, 0x0456, 0x0491, 0x00b5, 0x00b6, 0x00b7, 0x0451, 0x2116, 0x0454, 0x00bb, 0x0458, 0x0405, 0x0455, 0x0457,
    ...range(0x0410, 64),
  ),
}

/** Excel's code of a character outside its code page: `_` (measured on Mac Roman and Mac Cyrillic). */
const OUTSIDE = 95

const page = (cx: Cx) => CODE_PAGES[cx.rules.locale.tag]

export const functions: Record<string, FnDef> = {
  CHAR: { args: 'i', cx: true, fn: (cx: Cx, n: number) => (n < 1 || n > 255 ? E.VALUE : n < 128 ? String.fromCharCode(n) : page(cx)[n - 128]) },
  CODE: {
    args: 's',
    cx: true,
    fn: (cx: Cx, s: string) => {
      if (s === '') return E.VALUE
      const c = s.charCodeAt(0)
      if (c < 128) return c
      const at = page(cx).indexOf(s[0])
      return at < 0 ? OUTSIDE : at + 128
    },
  },
}
