/**
 * Functions of East Asian text. ASC, DBCS and JIS: full-width (double-byte) forms of letters, digits, punctuation and katakana to half-width ones and
 * back, as Windows maps them (`LCMapString` with `LCMAP_HALFWIDTH`, `LCMAP_FULLWIDTH`): ASCII from `!` to `~` and the
 * space against U+FF01–U+FF5E and U+3000, the half-width katakana against the full-width ones (a voiced katakana
 * against its letter and the sound mark), and the signs ¢ £ ¬ ¯ ¦ ¥ ₩ against U+FFE0–U+FFE6. PHONETIC: the readings
 * (furigana) a file stores with a cell's text, which the workbook does not keep, so a cell's own text.
 */
import type { Arg, Cx, FnDef } from '../../args.js'
import { E, isErr, Rng } from '../../values.js'

/** The full-width forms of U+FF61–U+FF9F, in order. */
const KANA = '。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン゛゜'
const SIGNS = '¢£¬¯¦¥₩'
const VOICED = 0xff9e
const SEMI = 0xff9f

const toFull = new Map<number, string>()
const toHalf = new Map<string, string>()
/** A full-width letter and a half-width sound mark after it, as the voiced letter (カ and ﾞ: ガ). */
const voiced = new Map<string, string>()
for (let i = 0; i < KANA.length; i++) toFull.set(0xff61 + i, KANA[i])
for (let i = 0; i < SIGNS.length; i++) toFull.set(SIGNS.codePointAt(i)!, String.fromCharCode(0xffe0 + i))
for (const [h, f] of toFull) toHalf.set(f, String.fromCodePoint(h))
const HA = [0x30cf, 0x30d2, 0x30d5, 0x30d8, 0x30db]
for (const [h, f] of [...toFull]) {
  const c = f.charCodeAt(0)
  const marks: [number, string][] = []
  if (c === 0x30a6) marks.push([VOICED, 'ヴ'])
  if ((c >= 0x30ab && c <= 0x30c8 && c !== 0x30c3) || HA.includes(c)) marks.push([VOICED, String.fromCharCode(c + 1)])
  if (HA.includes(c)) marks.push([SEMI, String.fromCharCode(c + 2)])
  for (const [mark, v] of marks) {
    voiced.set(f + String.fromCharCode(mark), v)
    toHalf.set(v, String.fromCodePoint(h, mark))
  }
}

/** Half-width forms to full width; a half-width letter and its sound mark become the voiced letter. */
function full(s: string): string {
  let out = ''
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i)
    if (c === 0x20) out += '\u3000'
    else if (c > 0x20 && c < 0x7f) out += String.fromCharCode(c + 0xfee0)
    else {
      const f = toFull.get(c)
      const v = f && voiced.get(f + s.charAt(i + 1))
      if (v) i++
      out += v || f || s[i]
    }
  }
  return out
}

/** Full-width forms to half width. */
function half(s: string): string {
  let out = ''
  for (const ch of s) {
    const c = ch.charCodeAt(0)
    if (c === 0x3000) out += ' '
    else if (c >= 0xff01 && c <= 0xff5e) out += String.fromCharCode(c - 0xfee0)
    else out += toHalf.get(ch) ?? ch
  }
  return out
}

/** PHONETIC of a reference: the texts of its cells in order, joined (numbers and other values give nothing); anything else #N/A. */
function phonetic(cx: Cx, x: Arg) {
  if (isErr(x)) return x
  if (!(x instanceof Rng)) return E.NA
  let out = ''
  for (const v of cx.flat(x)) if (typeof v === 'string') out += v
  return out
}

export const functions: Record<string, FnDef> = {
  ASC: { args: 's', fn: half },
  DBCS: { args: 's', fn: full },
  JIS: { args: 's', fn: full },
  PHONETIC: { args: 'R', cx: true, fn: phonetic },
}
