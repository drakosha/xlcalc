/**
 * Text in XML parts: entities, and Excel's `_xHHHH_` escapes of characters XML cannot hold (control
 * characters, a carriage return, lone surrogates). Text that itself reads `_x0041_` is written with its
 * underscore escaped (`_x005F_x0041_`), so it comes back as written.
 */

const NAMED: Record<string, string> = { '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'" }

/** Entities to characters in one pass: a character reference gives data that is not read again (`&#38;amp;` is `&amp;`). */
export function decodeXml(s: string): string {
  if (!s.includes('&')) return s
  return s.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[0-9A-Fa-f]+);/g, (m) => NAMED[m] ?? String.fromCodePoint(m[2] === 'x' ? parseInt(m.slice(3, -1), 16) : parseInt(m.slice(2, -1), 10)))
}

/** `_xHHHH_` to its character (`_x005F_` is `_`), one pass. */
export function decodeEscapes(s: string): string {
  if (!s.includes('_x')) return s
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, h: string) => String.fromCharCode(parseInt(h, 16)))
}

/** Entities to characters outside CDATA sections, whose content is taken as written (an attribute value holds none: it holds no `<`). */
function decodeData(s: string): string {
  if (!s.includes('<![CDATA[')) return decodeXml(s)
  let out = ''
  let at = 0
  for (let open = s.indexOf('<![CDATA[', at); open >= 0; open = s.indexOf('<![CDATA[', at)) {
    const close = s.indexOf(']]>', open + 9)
    if (close < 0) break
    out += decodeXml(s.slice(at, open)) + s.slice(open + 9, close)
    at = close + 3
  }
  return out + decodeXml(s.slice(at))
}

/** Text of an XML element (`<t>`, `<v>`): entities and CDATA sections, then Excel's escapes. */
export const readText = (s: string): string => decodeEscapes(decodeData(s))

const hex4 = (c: number) => `_x${c.toString(16).toUpperCase().padStart(4, '0')}_`

/**
 * Characters XML 1.0 cannot hold (controls but tab and line feed, a carriage return that parsing would turn into
 * a line feed, U+FFFE, U+FFFF, lone surrogates) as `_xHHHH_`, and the underscore of a literal `_xHHHH_`.
 */
const ESCAPED = /_(?=x[0-9A-Fa-f]{4}_)|[\x00-\x08\x0B-\x0D\x0E-\x1F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g

export const escapeChars = (s: string): string => s.replace(ESCAPED, (m) => hex4(m.charCodeAt(0)))

/** Element text: `&`, `<`, `>` as entities, after Excel's escapes. */
export function escapeText(s: string): string {
  const e = escapeChars(s)
  return /[&<>]/.test(e) ? e.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') : e
}

/** An attribute value: also quotes, and tab and line feed as character references (attribute parsing would turn them into spaces). */
export const escapeAttr = (s: string): string => escapeText(s).replace(/"/g, '&quot;').replace(/\n/g, '&#10;').replace(/\t/g, '&#9;')

/** Whether element text needs `xml:space="preserve"`: spaces at either end, or line breaks and tabs. */
export const needsPreserve = (s: string): boolean => /^\s|\s$|[\t\n\r]/.test(s)

/** Text as a text element (`<t>`), its spaces kept where they need it. */
export const textElement = (s: string): string => `<t${needsPreserve(s) ? ' xml:space="preserve"' : ''}>${escapeText(s)}</t>`
