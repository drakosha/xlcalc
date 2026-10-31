/** What the writer's parts share: the XML declaration and the SpreadsheetML namespace they start with, colours, and text cut to Excel's limits. */
import type { Issues, Where, XlsxMessage } from '../issues.js'
import { escapeAttr } from '../xml/text.js'

export const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
export const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
/** The namespace of a part's relationships (`r:id`), which their types begin with. */
export const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
/** The namespace of a package's relationships parts. */
export const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
/** Excel 2010's namespace (`x14`) and Excel 2013's (`x15`). */
export const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
export const X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'
/** What the content types of SpreadsheetML parts begin with. */
export const TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

/** A colour (`#RRGGBB`) as a file's opaque ARGB. */
export const rgb = (c: string) => `FF${c.slice(1).toUpperCase()}`

export { MAX_TEXT } from '../../model.js'

/**
 * Attributes of `o` by a table of entries `attr=key:kind` in their order, as `take` reads them: `s` text (the default),
 * `n` a number where given, `p` a number above 0, `t` a flag true by default, written `0` where false, `f` a flag false
 * by default, written `1` where true, `b` a flag written always; an entry with its value (`attr="v"`) as it is.
 */
export function give(o: object, spec: string): string {
  let out = ''
  for (const e of spec.split(' ')) {
    if (e.includes('"')) {
      out += ` ${e}`
      continue
    }
    const [name, kind = 's'] = e.split(':')
    const [attr, key = attr] = name.split('=')
    const v = (o as Record<string, unknown>)[key]
    if (kind === 'b') out += ` ${attr}="${v ? 1 : 0}"`
    else if (kind === 't' ? v === false : kind === 'f' || kind === 'p' ? v : v !== undefined) out += ` ${attr}="${kind === 't' ? 0 : kind === 'f' ? 1 : kind === 's' ? escapeAttr(v as string) : v}"`
  }
  return out
}

/** A list element of `items` with their count: `<cols count="2">…</cols>`. */
export const counted = (tag: string, items: string[]) => `<${tag} count="${items.length}">${items.join('')}</${tag}>`

/** `counted` where there are items, else nothing. */
export const countedIf = (tag: string, items: string[] | undefined) => (items?.length ? counted(tag, items) : '')

/** Text cut to Excel's limit `max`, reported as `textLength` when it was longer; a character of two code units is not split. */
export function clip(text: string, max: number, issues: Issues, message: XlsxMessage, where: Where): string {
  if (text.length <= max) return text
  issues.add('textLength', message, where)
  const high = (text.charCodeAt(max - 1) & 0xfc00) === 0xd800 && (text.charCodeAt(max) & 0xfc00) === 0xdc00
  return text.slice(0, high ? max - 1 : max)
}

/** A part's relationships, `rId1` on; an external one (a hyperlink's address, another workbook's path) keeps its target as written. */
export const relsXml = (items: [type: string, target: string, external?: boolean][]) =>
  `${HEAD}<Relationships xmlns="${PKG_REL}">${items.map(([type, target, external], i) => `<Relationship Id="rId${i + 1}" Type="${type}" Target="${external ? escapeAttr(target) : target}"${external ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`
