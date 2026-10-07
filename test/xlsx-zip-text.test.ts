/**
 * The package layer: zip archives without dependencies (deflated and stored entries, names in any case, what is
 * refused) and the text of XML parts (entities, Excel's `_xHHHH_` escapes both ways, line ends).
 */
import { describe, expect, it } from 'vitest'
import { decodeEscapes, decodeXml, escapeAttr, escapeChars, escapeText, needsPreserve, readText } from '../src/xlsx/xml/text.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { zip } from '../src/xlsx/zip/write.js'

const enc = new TextEncoder()

/** A zip of one stored (not deflated) entry, written by hand. */
function stored(name: string, data: Uint8Array, flags = 0): Uint8Array {
  const n = enc.encode(name)
  const local = new Uint8Array(30 + n.length + data.length)
  const lv = new DataView(local.buffer)
  lv.setUint32(0, 0x04034b50, true)
  lv.setUint16(6, flags, true)
  lv.setUint32(18, data.length, true)
  lv.setUint32(22, data.length, true)
  lv.setUint16(26, n.length, true)
  local.set(n, 30)
  local.set(data, 30 + n.length)
  const dir = new Uint8Array(46 + n.length)
  const dv = new DataView(dir.buffer)
  dv.setUint32(0, 0x02014b50, true)
  dv.setUint16(8, flags, true)
  dv.setUint32(20, data.length, true)
  dv.setUint32(24, data.length, true)
  dv.setUint16(28, n.length, true)
  dir.set(n, 46)
  const end = new Uint8Array(22)
  const ev = new DataView(end.buffer)
  ev.setUint32(0, 0x06054b50, true)
  ev.setUint16(8, 1, true)
  ev.setUint16(10, 1, true)
  ev.setUint32(12, dir.length, true)
  ev.setUint32(16, local.length, true)
  const out = new Uint8Array(local.length + dir.length + end.length)
  out.set(local)
  out.set(dir, local.length)
  out.set(end, local.length + dir.length)
  return out
}

describe('xlsx: zip', () => {
  it('writes and reads back deflated entries; the same parts give the same bytes', async () => {
    const big = enc.encode('<row>'.repeat(50_000))
    const files = [{ name: '[Content_Types].xml', data: enc.encode('<Types/>') }, { name: 'xl/big.xml', data: big }, { name: 'empty', data: new Uint8Array() }]
    const a = await zip(files)
    expect(await zip(files)).toEqual(a)
    expect(a.length).toBeLessThan(big.length / 100)
    const archive = openArchive(a)
    expect(archive.names).toEqual(['[Content_Types].xml', 'xl/big.xml', 'empty'])
    expect(await archive.bytes('xl/big.xml')).toEqual(big)
    expect(await archive.text('empty')).toBe('')
    expect(await archive.bytes('missing')).toBeUndefined()
  })

  it('stored entries; names in any case; a comment after the archive', async () => {
    const data = enc.encode('<a>\r\nb\rc</a>')
    const bytes = stored('XL/Part.xml', data)
    const withComment = new Uint8Array(bytes.length + 5)
    withComment.set(bytes)
    new DataView(withComment.buffer).setUint16(bytes.length - 2, 5, true)
    const archive = openArchive(withComment)
    expect(await archive.bytes('xl/part.XML')).toEqual(data)
    expect(await archive.text('xl/part.xml')).toBe('<a>\nb\nc</a>')
  })

  it('refuses what it cannot read: not a zip, encrypted entries, a broken directory', async () => {
    expect(() => openArchive(enc.encode('PK not really'))).toThrow(/no end of central directory/)
    expect(() => openArchive(stored('a', enc.encode('x'), 1))).toThrow(/encrypted entry a/)
    const broken = stored('a', enc.encode('x'))
    new DataView(broken.buffer).setUint32(broken.length - 6, 3, true)
    expect(() => openArchive(broken)).toThrow(/broken central directory/)
  })

  it('refuses names that are not ASCII', async () => {
    await expect(zip([{ name: 'xl/лист.xml', data: new Uint8Array() }])).rejects.toThrow(/not ASCII/)
  })
})

describe('xlsx: XML text', () => {
  it('entities in one pass', () => {
    expect(decodeXml('&lt;&amp;&gt;&quot;&apos;&#65;&#x42;&#38;amp;')).toBe(`<&>"'AB&amp;`)
  })

  it("Excel's escapes: characters XML cannot hold, a carriage return, and literal text that reads like an escape", () => {
    const cases: [string, string][] = [
      ['plain', 'plain'],
      ['_x0041_', '_x005F_x0041_'],
      ['_x004_ and _xZZZZ_', '_x004_ and _xZZZZ_'],
      ['a\u000bb', 'a_x000B_b'],
      ['\u0000\u0008\u001f', '_x0000__x0008__x001F_'],
      ['tab\tlf\n', 'tab\tlf\n'],
      ['cr\r\n', 'cr_x000D_\n'],
      ['￾￿', '_xFFFE__xFFFF_'],
      ['lone \uD800 and \uDC00', 'lone _xD800_ and _xDC00_'],
      ['pair 😀', 'pair 😀'],
    ]
    for (const [text, escaped] of cases) {
      expect(escapeChars(text), text).toBe(escaped)
      expect(decodeEscapes(escaped), escaped).toBe(text)
    }
    expect(decodeEscapes('_x005f_x0041_')).toBe('_x0041_')
  })

  it('element text and attributes', () => {
    expect(escapeText('a < b & "c" \u000b')).toBe('a &lt; b &amp; "c" _x000B_')
    expect(escapeAttr('say "hi"\n\tto & <me>')).toBe('say &quot;hi&quot;&#10;&#9;to &amp; &lt;me&gt;')
    expect(readText('&lt;_x000B_&gt;')).toBe('<\u000b>')
    expect([' a', 'a ', 'a\nb', 'a\tb', 'a b'].map(needsPreserve)).toEqual([true, true, true, true, false])
  })
})
