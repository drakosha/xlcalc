/**
 * The converter's XML buffer: what it writes as UTF-8 comes back as the same text, a lone surrogate (which UTF-8
 * cannot hold) comes back as the `_xHHHH_` the reader decodes, and the limits count bytes, not characters.
 */
import { describe, expect, it } from 'vitest'
import { Xml } from '../src/xlsb/xml.js'

const xml = (limits = {}) => new Xml('xl/worksheets/sheet1.bin', limits, { spent: 0 })

describe('xlsb XML buffer', () => {
  it('gives back text of every width, across growth and appended fragments', () => {
    const out = xml()
    const tail = out.fragment()
    const long = 'é'.repeat(5000)
    out.add('<c><v>Щ😀')
    out.add(long)
    tail.add('</v></c>')
    out.append(tail)
    expect(out.text()).toBe(`<c><v>Щ😀${long}</v></c>`)
  })

  it('writes a lone surrogate as the escape the reader decodes', () => {
    const out = xml()
    out.add('a\uD800b\uDC00c😀')
    expect(out.text()).toBe('a_xD800_b_xDC00_c😀')
  })

  it('counts UTF-8 bytes against the part and total limits', () => {
    const fits = xml({ entryBytes: 4 })
    fits.add('Щщ')
    expect(fits.text()).toBe('Щщ')
    expect(() => xml({ entryBytes: 3 }).add('Щщ')).toThrow(/more than 3 bytes/)
    expect(() => xml({ totalBytes: 3 }).add('Щщ')).toThrow(/past the 3 byte total/)
  })
})
