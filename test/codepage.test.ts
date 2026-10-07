/**
 * The code pages of CHAR and CODE (`src/functions/text/codepage.ts`) against the published tables: Unicode's
 * mappings of Windows-1252 and Windows-1251 (`test/fixtures/unicode/`, from
 * https://www.unicode.org/Public/MAPPINGS/VENDORS/MICSFT/WINDOWS/) and the platform's `TextDecoder`, where it
 * has them. A byte the published table leaves undefined is the C1 control of the same value, as Windows and
 * `TextDecoder` convert it.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { CODE_PAGES } from '../src/functions/text/codepage.js'
import { LOCALES } from '../src/locale/registry.js'
import { compute } from '../src/index.js'

const PAGES = { 'en-US': 'windows-1252', 'ru-RU': 'windows-1251' } as const

/** Bytes 128–255 of a published table: the character, or the C1 control of the same value where it is undefined. */
function published(name: string): string {
  const text = readFileSync(new URL(`./fixtures/unicode/${name.replace('windows-', 'CP')}.TXT`, import.meta.url), 'utf8')
  const map = new Map<number, number>()
  for (const m of text.matchAll(/^0x([0-9A-F]{2})\t0x([0-9A-F]{4})\t/gm)) map.set(parseInt(m[1], 16), parseInt(m[2], 16))
  return String.fromCharCode(...Array.from({ length: 128 }, (_, i) => map.get(128 + i) ?? 128 + i))
}

const decoder = (name: string) => {
  try {
    return new TextDecoder(name)
  } catch {
    return null
  }
}

describe('code pages of CHAR and CODE', () => {
  it('every locale has one', () => expect(Object.keys(CODE_PAGES).sort()).toEqual([...LOCALES.keys()].sort()))

  for (const [tag, name] of Object.entries(PAGES)) {
    it(`${tag}: ${name} as Unicode publishes it`, () => expect(CODE_PAGES[tag]).toBe(published(name)))

    it.skipIf(!decoder(name))(`${tag}: ${name} as TextDecoder reads it`, () => expect(CODE_PAGES[tag]).toBe(decoder(name)!.decode(Uint8Array.from({ length: 128 }, (_, i) => 128 + i))))

    it(`${tag}: CODE(CHAR(n)) is n for 1 to 255, a character outside the page is 63`, () => {
      const cells = Object.fromEntries(Array.from({ length: 255 }, (_, i) => [`A${i + 1}`, { formula: `CODE(CHAR(${i + 1}))` }]))
      Object.assign(cells, { B1: { formula: 'CODE("中")' }, B2: { formula: `CODE("${tag === 'en-US' ? 'Ж' : 'é'}")` } })
      const out = compute({ locale: tag, sheets: [{ name: 'S', cells }] }).sheets[0].cells
      expect(Array.from({ length: 255 }, (_, i) => out[`A${i + 1}`].value)).toEqual(Array.from({ length: 255 }, (_, i) => i + 1))
      expect([out.B1.value, out.B2.value]).toEqual([63, 63])
    })
  }
})
