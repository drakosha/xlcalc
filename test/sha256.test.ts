/**
 * The plain-JS SHA-256 against published test vectors and against `node:crypto` (the 0.1.0 engine hashed
 * with it), and the synthetic book hashes pinned: with every earlier configuration version, and with the
 * current one (a change of values or of CONFIG_VERSION has to update this pin consciously).
 */
import { createHash } from 'node:crypto'
import { describe, expect, it } from 'vitest'
import { compute, type CellValue } from '../src/index.js'
import { inputHash, valuesHash } from '../src/canon.js'
import { Sha256 } from '../src/sha256.js'
import { mulberry32, synthBook } from '../src/synth.js'

const js = (...parts: string[]) => parts.reduce((h, p) => h.update(p), new Sha256()).hex()
const node = (...parts: string[]) => parts.reduce((h, p) => h.update(p), createHash('sha256')).digest('hex')

/** A seeded string mixing ASCII, Cyrillic, 3-byte, astral and lone surrogate code units. */
function randomText(rnd: () => number, length: number): string {
  const pools = [
    () => 32 + Math.floor(rnd() * 95),
    () => 0x410 + Math.floor(rnd() * 64),
    () => 0x800 + Math.floor(rnd() * 0xd000),
    () => 0xd800 + Math.floor(rnd() * 0x800),
  ]
  let s = ''
  while (s.length < length) {
    const x = rnd()
    if (x < 0.05) s += String.fromCodePoint(0x10000 + Math.floor(rnd() * 0xfffff))
    else s += String.fromCharCode(pools[Math.floor(rnd() * pools.length)]())
  }
  return s
}

describe('sha256', () => {
  it('FIPS 180-4 test vectors', () => {
    expect(js('')).toBe('e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855')
    expect(js('abc')).toBe('ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad')
    expect(js('abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq')).toBe('248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1')
    expect(js('abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu')).toBe(
      'cf5b16a778af8380036ce59e7b0492370b249b11e8f07a51afac45037afee9d1',
    )
    expect(js('a'.repeat(1_000_000))).toBe('cdc76e5c9914fb9281a1c7e284d73e67f1809a48a497200e046d39ccc7112cd0')
  })

  it('every length around the block boundaries matches node:crypto', () => {
    for (let n = 0; n <= 200; n++) {
      const s = 'x'.repeat(n)
      expect(js(s), String(n)).toBe(node(s))
    }
  })

  it('UTF-8: Cyrillic, astral characters and lone surrogates match node:crypto', () => {
    for (const s of ['Данные', 'Ёлка «й»', '€', '😀', 'a\ud800', '\udc00b', '\udc00\ud800', '\ud83d', 'x😀\ud83d']) expect(js(s), JSON.stringify(s)).toBe(node(s))
    // a surrogate pair split across two updates is two lone surrogates, as in node:crypto
    expect(js('\ud83d', '\ude00')).toBe(node('\ud83d', '\ude00'))
  })

  it('500 random strings split into random updates match node:crypto', () => {
    const rnd = mulberry32(11)
    for (let i = 0; i < 500; i++) {
      const s = randomText(rnd, Math.floor(rnd() * 600))
      const parts: string[] = []
      for (let at = 0; at < s.length; ) {
        const len = 1 + Math.floor(rnd() * 130)
        parts.push(s.slice(at, at + len))
        at += len
      }
      expect(js(...parts), `case ${i}`).toBe(node(...parts))
    }
  })

  it('the synthetic book hashes: those of earlier configuration versions, and the current one', () => {
    const { book } = synthBook({ seed: 1 })
    const r = compute(book, { clock: { now: '2026-10-04T12:00:00Z', seed: 7 } })
    const values = new Map(r.sheets.map((s) => [s.name, new Map(Object.entries(s.cells).map(([a, c]): [string, CellValue] => [a, c.value]))]))
    // 0.2.0 and the formula-engine development builds hashed with these versions; the hash of the same values must not move
    expect(valuesHash(values, 'formula-engine/1;compiled-js;formulajs@4.6.1;canon15')).toBe('18be025cb4581543a8953790d380ecf9f3383002d8112eb6b15da4f182c0bba3')
    expect(valuesHash(values, 'formula-engine/2;compiled-js;canon15')).toBe('8d63bad2c247ae26aa632f2823e29d07947af262a623365e03d5856e67dd30fc')
    expect(inputHash(book)).toBe('1fb3575501a5acbff64c388b43b5e3a1176e312243f01de1e4610e4ac7bc67a2')
    // the configuration version carries the package name: renaming the package changes this hash too
    expect(r.configVersion).toBe('xlcalc/1;canon15')
    expect(r.hash).toBe('101578a9c6b600b197bfd8f96d2d28eb3ffc19044a947949d8473d45ce2ed901')
  })
})
