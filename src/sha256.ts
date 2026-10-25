/**
 * SHA-256 (FIPS 180-4) in plain JS: synchronous, incremental, no platform APIs, so hashing works the
 * same in Node, browsers and bare JS engines. Strings are hashed as UTF-8, a lone surrogate as U+FFFD
 * (what `TextEncoder` and Node's `hash.update(string)` do), so digests match `node:crypto` byte for byte.
 * `sha256()` hashes with Node's own where there is one.
 */

type NodeCrypto = { createHash(algorithm: 'sha256'): { update(text: string): unknown; digest(encoding: 'hex'): string } }

/**
 * Node's crypto, taken without an import (`process.getBuiltinModule`): a bundle for a browser holds no Node module.
 * None where it throws: a Node built without OpenSSL, or one whose crypto has no SHA-256.
 */
const nodeCrypto = ((): NodeCrypto | undefined => {
  try {
    const c = (globalThis as { process?: { getBuiltinModule?(id: string): unknown } }).process?.getBuiltinModule?.('node:crypto') as NodeCrypto | undefined
    c?.createHash('sha256')
    return c
  } catch {
    return undefined
  }
})()

/** A hasher of strings: Node's SHA-256 where there is one (the same digests, three to four times faster), else `Sha256`. */
export function sha256(): { update(text: string): unknown; hex(): string } {
  if (!nodeCrypto) return new Sha256()
  const h = nodeCrypto.createHash('sha256')
  return { update: (text) => h.update(text), hex: () => h.digest('hex') }
}

const K = new Int32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
])

export class Sha256 {
  private readonly h = new Int32Array([0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19])
  private readonly w = new Int32Array(64)
  /** The pending block. */
  private readonly buf = new Uint8Array(64)
  private used = 0
  /** Total bytes hashed so far. */
  private length = 0

  /** Adds a string, encoded as UTF-8. */
  update(text: string): this {
    for (let i = 0; i < text.length; i++) {
      let c = text.charCodeAt(i)
      if (c < 0x80) {
        this.byte(c)
        continue
      }
      if (c < 0x800) {
        this.byte(0xc0 | (c >> 6))
        this.byte(0x80 | (c & 0x3f))
        continue
      }
      if (c >= 0xd800 && c <= 0xdfff) {
        const d = c <= 0xdbff && i + 1 < text.length ? text.charCodeAt(i + 1) : 0
        if (d >= 0xdc00 && d <= 0xdfff) {
          const cp = 0x10000 + ((c - 0xd800) << 10) + (d - 0xdc00)
          i++
          this.byte(0xf0 | (cp >> 18))
          this.byte(0x80 | ((cp >> 12) & 0x3f))
          this.byte(0x80 | ((cp >> 6) & 0x3f))
          this.byte(0x80 | (cp & 0x3f))
          continue
        }
        c = 0xfffd
      }
      this.byte(0xe0 | (c >> 12))
      this.byte(0x80 | ((c >> 6) & 0x3f))
      this.byte(0x80 | (c & 0x3f))
    }
    return this
  }

  /** The digest as lower-case hex. The hasher must not be used afterwards. */
  hex(): string {
    const bits = this.length * 8
    this.byte(0x80)
    while (this.used !== 56) this.byte(0)
    const hi = Math.floor(bits / 0x100000000)
    const lo = bits >>> 0
    for (const word of [hi, lo]) for (let s = 24; s >= 0; s -= 8) this.byte((word >>> s) & 0xff)
    let out = ''
    for (const x of this.h) out += (x >>> 0).toString(16).padStart(8, '0')
    return out
  }

  private byte(b: number): void {
    this.buf[this.used++] = b
    this.length++
    if (this.used === 64) {
      this.block()
      this.used = 0
    }
  }

  private block(): void {
    const { w, buf, h } = this
    for (let i = 0; i < 16; i++) w[i] = (buf[4 * i] << 24) | (buf[4 * i + 1] << 16) | (buf[4 * i + 2] << 8) | buf[4 * i + 3]
    for (let i = 16; i < 64; i++) {
      const x = w[i - 15]
      const y = w[i - 2]
      const s0 = ((x >>> 7) | (x << 25)) ^ ((x >>> 18) | (x << 14)) ^ (x >>> 3)
      const s1 = ((y >>> 17) | (y << 15)) ^ ((y >>> 19) | (y << 13)) ^ (y >>> 10)
      w[i] = (w[i - 16] + s0 + w[i - 7] + s1) | 0
    }
    let a = h[0]
    let b = h[1]
    let c = h[2]
    let d = h[3]
    let e = h[4]
    let f = h[5]
    let g = h[6]
    let k = h[7]
    for (let i = 0; i < 64; i++) {
      const S1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7))
      const t1 = (k + S1 + ((e & f) ^ (~e & g)) + K[i] + w[i]) | 0
      const S0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10))
      const t2 = (S0 + ((a & b) ^ (a & c) ^ (b & c))) | 0
      k = g
      g = f
      f = e
      e = (d + t1) | 0
      d = c
      c = b
      b = a
      a = (t1 + t2) | 0
    }
    h[0] += a
    h[1] += b
    h[2] += c
    h[3] += d
    h[4] += e
    h[5] += f
    h[6] += g
    h[7] += k
  }
}
