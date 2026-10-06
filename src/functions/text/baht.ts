/** BAHTTEXT, this engine's own: a number in Thai words with baht and satang, as Excel writes it. */
import type { FnDef } from '../../args.js'
import { roundDigits } from '../common.js'

const DIGITS = ['ศูนย์', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า']
const PLACES = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน']
const MILLION = 'ล้าน'

/** Up to six digits: ten is สิบ, twenty ยี่สิบ, a final one after other digits เอ็ด. */
function group(n: number, after: boolean): string {
  let out = ''
  const ds = String(n).split('').map(Number)
  ds.forEach((d, i) => {
    const place = ds.length - 1 - i
    if (d === 0) return
    if (place === 1) out += (d === 1 ? '' : d === 2 ? 'ยี่' : DIGITS[d]) + PLACES[1]
    else if (place === 0) out += d === 1 && (after || n > 9) ? 'เอ็ด' : DIGITS[d]
    else out += DIGITS[d] + PLACES[place]
  })
  return out
}

/** A whole number in groups of six digits joined by ล้าน (million). */
function words(n: number): string {
  if (n < 1_000_000) return group(n, false)
  const rest = n % 1_000_000
  return words(Math.floor(n / 1_000_000)) + MILLION + (rest ? group(rest, true) : '')
}

function bahttext(x: number): string {
  const cents = Math.round(roundDigits(Math.abs(x), 2, Math.round) * 100)
  const baht = Math.floor(cents / 100)
  const satang = cents % 100
  // the minus stays on an amount rounded to zero
  const sign = x < 0 ? 'ลบ' : ''
  if (cents === 0) return sign + DIGITS[0] + 'บาทถ้วน'
  return sign + (baht ? words(baht) + 'บาท' : '') + (satang ? group(satang, false) + 'สตางค์' : 'ถ้วน')
}

export const functions: Record<string, FnDef> = {
  BAHTTEXT: { args: 'n', fn: bahttext },
}
