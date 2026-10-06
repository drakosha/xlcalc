/**
 * Excel en-US cases for the parts of complex text: whether a real or an imaginary part is read as VALUE reads a
 * number (`"50%+i"`, `"$1+2i"`, `"1,000-2i"`) or only as a plain decimal. Computed as dynamic array formulas, in
 * Excel en-US (`test/excel/cases/complex-text.json`).
 */
import type { Probe } from '../probes.js'
import { xl } from './arrays/common.js'

const TEXTS = [
  '50%+i', '1+50%i', '50%-50%i', '%50+i', '$1+2i', '1+$2i', '$1,000.5-2i', '1,000+2i', '1+1,000i', '1,00+2i', '(1)+2i', '1+(2)i', '-(1)+i',
  '1 1/2+i', '1+1 1/2i', '1/2+3i', '1/2/2025+i', '18:30+i', '1+18:30i', '6 PM+i', '.5+i', '1.+i', '1.5e1+2i', '1E+2+3i', '1+2E-1i', '1e+i',
  ' 1 + 2 i', '+1+2i', '+1-+2i', '1+-2i', '1--2i', 'TRUE+i', '1+TRUE i', '1 000+2i', '1+1 000i', '1.5.5+i', '0x1+i', '1+2ii', '1+2j', '50%j',
  '1E+308+i', '1E+309+i', '9.99999999999999E+307+i', '1+ i', 'i', '-i', '50%', '$1', '1,000', '(5)',
]

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const add = (fn: string, formula: string) => out.push({ id: `${fn}#${out.filter((p) => p.fn === fn).length + 1}`, group: fn, fn, formula: xl(formula) })
  for (const t of TEXTS) {
    add('IMREAL', `IMREAL("${t}")`)
    add('IMAGINARY', `IMAGINARY("${t}")`)
  }
  for (const t of ['50%+i', '$1+2i', '1,000+2i', '1+50%i', '(1)+2i']) {
    add('IMSUM', `IMSUM("${t}","1+i")`)
    add('IMABS', `IMABS("${t}")`)
  }
  return out
}
