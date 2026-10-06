/**
 * Excel probes of the engine's rules outside the formulajs functions: numbers as text (`&`, CONCAT, text
 * arguments, TEXT with General), reading long number literals and numeric text, criteria over whole columns of
 * sheets with different numbers of filled rows. Excel en-US's answers are in `test/excel/probes.json`, checked
 * by `test/excel-probes.test.ts`.
 */

export interface Probe {
  id: string
  group: string
  formula: string
  /** The function the case checks (in the sets `sets/*`): acceptance sees by it that the function has cases. */
  fn?: string
}

/** A set's data sheets: a cell's value or formula (as Excel writes it, with `_xlfn.` prefixes). */
export type SheetCells = Record<string, Record<string, number | string | boolean | { formula: string }>>

/** The probes' data sheets: S2!A1:A5 = 1…5, S3!B1:B3 = 1, −1, 1. */
export const PROBE_SHEETS: Record<string, Record<string, number>> = {
  S2: { A1: 1, A2: 2, A3: 3, A4: 4, A5: 5 },
  S3: { B1: 1, B2: -1, B3: 1 },
}

/** A number literal with mantissa m and exponent e, as written in a formula: `1.5E-07`. */
const lit = (m: string, e: number) => `${m}E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const add = (group: string, formula: string) => out.push({ id: `${group}#${out.filter((p) => p.group === group).length + 1}`, group, formula })

  // numbers as text: mantissas of 1…15 significant digits at exponents −22…22
  for (const m of ['1', '1.5', '7.45', '1.2345', '-1.5', '1.23456789', '1.2345678901234', '1.23456789012345', '-1.23456789012345'])
    for (let e = -22; e <= 22; e++) add('text', `"a"&${lit(m, e)}`)
  // more than 15 significant digits: a quotient of literals
  for (const [n, d] of [['1', '3'], ['2', '3'], ['-1', '7']] as const) for (let e = -20; e <= 20; e++) add('text', `"a"&(${n}/${lit(d, e)})`)
  add('text', '"a"&(1/134228)')
  add('text', '"a"&0.0000000745')
  add('text', '"a"&0.00000745')
  add('text', '"a"&0')
  add('text', '"a"&-0.000001')
  // other paths of the same conversion
  for (const x of ['0.0000000745', '(1/134228)', '1E+21', '1E+15', '123456789012345', '(1/3)', '(2/3E-15)', '-1.5E-11', '1E+20', '(10/3)'])
    for (const f of [`_xlfn.CONCAT(${x})`, `LEN(${x})`, `${x}&"a"`, `LEFT(${x},30)`, `TEXT(${x},"General")`]) add('paths', f)
  // TEXT with the General format: a rule other than the conversion's
  for (const m of ['1', '1.5', '1.2345', '1.23456789', '1.23456789012345', '9.99999', '9.999999', '-1.5', '-1.23456789012345'])
    for (let e = -12; e <= 15; e++) add('general', `TEXT(${lit(m, e)},"General")`)
  for (const [n, d] of [['1', '3'], ['2', '3'], ['-1', '7']] as const) for (let e = -12; e <= 14; e++) add('general', `TEXT(${n}/${lit(d, e)},"General")`)
  // edges of the conversion: three-digit exponents, subnormals, rounding at the 15-digit boundary
  for (const x of ['1E-100', '1.5E+300', '-1E-300', '(1E-300/1E+20)', '9.999999999999999E+19', '99999999999999999999', '(1E+20-1)', '(0.1+0.2)', '(1/3*3)', '(9.99999999999999E-05+1E-19)'])
    add('text', `"a"&${x}`)
  // below the smallest normal number (2.2250738585072E-308): Excel has no subnormals
  for (const x of ['(1E-300*1E-10)', '(3E-308/2)', '(2.2250738585072E-308*1)', 'POWER(1E-160,2)', '(-1E-300/1E+15)']) add('text', `"a"&${x}`)
  add('text', '(1E-300*1E-10)=0')
  // General with a three-digit exponent
  for (const x of ['1E-100', '1.23456789E-100', '1.5E+300', '-1.23456789E+200']) add('general', `TEXT(${x},"General")`)
  // literals and numeric text longer than 15 significant digits
  for (const f of [
    '123456789012345678-123456789012345000',
    '-123456789012345678+123456789012345000',
    '1.99999999999999999=2',
    '1.99999999999999999-1.9999999999999',
    '0.99999999999999999=1',
    '9999999999999999-9999999999999990',
    '"a"&123456789012345678',
    '"a"&0.1234567890123456789',
    '"a"&1.234567890123456789E-10',
    '"a"&12345678901234567890123',
    '"123456789012345678"*1-123456789012345000',
    '"1.99999999999999999"*1=2',
    '"0.99999999999999999"+0=1',
  ])
    add('parse', f)
  // criteria over whole columns of different sheets
  for (const f of [
    "SUMIFS('S2'!A:A,'S3'!B:B,\">0\")",
    "SUMIF('S3'!B:B,\">0\",'S2'!A:A)",
    "COUNTIFS('S3'!B:B,\">0\",'S2'!A:A,\">2\")",
    "SUMIFS('S2'!A:A,'S3'!B:B,\"\")",
    "SUMIFS('S2'!A:A,'S3'!B:B,\"<>1\")",
    "COUNTIFS('S2'!A:A,\"\")",
    "COUNTIFS('S2'!A:A,\"<>3\")",
    "COUNTIFS('S2'!A:B,\"\")",
    "SUMIFS('S2'!A:A,'S3'!B1:B3,\">0\")",
    "SUMIF('S3'!B:B,\"\",'S2'!A1)",
  ])
    add('criteria', f)
  return out
}
