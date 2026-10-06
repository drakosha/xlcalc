/**
 * Excel cases for the math and engineering functions beyond the coverage measurement: domain edges (whole
 * numbers, ranges, digit counts of number systems, limits of trigonometry and factorials), large and
 * negative arguments, complex number syntax and output, CONVERT units and prefixes, functions formula.js
 * lacks (MDETERM, ERF.PRECISE, ERFC.PRECISE). Answers: `docs/excel-cases/matheng.json`
 * (`npx tsx bench/coverage/run.ts probes matheng`).
 */
import type { Probe } from '../probes.js'

/** Functions newer than Excel 2007: stored with the `_xlfn.` prefix. */
const XLFN = new Set([
  'ACOT', 'ARABIC', 'BASE', 'BITAND', 'BITLSHIFT', 'BITOR', 'BITRSHIFT', 'BITXOR', 'COMBINA', 'COT', 'COTH', 'CSC', 'DECIMAL',
  'ERF.PRECISE', 'ERFC.PRECISE', 'IMCOSH', 'IMCOT', 'IMCSC', 'IMCSCH', 'IMSEC', 'IMSECH', 'IMSINH', 'IMTAN', 'ISO.CEILING', 'PERCENTOF',
  'SEC',
])

/** Data of the set's own sheet M: ranges with mixed content, matrices. */
export const SHEETS: Record<string, Record<string, number | string | boolean>> = {
  M: {
    // GCD, LCM, MULTINOMIAL: text number, boolean, text, blank, fractions
    A1: 4, A2: '6',
    B1: 4, B2: true,
    C1: 4, C2: 'abc',
    D1: 4, D3: 6,
    E1: 4.5, E2: 6.9,
    // MDETERM
    G4: 1, H4: 3, I4: 8, J4: 5,
    G5: 1, H5: 3, I5: 6, J5: 1,
    G6: 1, H6: 1, I6: 1, J6: 0,
    G7: 7, H7: 3, I7: 10, J7: 2,
    L1: 3, M1: 6, N1: 1,
    L2: 1, M2: 1, N2: 0,
    L3: 3, M3: 10, N3: 2,
    L5: 1, M5: 2, N5: 3,
    L6: 4, M6: 5, N6: 6,
    L7: 7, M7: 8, N7: 9,
    G9: 1, G10: 3, H10: 4,
    G12: 1, H12: '2', G13: 3, H13: 4,
    G15: 1, H15: true, G16: 3, H16: 4,
    P1: 1, Q1: 2, P2: 3, Q2: 4, P3: 5, Q3: 6,
    L9: 0, M9: 1, N9: 2,
    L10: 1, M10: 0, N10: 3,
    L11: 4, M11: -3, N11: 8,
    ...hilbert('S10', 5),
    // SUMPRODUCT and the SUMX functions: mixed against numbers
    S1: 1, S2: '2', S3: true,
    T1: 5, T2: 6, T3: 7, T4: 8,
    // SERIESSUM coefficients: blank, text number, boolean
    V1: 1, V3: 2,
    W1: 1, W2: '2', W3: 3,
    X1: 1, X2: true,
    // IMSUM, IMPRODUCT over ranges
    Y1: '1+i', Y2: '2', Y3: 3,
    Z1: '1+i', Z3: '2',
    AA1: '1+i', AA2: true,
    AB1: '1+j', AB2: '2+j',
    AC1: '1+i', AC2: 'abc',
  },
}

function hilbert(at: string, n: number): Record<string, number> {
  const col = at.charCodeAt(0) - 65
  const row = Number(at.slice(1))
  const out: Record<string, number> = {}
  for (let i = 0; i < n; i++) for (let j = 0; j < n; j++) out[`${String.fromCharCode(65 + col + j)}${row + i}`] = 1 / (i + j + 1)
  return out
}

/** Complex number syntax: valid and invalid forms. */
const COMPLEX_TEXTS = [
  '3+4i', '3-4i', '3+4j', '-3-4j', '3+4I', '3+4J', '+3+4i', '+3', '-3', '3', '3.', '3.5', '.5', '.5i', '5.i', '4i', '+4i', '-4i',
  'i', '+i', '-i', 'j', '-j', '1+i', '1-i', '1-j', '1e3+2e-1i', '1E3-2E+1i', '1E3i', '1e3', '1e+3', '1e-3i', '2e3e', '3+4', '3i+4',
  '3 + 4i', ' 3+4i', '3+4i ', '3+i4', 'i4', '3+4ii', '3+-4i', '3++4i', '--3', '+-3', '0', '0i', '-0', '-0i', '1e', '1e+', 'e5',
  '1e400', '1e308+1e308i', '0x10', '1,5', '1.2.3', '3+4k', 'i3', 'ii', '+', '-', '3-', '3+', '3+i', '3+j', '4+3.5e2i',
  '1.5e-3-2.5e+2j', '00012', '-.5-.5i', ' ', '1i', '0+0i', '-0-0i', '1+0i', '0+1i', '123456789012345678+1i',
]

/** Arguments for the one-argument complex functions. */
const COMPLEX_ARGS = ['"1+i"', '"-1"', '"i"', '"-i"', '"0"', '"3-4i"', '"-2.5+0.5j"', '"1e-10+i"', '5', '-5.5']

const BIN = ['BIN2DEC', 'BIN2HEX', 'BIN2OCT']

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const count = new Map<string, number>()
  const add = (fn: string, args: string) => {
    const k = (count.get(fn) ?? 0) + 1
    count.set(fn, k)
    out.push({ id: `${fn}#${k}`, group: fn, fn, formula: `${XLFN.has(fn) ? '_xlfn.' : ''}${fn}(${args})` })
  }
  const each = (fn: string, args: string[]) => args.forEach((a) => add(fn, a))

  // ---------- engineering ----------

  // Bessel functions: a grid of x and orders, negative x, large orders, overflow, argument coercion
  const xs = ['0.1', '0.5', '1', '1.5', '2', '2.5', '3', '3.75', '4', '5', '7.5', '8', '8.5', '10', '15', '30', '100']
  for (const fn of ['BESSELI', 'BESSELJ', 'BESSELK', 'BESSELY']) {
    for (const x of xs) for (const n of ['0', '1', '2', '3', '5', '10']) add(fn, `${x},${n}`)
    for (const x of ['-0.5', '-3', '-10']) for (const n of ['0', '1', '2', '3']) add(fn, `${x},${n}`)
    for (const x of ['1', '30']) for (const n of ['20', '50', '100', '1000', '10000', '100000']) add(fn, `${x},${n}`)
    each(fn, ['0,0', '0,1', '0,2', '2,2.9', '2,0.5', '2,-0.5', '2,-1', '"2","2"', '2,TRUE', 'TRUE,2', '2,D!A1', 'D!A1,2', 'D!A4,1', 'D!A2,1', '2,"abc"', '1E10,1', '1E15,0', '700,0', '709,1', '710,0', '713,1', '0.001,10', '0.00001,5', '750,1', 'D!B1:B2,1'])
  }
  for (const fn of ['BESSELI', 'BESSELJ']) each(fn, ['1,1000000', '1,10000000', '5000,4000', '0.5,30'])

  // number systems: digit counts, two's complement ranges, places, inputs as numbers and blanks
  const textIn = ['"0"', '"1"', '""', 'D!A1', 'D!A13', 'TRUE', 'D!A4', 'D!A2', '" 1"', '"1 "', 'D!B1:B2', '0.1', '-1', '"-1"', '"+1"']
  for (const fn of BIN) each(fn, [...textIn, '"101"', '"0111111111"', '"1000000000"', '"1111111111"', '"11111111111"', '"0000000001"', '"2"', '101', '1111111111', '1E+3', '11111111111', '"1e2"'])
  for (const fn of ['BIN2HEX', 'BIN2OCT'])
    each(fn, ['"11111011",2', '"11111011",1', '"11111011",10', '"11111011",11', '"1111111111",3', '"1111111111",11', '"1111111111",-1', '"1000000000",0', '"0",0', '"0",1', '"101","2.9"', '"101",-1', '"101",D!A1', '"101",TRUE', '"101","abc"', '"101",1E10', '"101",10.9', '"101",D!A4', '"101",D!A2'])
  for (const fn of ['DEC2BIN', 'DEC2HEX', 'DEC2OCT'])
    each(fn, ['0', '1', '-1', '"9"', '"1e1"', 'D!A4', 'D!A2', 'D!A1', '" 9 "', 'TRUE', '9,"4"', '9,0', '9,10', '9,11', '-1,3', '-1,11', '-1,-1', '-1,0', '-0.5', '0.999', '"abc"', '9,TRUE', '9,D!A4', '9,1E10', 'D!B1:B2', '1E10'])
  each('DEC2BIN', ['511', '512', '-512', '-513', '511.9', '-512.9', '9,4', '9,3', '255,8', '255,7'])
  each('DEC2HEX', ['549755813887', '549755813888', '-549755813888', '-549755813889', '255,2', '255,1', '-1,2', '1E12', '4294967296', '255.99', '-54'])
  each('DEC2OCT', ['536870911', '536870912', '-536870912', '-536870913', '8,2', '8,1', '-100', '511'])
  const hexIn = ['"1FF"', '"200"', '"FFFFFFFE00"', '"FFFFFFFDFF"', '"FFFFFFFFFF"', '"7FFFFFFFFF"', '"8000000000"', '"ff"', '"Ff"', '"FFFFFFFFFFF"', '"G"', '"1G"', '10', '1E+3', '"0x10"', '"1FFFFFFF"', '"20000000"', '"FFE0000000"', '"FFDFFFFFFF"', '"A5"', '"3DA408B9"']
  for (const fn of ['HEX2BIN', 'HEX2DEC', 'HEX2OCT']) each(fn, [...textIn, ...hexIn])
  for (const fn of ['HEX2BIN', 'HEX2OCT']) each(fn, ['"F",4', '"F",3', '"F",2', '"F",1', '"F",10', '"F",11', '"FFFFFFFFFF",3', '"0",0', '"F","2.9"', '"F",TRUE', '"F",D!A4', '"F",-1'])
  const octIn = ['"777"', '"1000"', '"7777777000"', '"7777776777"', '"7777777777"', '"3777777777"', '"4000000000"', '"77777777777"', '"8"', '17', '"12"', '"0012"']
  for (const fn of ['OCT2BIN', 'OCT2DEC', 'OCT2HEX']) each(fn, [...textIn, ...octIn])
  for (const fn of ['OCT2BIN', 'OCT2HEX']) each(fn, ['"12",3', '"12",2', '"17",1', '"17",2', '"17",10', '"17",11', '"7777777777",2', '"0",0', '"17","2.9"', '"17",TRUE', '"17",-1'])

  // bitwise functions: the 48-bit range, whole numbers only, shifts within 53 bits
  for (const fn of ['BITAND', 'BITOR', 'BITXOR'])
    each(fn, ['0,0', '13,25', '281474976710655,1', '281474976710656,1', '1,281474976710656', 'POWER(2,47)+1,POWER(2,47)+3', '5.0000000001,1', '-1,1', '"7","3"', 'TRUE,TRUE', 'D!A4,1', 'D!A2,3', 'D!B1:B2,1', '1E15,1', '0.1*3*10,3', '2.5-0.5,3', '"abc",1', 'D!A1,5'])
  for (const fn of ['BITLSHIFT', 'BITRSHIFT'])
    each(fn, ['1,0', '1,1', '1,47', '1,48', '1,53', '1,-53', '1,54', '1,-54', '0,60', 'POWER(2,47),1', 'POWER(2,48)-1,0', 'POWER(2,48)-1,1', 'POWER(2,48)-1,-1', '281474976710655,53', '281474976710655,-53', '3,2.9', '3,-2.9', '12,"2"', '12,TRUE', '12,D!A1', '12,D!A4', '5,-2', '20,2', '-1,1', '1.5,1', '281474976710656,0', '1,"abc"'])

  // complex numbers: syntax through IMREAL and IMAGINARY, construction, arithmetic, functions
  for (const t of COMPLEX_TEXTS) for (const fn of ['IMREAL', 'IMAGINARY']) add(fn, JSON.stringify(t))
  each('COMPLEX', ['1,1', '1,-1', '0,-1', '0,0', '-1,0', '1/3,2/3', '1E20,1E-20', '3,4,"I"', '3,4,"J"', '3,4,""', '3,4,D!A1', '3,4,TRUE', '3,4,"k"', '3,4,"ij"', '3,4,1', '0.1+0.2,0', '-0.5,-2.5,"j"', '123456789012345678,1', 'D!A4,1', '"3","4"', '1,1,"j"', '0,1,"j"', '0,-1,"j"', '1E300*10,1', '"abc",1', '3,"abc"'])
  const single = ['IMABS', 'IMARGUMENT', 'IMCONJUGATE', 'IMCOS', 'IMCOSH', 'IMCOT', 'IMCSC', 'IMCSCH', 'IMEXP', 'IMLN', 'IMLOG10', 'IMLOG2', 'IMSEC', 'IMSECH', 'IMSIN', 'IMSINH', 'IMSQRT', 'IMTAN']
  for (const fn of single) each(fn, [...COMPLEX_ARGS, 'D!A4', 'D!B1:B2', '"1+j"'])
  each('IMABS', ['"1e308+1e308i"', '"1e-200+1e-200i"', '"-3-4j"'])
  each('IMARGUMENT', ['"-1-0i"', '"1e308+1e308i"', '"-1-i"', '"-1+i"', '"1e-300i"'])
  each('IMCONJUGATE', ['"3+4j"', '"0-j"'])
  each('IMEXP', ['"710"', '"709+i"', '"1+1e10i"', '"-800"'])
  each('IMCOS', ['"1+800i"', '"1E10"', '"134217728"', '"134217727"', '"1+1000i"'])
  each('IMSIN', ['"1+800i"', '"1E10"', '"134217728"', '"134217727"'])
  each('IMTAN', ['"1+400i"', '"1.5707963267949"', '"134217728"'])
  each('IMCOT', ['"1e-300"', '"1+400i"'])
  each('IMCOSH', ['"710"', '"1+1E10i"'])
  each('IMSINH', ['"710"', '"-1+2j"'])
  each('IMSQRT', ['"1e308+1e308i"', '"-4"', '"4"', '"-4-0i"', '"i"', '"-1e-300"'])
  each('IMLN', ['"1e308+1e308i"', '"-1-i"', '"1e-300"', '"-0"'])
  each('IMLOG10', ['"100"', '"-1"', '"0.001i"'])
  each('IMLOG2', ['"8"', '"-1"', '"0.25j"'])
  each('IMPOWER', ['"2+3i",3', '"i",2', '"-1",0.5', '"0",0', '"0",1', '"0",-1', '"1+i",-1', '"2",0.5', '"1+i",0', '"1+i",100', '"1+i",2000', '"-8",1/3', '"1+i","2"', '"1+i",TRUE', '"1+i",D!A1', '"4",0.5', '"1e200",2', '"1+j",2', '"-2",2', '"-2",3', '"1+i",D!A4', '"abc",2', '"1+i",2.5'])
  each('IMDIV', ['"1+i","1-i"', '"1","0"', '"0","1+i"', '"1+j","1+i"', '"1+j","2"', '"1","1+j"', '"1E200","1E-200"', '"1+i","1E300+1E300i"', '"0","0"', '"1+j","1+j"', '"6","3"', 'D!A4,"1"', '"1",D!B1:B2'])
  each('IMPRODUCT', ['"1+i"', '"1+i","1-i"', '"1+j","1+i"', '"2","3","4"', 'M!Y1:Y3', 'M!Z1:Z3', 'M!AA1:AA2', 'M!AB1:AB2', 'M!AC1:AC2', '"1E200","1E200"', '"i","i","i"', 'D!B1:B3', 'D!K1:K8', 'D!K1:K8,"2"', '"1+j","2"', 'D!A4', '"1+j",M!Z1:Z3'])
  each('IMSUB', ['"1+j","1+i"', '"1+j","1"', '"1","1"', '"1E308","-1E308"', '"1+j","1+j"', '"3","4j"', 'D!A4,"1"', 'D!B1:B2,"1"'])
  each('IMSUM', ['"1","2+i","3-2i"', '"1+j","2+j"', '"1+j","1+i"', 'M!Y1:Y3', 'M!Z1:Z3', 'M!AA1:AA2', 'M!AB1:AB2', 'M!AC1:AC2', 'D!B1:B8', 'D!D1:D8', 'D!K1:K8', 'D!K1:K8,"2"', '"1E308","1E308"', '"0.1","0.2"', 'D!A4', '"1+i",M!AB1:AB2', '"i","-i"'])

  // CONVERT: every unit and alias against its quantity's base unit, prefixes, temperatures
  const units: [string, string[]][] = [
    ['g', ['g', 'sg', 'lbm', 'u', 'ozm', 'grain', 'cwt', 'shweight', 'uk_cwt', 'lcwt', 'hweight', 'stone', 'ton', 'uk_ton', 'LTON', 'brton']],
    ['m', ['m', 'mi', 'Nmi', 'in', 'ft', 'yd', 'ang', 'ell', 'ly', 'parsec', 'pc', 'Picapt', 'Pica', 'pica', 'survey_mi']],
    ['sec', ['yr', 'day', 'd', 'hr', 'mn', 'min', 'sec', 's']],
    ['Pa', ['Pa', 'p', 'atm', 'at', 'mmHg', 'psi', 'Torr']],
    ['N', ['N', 'dyn', 'dy', 'lbf', 'pond']],
    ['J', ['J', 'e', 'c', 'cal', 'eV', 'ev', 'HPh', 'hh', 'Wh', 'wh', 'flb', 'BTU', 'btu']],
    ['W', ['HP', 'h', 'PS', 'W', 'w']],
    ['T', ['T', 'ga']],
    ['m3', ['tsp', 'tspm', 'tbs', 'oz', 'cup', 'pt', 'us_pt', 'uk_pt', 'qt', 'uk_qt', 'gal', 'uk_gal', 'l', 'L', 'lt', 'ang3', 'ang^3', 'barrel', 'bushel', 'ft3', 'ft^3', 'in3', 'in^3', 'ly3', 'ly^3', 'm3', 'm^3', 'mi3', 'mi^3', 'yd3', 'yd^3', 'Nmi3', 'Nmi^3', 'Picapt3', 'Picapt^3', 'Pica3', 'Pica^3', 'GRT', 'regton', 'MTON']],
    ['m2', ['uk_acre', 'us_acre', 'ang2', 'ang^2', 'ar', 'ft2', 'ft^2', 'ha', 'in2', 'in^2', 'ly2', 'ly^2', 'm2', 'm^2', 'Morgen', 'mi2', 'mi^2', 'Nmi2', 'Nmi^2', 'Picapt2', 'Pica2', 'Pica^2', 'Picapt^2', 'yd2', 'yd^2']],
    ['bit', ['bit', 'byte']],
    ['m/s', ['admkn', 'kn', 'm/h', 'm/hr', 'm/s', 'm/sec', 'mph']],
  ]
  for (const [base, us] of units) for (const u of us) add('CONVERT', `1,"${u}","${base}"`)
  // a decimal prefix on every unit: which units take prefixes
  for (const [base, us] of units) for (const u of us) add('CONVERT', `1,"k${u}","${base}"`)
  for (const p of ['Y', 'Z', 'E', 'P', 'T', 'G', 'M', 'k', 'h', 'da', 'e', 'd', 'c', 'm', 'u', 'n', 'p', 'f', 'a', 'z', 'y']) add('CONVERT', `1,"${p}m","m"`)
  for (const p of ['Yi', 'Zi', 'Ei', 'Pi', 'Ti', 'Gi', 'Mi', 'ki']) each('CONVERT', [`1,"${p}byte","bit"`, `1,"${p}bit","bit"`])
  each('CONVERT', [
    '1,"kim","m"', '1,"Kbyte","byte"', '1,"cm2","m2"', '1,"km2","m2"', '1,"km^2","m2"', '1,"mm3","m3"', '1,"cm3","m3"', '1,"dm3","l"', '1,"km3","m3"',
    '1,"ml","l"', '1,"mL","l"', '1,"dl","l"', '1,"hl","l"', '1,"ms","s"', '1,"msec","s"', '1,"us","s"', '1,"kPa","Pa"', '1,"hPa","Pa"', '1,"matm","Pa"',
    '1,"kmmHg","Pa"', '1,"kWh","J"', '1,"MeV","J"', '1,"keV","eV"', '1,"kcal","J"', '1,"km/h","m/s"', '1,"km/hr","m/s"', '1,"mT","T"', '1,"mga","T"',
    '1,"mK","K"', '1,"kK","K"', '1,"mC","C"', '1,"kC","K"', '1,"kF","F"', '1,"kRank","Rank"', '1,"kpc","m"', '1,"kparsec","m"', '1,"kly","m"', '1,"kang","m"',
    '1,"mg","g"', '1,"kg","lbm"', '1,"ug","g"', '1,"Mg","ton"', '1,"dam","m"', '1,"em","m"', '1,"hm","m"', '1,"ec","J"', '1,"ha","m2"', '1,"har","m2"',
  ])
  each('CONVERT', [
    '100,"C","F"', '212,"F","C"', '0,"K","C"', '0,"Rank","K"', '100,"Reau","C"', '100,"C","Rank"', '1,"cel","fah"', '1,"kel","K"', '-40,"C","F"',
    '300,"K","Reau"', '50,"F","Rank"', '1,"Reau","Rank"', '0,"C","K"', '-273.15,"C","K"', '-500,"C","K"', '1,"C","C"', '1,"Reau","Reau"',
    '1,"m","m"', '1,"m","sec"', '1,"LBM","kg"', '1," m","m"', '1,"m ","m"', '1,"","m"', '"2","m","ft"', 'TRUE,"m","ft"', 'D!A4,"m","ft"', '1,TRUE,"m"',
    '1,"m",D!A1', 'D!A1,"m","ft"', '1E308,"km","m"', '1,"yd","ft"', '1,"mi","Nmi"', '1,"gal","uk_gal"', '1,"ft2","in2"', '1,"BTU","cal"',
    '1,"hh","HPh"', '1,"lbm","ozm"', '1,"day","yr"', '1,"psi","mmHg"', '1,"atm","Torr"', '1,"mph","kn"', '1,"admkn","kn"', '1,"byte","bit"',
    '1,"Wh","kJ"', '1,"HP","PS"', '1,"ly","parsec"', '1,"survey_mi","mi"', '1,"us_acre","uk_acre"', '1,"pica","Pica"', '1,"cup","oz"', '1,"tbs","tsp"',
    '6,"tsp","tbs"', '1,"kn","m/h"', '1,"Picapt3","Pica^3"', '1,"ang","Pica"', '1,"g","u"', '1,"eV","e"', '1,"hr","mn"', '1,"MTON","GRT"', '1,"lbf","dyn"',
  ])

  // DELTA, GESTEP, ERF and ERFC
  each('DELTA', ['1,1', '1,2', '0.1+0.2,0.3', '1', '0', '"a"', 'TRUE', 'D!A4', 'D!A2,2', '1E-20,0', '"1","1"', '1,D!A1', 'D!B1:B2,1'])
  each('GESTEP', ['0.1+0.2,0.3', '5', '-5', '0,0', 'TRUE', 'D!A4', '"5","4"', '1E-20', '-1E-20', 'D!A1,D!A1', '4,"abc"'])
  each('ERF', ['0.1', '0.5', '1', '2', '3', '4', '5', '6', '-1', '1,2', '2,1', '-1,1', '0,0', '1,1', '5,6', '1,D!A1', 'D!A1,1', '1E-10', '27', 'TRUE', 'D!A4', '"1"', '1,TRUE', '-30,30', '0.3,0.30001'])
  each('ERF.PRECISE', ['0.5', '1', '-1', '3', '6', '1E-10', 'TRUE', 'D!A4', '"1"', 'D!A1', '"abc"'])
  each('ERFC', ['0.1', '0.5', '1', '2', '3', '4', '5', '6', '8', '10', '20', '26', '27', '-1', '-5', '-30', '1E-10', 'TRUE', 'D!A4', '"2"'])
  each('ERFC.PRECISE', ['1', '5', '-1', '27', '0', '10', 'TRUE', 'D!A4', '"2"', 'D!A1'])

  // ---------- math ----------

  const trig = ['134217727', '134217728', '-134217727', '-134217728', '134217727.9', '1E8', 'PI()', 'PI()/2', '-PI()', '0', '1E-300', 'TRUE', 'D!A4', '"1"', '1E300']
  for (const fn of ['SIN', 'COS', 'TAN', 'SEC', 'CSC', 'COT']) each(fn, trig)
  each('ACOT', ['1', '-1', '0', '1E-300', '-1E10', '1E300', '5', 'TRUE', 'D!A4', '-1E-300', '"2"'])
  each('COTH', ['1E-300', '20', '-20', '0', '710', '-710', '1E-10', 'TRUE', 'D!A4', '1E300', '0.5'])
  each('TANH', ['20', '-20', '1E-10', '710', '1E300', '-1E300', 'D!A4', '0.5'])
  each('ATAN2', ['0,0', '0,1', '0,-1', '-1,0', '1,0', '1E300,1E300', '-1,-1E-10', '"1","1"', '1E-300,1E-300', 'TRUE,FALSE', 'D!A4,1', '-1,1E-300'])
  each('BASE', ['0,2', '255,16', '255,36', '35,36', '36,36', '7,2,0', '7,2,1', '7,2,255', '7,2,256', '7,2,-1', '7,1', '7,37', '7,36.9', 'POWER(2,53)-1,2', 'POWER(2,53),2', 'POWER(2,53)-1,36', '1E15,36', '-1,2', '7.9,2', '7,2,2.9', '7,"2"', '7,2,TRUE', 'D!A4,2', '"7",2', '7,2,D!A1', '1E16,2'])
  each('COMBIN', ['0,0', '5,5', '5,6', '5,-1', '-1,0', '1000,2', '1000,500', '1030,515', '60,30', '170,85', '1E9,2', '2147483647,2', '2147483648,2', '4294967295,1', '4294967296,1', '4294967296,0', '9999999999,1', '1E15,1', '8.9,2.9', 'TRUE,TRUE', 'D!A4,1', '"8","2"', '50,25', '1029,514', '10000,5000'])
  each('COMBINA', ['4,3', '0,0', '0,1', '1,0', '4,0', '10,3', '1000,2', '2147483647,1', '2147483648,1', '4294967296,1', '1E9,1', '-1,1', '4,-1', '2.9,2.9', 'D!A4,1', '"4","3"', '515,515', '100,100'])
  each('DECIMAL', ['"FF",16', '"ff",16', '"Ff",16', '"zz",36', '"ZZ",36', '"10",2', '"12",2', '"",2', '"0",2', '"1"&REPT("0",52),2', 'REPT("1",53),2', 'REPT("1",54),2', 'REPT("F",14),16', 'REPT("Z",255),36', 'REPT("0",256),2', '" FF",16', '"FF ",16', '"-FF",16', '"FF",1', '"FF",37', '"FF",16.9', '"FF","16"', '12,10', 'TRUE,36', 'D!A4,36', '"1.5",10', '"1E3",16', 'REPT("1",255),2', '"FF",TRUE'])
  each('FACT', ['0', '1', '5', '20', '21', '25', '100', '170', '171', '-1', '-0.5', '0.99', '169.99', 'TRUE', 'D!A4', '"5"', '-1E-300'])
  each('FACTDOUBLE', ['-1', '-2', '-0.5', '0', '1', '2', '10', '11', '100', '299', '300', '301', '302', '5.9', 'TRUE', 'D!A4', '"6"'])
  each('GCD', ['0', '0,0', '12', '12,0', '12,18,24', '5.9,3.9', '6.9,4.1', 'POWER(2,53)-1,3', 'POWER(2,53),2', '1E15,5', '-1,2', 'M!A1:A2', 'M!B1:B2', 'M!C1:C2', 'M!D1:D3', 'M!E1:E2', '"6",4', 'D!A2,4', 'D!B1:B8', 'D!A4', 'D!A13,4', 'D!B1:B6,6'])
  each('LCM', ['0', '5', '4,6', '4,0', 'M!A1:A2', 'M!B1:B2', 'M!C1:C2', 'M!D1:D3', 'M!E1:E2', 'POWER(2,52),3', '1E10,1E10+1', '9007199254740991,2', '4.9,6.9', '-4,6', 'D!A2,3', 'D!A13,3', '1,2,3,4,5,6,7,8,9,10'])
  each('MULTINOMIAL', ['0', '1', '1,1,1', '10,10', '170,0', '171', '100,70', '100,71', 'M!A1:A2', 'M!D1:D3', '2.9,3.9', '-1,2', '"2",3', 'TRUE,1', 'M!B1:B2', 'M!C1:C2', 'D!A13,1'])
  each('ISO.CEILING', ['4.3', '-4.3', '4.3,2', '-4.3,-2', '4.3,0', '0,2', '2.5,0.1', '1E20,3', '-2.5,2.5', '0.3,0.1', 'TRUE,1', '"4.3"'])
  each('MOD', ['3,0', '0,0', '0.3,0.1', '1,0.1', '1.1,0.1', '10,0.2', '-1,0.1', '0.7,0.1', '5.5,1E-10', '1E15,7', '1E16,7', '1E20,7', 'POWER(2,53),3', '134217728*3,2', '1E10,3', '-1E10,3', '1E-300,1', '-1E-300,1', '2.5,1E300', '-2.5,1E300', '-0.75,0.25', '7,-0.5', 'POWER(2,27),1', 'POWER(2,27)+0.5,1', '1E15+0.5,1', 'D!A4,2'])
  each('MROUND', ['2.5,1', '-2.5,-1', '1.5,1', '0.5,1', '6.05,0.1', '1.05,0.1', '0.15,0.1', '0.25,0.1', '5,0', '0,-1', '0,5', '1E20,3', '7,1E20', '3.5,7', '-3.5,-7', 'D!A4,1', '4.444,0.002', '-6.05,-0.1', '1.005,0.01', '"10","3"', '10,TRUE'])
  each('QUOTIENT', ['-5,2', '5,-2', '-5,-2', '0.3,0.1', '1,0.1', '1E20,3', '7,0', '0,0', '1E-300,1E300', '1E300,1E-300', 'D!A4,1', '"7","2"', '1,3', '-1,3', '6,0.2', '2.9999999999999996,1', '7,TRUE'])
  for (const n of [1, 4, 5, 9, 14, 40, 45, 46, 49, 90, 95, 99, 400, 450, 490, 495, 499, 900, 950, 990, 995, 999, 1999, 3999, 444, 1444, 3449, 1495, 48, 98, 498, 998])
    for (let f = 0; f <= 4; f++) add('ROMAN', `${n},${f}`)
  each('ROMAN', ['4000', '-1', '0', '3999.9', '1,5', '1,-1', '499,TRUE', '499,FALSE', '499,"2"', '499,D!A1', '499,1.9', '"12"', 'TRUE', 'D!A4,0', '12,"abc"', '499,D!A4', '0.5', '1999'])
  each('ARABIC', ['"I"', '"IV"', '"MCMXCIX"', '"mcmxcix"', '"MMMCMXCIX"', '"MMMM"', '"IIII"', '"IC"', '"VX"', '"XM"', '"LL"', '"DM"', '"IIIIIIIIII"', '"-MCM"', '"--I"', '"-"', '" MCM"', '"MCM "', '"M CM"', '"0"', '"A"', 'REPT("M",255)', 'REPT("M",256)', '"XIIII"', '"CCCCC"', '"IM"', '"VV"', '"XLX"', '"iV"', 'D!A13', '" "', '"MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM"'])
  each('SERIESSUM', ['2,1,1,M!W1:W3', '2,1,1,M!V1:V3', '2,1,1,M!X1:X2', '2,1,1,3', '2,0,1,D!AR1:AR4', '0,1,1,D!AR1:AR4', '0,0,1,D!AR1:AR4', '-2,0.5,1,D!AR1:AR4', '2,1.5,1,D!AR1:AR4', '2,1,1,D!E1:E8', '2,1,1,D!K1:K8', '2,1,1,"3"', '2,1,1,TRUE', '1E200,2,1,D!AR1:AR4', '2,-1,-1,D!AR1:AR4', 'D!A4,1,1,D!AR1:AR4', '2,1,1,M!C1:C2', '0,-1,1,D!AR1:AR4'])
  each('SQRTPI', ['0', '1', '1E300', '1E308', '-1', '"4"', 'D!A4', '1E-300'])
  each('SUMPRODUCT', ['D!B1:B8', 'D!B1:B8,D!C1:C8,D!B1:B8', '3,4', '"3",4', 'TRUE,4', 'D!B1,D!C1', 'D!P1:Q2,D!B1:B4', 'M!S1:S4,M!T1:T4', 'D!B:B,D!C:C', 'D!I13:I17,D!E2:E6', 'D!E2:E6,D!I13:I17', 'D!K1:K8', 'D!A7', 'D!A2,D!A2', 'D!A4,1', '"abc",1', 'D!P1:Q2,D!R1:S2,D!P1:Q2', 'D!B1:C8'])
  each('SUMSQ', ['D!B:B', '1E200,1E200', '"3"', 'TRUE', 'D!A4', 'D!A2', 'D!A13', 'D!B1:C8'])
  for (const fn of ['SUMX2MY2', 'SUMX2PY2', 'SUMXMY2'])
    each(fn, ['3,4', 'M!S1:S4,M!T1:T4', 'D!B1:B4,D!C1:C4', 'D!P1:Q2,D!B1:B4', 'D!B:B,D!C:C', 'D!K1:K8,D!K1:K8', 'D!I13:I17,D!E2:E6', 'D!E2:E6,D!I13:I17', '"3",4', 'TRUE,4', 'D!P1:Q2,D!R1:S2', 'D!B1:B8,D!C1:C4'])
  each('PERCENTOF', ['D!B1:B2,D!B1:B8', '5,10', 'D!B1,D!B1:B8', 'D!D1:D8,D!B1:B8', '0,D!K1:K8', 'D!B1:B8,D!K1:K8', 'D!E1:E8,D!B1:B8', '"5",10', 'TRUE,10'])
  each('MDETERM', ['D!P1:Q2', 'M!G4:J7', 'M!L1:N3', 'M!L5:N7', 'M!G9:H10', 'M!G12:H13', 'M!G15:H16', 'M!P1:Q3', '5', '"5"', 'D!B1', 'D!A1', 'D!A7', 'M!S10:W14', 'M!L9:N11', 'TRUE', 'D!B1:B8'])
  return out
}
