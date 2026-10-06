/**
 * Excel cases of the text, lookup, database, logical and information functions beyond the coverage
 * measurement: wildcards, case and Cyrillic, approximate search over sorted data, errors on input, number
 * text, database criteria. Sheets: T texts, L lookup data, B a database with criteria blocks.
 */
import type { Probe } from '../probes.js'

type Cell = number | string | boolean

const column = (col: string, values: (Cell | null)[], from = 1): Record<string, Cell> =>
  Object.fromEntries(values.flatMap((v, i) => (v === null ? [] : [[`${col}${from + i}`, v]])))

const DB: Cell[][] = [
  ['Tree', 'Height', 'Age', 'Yield', 'Profit'],
  ['Apple', 18, 20, 14, 105],
  ['Pear', 12, 12, 10, 96],
  ['Cherry', 13, 14, 9, 105],
  ['Apple', 14, 15, 10, 75],
  ['Pear', 9, 8, 8, 76.8],
  ['Apple', 8, 9, 6, 45],
  ['Applesauce', 10, '12', true, 50],
  ['pineapple', 7, 5, '', 30],
]

/** Criteria blocks of sheet B by column: header, then conditions (null for a blank cell). */
const CRITERIA: Record<string, (Cell | null)[]> = {
  G: ['Tree', 'App'],
  H: ['Tree', '=Apple'],
  I: ['Tree', '=App'],
  J: ['Tree', '*apple'],
  K: ['Tree', '<>Apple'],
  L: ['Tree', '?ear'],
  M: ['Tree', 'Pear', 'Cherry'],
  R: ['Age', 12],
  S: ['Age', '12'],
  T: ['Yield', true],
  U: ['Yield', '='],
  V: ['Yield', '<>'],
  W: ['Weight', 'x'],
  X: ['Weight', null],
  Y: ['Tree'],
  Z: ['Tree', null, 'Pear'],
  AA: ['tree', 'Pear'],
  AB: ['Profit', '>=96'],
  AC: ['Tree', 'PEAR'],
  AD: [null, 'Pear'],
  AE: ['Tree', '=Cherry'],
  AF: ['Tree', '=pineapple'],
  AG: [1, null],
  AH: [1, '>3'],
}

function database(): Record<string, Cell> {
  const out: Record<string, Cell> = {}
  DB.forEach((row, r) => row.forEach((v, c) => v !== '' && (out[`${'ABCDE'[c]}${r + 1}`] = v)))
  for (const [col, vs] of Object.entries(CRITERIA)) Object.assign(out, column(col, vs))
  // two-column blocks: Tree and Height, Height twice
  Object.assign(out, { N1: 'Tree', O1: 'Height', N2: 'Apple', O2: '>10', P1: 'Height', Q1: 'Height', P2: '>10', Q2: '<15' })
  return out
}

export const SHEETS: Record<string, Record<string, Cell>> = {
  T: {
    ...column('A', ['-', '+']),
    ...column('B', ['b', 'c', 'a', 'ab']),
    ...column('C', ['1\t2', '1 2', ' 12\t', '日本語', 'ｱｲｳ', 'ＡＢＣ１２３', 'abc', 'x\ny']),
  },
  L: {
    ...column('A', [0, 1, 2, 2, 2, 5, 8, 13]),
    ...column('B', ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h']),
    ...column('D', ['apple', 'banana', 'Cherry', 'date', 'fig', 'grape']),
    ...column('E', [1, 2, 3, 4, 5, 6]),
    ...column('G', [1, 3, 5, 'a', 'c', 'e', false, true]),
    ...column('I', [1, null, 3, null, 5, null, 7]),
    ...column('K', [5, 1, 9, 3, 7, 2, 8]),
    ...column('M', [3, 'x', 1, true, 5, 'a', 2]),
    ...column('O', [50, 40, 30, 20, 10, 0]),
    ...column('Q', ['Ёж', 'еж', 'Йод', 'иод']),
    ...column('S', ['a*c', 'a?c', 'a~c', 'abc']),
  },
  B: database(),
}

/** Functions written with the `_xlfn.` prefix in a workbook (newer than Excel 2007). */
const XLFN = /\b(IFNA|SWITCH|XLOOKUP|XMATCH|NUMBERVALUE|UNICHAR|TEXTJOIN|TEXTBEFORE|TEXTAFTER|VALUETOTEXT|ARRAYTOTEXT|DBCS)\(/g

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const add = (fn: string, ...formulas: string[]) => {
    for (const f of formulas) out.push({ id: `${fn}#${out.filter((p) => p.fn === fn).length + 1}`, group: fn, fn, formula: f.replace(XLFN, '_xlfn.$1(') })
  }
  const q = (s: string) => `"${s.replace(/"/g, '""')}"`

  // ---------- text ----------
  add('MID', 'MID("abc",32767,1)', 'MID("abc",32768,1)', 'MID("abc",2147483647,1)', 'MID("abc",2147483648,1)', 'MID("abc",1,10000000000)', 'MID("abc",0,1)', 'MID("abc",1,-1)', 'MID("abc",4,1)', 'MID("abc",3,0)', 'MID("Ёлка",2,2)')
  add('REPLACE', 'REPLACE("abc",4,0,"x")', 'REPLACE("abc",10,1,"x")', 'REPLACE("abc",0,1,"x")', 'REPLACE("abc",1,-1,"x")', 'REPLACE("abc",45688,1,"x")', 'REPLACE("abc",2147483648,1,"x")', 'REPLACE("abc",2,10000000000,"x")', 'REPLACE("Привет",2,2,"ЁЙ")')
  add('SUBSTITUTE', 'SUBSTITUTE("aaa","aa","x",2)', 'SUBSTITUTE("aaaa","aa","x")', 'SUBSTITUTE("aaaa","aa","x",2)', 'SUBSTITUTE("abc","","x")', 'SUBSTITUTE("abc","b","x",0)', 'SUBSTITUTE("abc","B","x")', 'SUBSTITUTE("aXa","a","",2)', 'SUBSTITUTE("ёлка Ёлка","ё","е")', 'SUBSTITUTE("abcb","b","x",1.9)', 'SUBSTITUTE("abc","b","x",-1)', 'SUBSTITUTE("abc","b","x",5)')
  add(
    'SEARCH',
    'SEARCH("","abc",3)',
    'SEARCH("","abc",4)',
    'SEARCH("","",1)',
    'SEARCH("a","",1)',
    'SEARCH("b","abc",0)',
    'SEARCH("*","abc")',
    'SEARCH("?","abc",3)',
    'SEARCH("c*","abc")',
    'SEARCH("~*","a*b")',
    'SEARCH("~","a~b")',
    'SEARCH("~~","a~~b")',
    'SEARCH("~a","x~a")',
    'SEARCH("Й","ий")',
    'SEARCH("Ё","еёЁ")',
    'SEARCH("*b","abc")',
    'SEARCH("b*","abc")',
    'SEARCH("?b","abc")',
    'SEARCH("x?","abx")',
    'SEARCH("a?","a")',
    'SEARCH("~?","a?b",3)',
    'SEARCH("B?D","abcde")',
    'SEARCH("мИр","Привет МИР")',
    'SEARCH("a*c","xxabxxcxx")',
    'SEARCH(".","a.b")',
    'SEARCH("(","a(b")',
  )
  add('PROPER', 'PROPER("don\'t stop")', 'PROPER("o\'neil mc-donald")', 'PROPER("ÉCOLE élève")', 'PROPER("hello_world")', 'PROPER("a.b,c")', 'PROPER("мир-труд май")', 'PROPER("a1b2c3")', 'PROPER("straße")', 'PROPER(T!C8)', 'PROPER("ёЛКА йОД")')
  add('UPPER', 'UPPER("ﬁx")', 'UPPER("straße")', 'UPPER("ёлка йод")', 'UPPER("éàü")', 'UPPER(T!C8)')
  add('DOLLAR', 'DOLLAR(-0.001)', 'DOLLAR(1E+15,2)', 'DOLLAR(0.5,0)', 'DOLLAR(-0.5,0)', 'DOLLAR(1234.5,-5)', 'DOLLAR(1.005,2)', 'DOLLAR(1,100)', 'DOLLAR(1,101)', 'DOLLAR(1,127)', 'DOLLAR(1,128)', 'DOLLAR(-1234567.891,1)', 'DOLLAR(0.125,2)')
  add(
    'NUMBERVALUE',
    ...[
      '"1.000,5",","',
      '"1,000.5"',
      '"1 000,5",","',
      '"12%%"',
      '"%5"',
      '"1e3"',
      '"1E+3"',
      '"-5"',
      '"5-"',
      '"(5)"',
      '"$5"',
      '"1,5"',
      '"1.5.5"',
      '"1.5,5"',
      '" 1 2 3 "',
      'T!C1',
      'T!C2',
      '"1;5",";"',
      '"1.5",".."',
      '"1.5","","."',
      '"1.5",".","."',
      '"1,5",",",","',
      '"1,000.5","."',
      '"1.000.000",","',
      '".5"',
      '"5."',
      '"."',
      '"-"',
      '"+5"',
      '"--5"',
      '"1e"',
      '"1,2,3"',
      '"a"',
      '"1 1/2"',
      '"2025-01-31"',
      '"5 %"',
      '"1,5%",","',
      '"12,5",",","."',
    ].map((a) => `NUMBERVALUE(${a})`),
  )
  add(
    'VALUE',
    ...[
      '1,0000',
      '1,000,0',
      '10,00',
      '1,234,567.89',
      ',123',
      '123,',
      '$-5',
      '-$5',
      '($5)',
      '(5)',
      '-(5)',
      '(-5)',
      '5-',
      '+5',
      '5%',
      '%5',
      '5 %',
      '(5%)',
      '1e3',
      '1E+3',
      '1.5e-2',
      '.5',
      '5.',
      '1 1/2',
      '-1 1/2',
      '0 1/2',
      '1 3/0',
      '1 1/2%',
      '1/2/2025',
      '2025-01-31 18:30',
      '12:30 PM',
      '25:00',
      '24:00',
      '1:30:00.5',
      'Jan 31, 2025',
      '31-Jan-2025',
      'TRUE',
      '$',
      '$ 5',
      '5$',
      '1 000',
      '12abc',
      '0x10',
      '1.2.3',
      '$1,000.50',
      '-.5',
      '1e400',
      '$1e3',
      '1,000e3',
      '( 5 )',
      '- 5',
      '1.5%%',
      '00012',
      '-0',
    ].map((a) => `VALUE(${q(a)})`),
    'VALUE(T!C3)',
    'VALUE(T!C2)',
  )
  add('TEXTJOIN', 'TEXTJOIN(T!A1:A2,TRUE,"a","b","c","d")', 'TEXTJOIN(",",TRUE,"","")', 'TEXTJOIN(",",FALSE,D!K1:K3)', 'TEXTJOIN(",",TRUE,D!K1:K3)', 'TEXTJOIN(",",TRUE,D!P1:Q2)', 'TEXTJOIN(",",FALSE,D!A13,"x")', 'TEXTJOIN(",",TRUE,REPT("a",20000),REPT("b",20000))', 'TEXTJOIN(D!A8,TRUE,"a","b")', 'TEXTJOIN(D!K1:K2,FALSE,"a","b","c")', 'TEXTJOIN(T!A1:A2,TRUE,"a")')
  add('UNICHAR', 'UNICHAR(55296)', 'UNICHAR(57343)', 'UNICHAR(65535)', 'UNICHAR(1114111)', 'UNICHAR(1114112)', 'UNICHAR(128512)', 'UNICHAR(65534)', 'UNICHAR(160)')
  add(
    'TEXTBEFORE',
    ...[
      '"a-b-c","-"',
      '"a-b-c","-",2',
      '"a-b-c","-",-1',
      '"a-b-c","-",3',
      '"a-b-c","-",3,0,1',
      '"a-b-c","-",-3,0,1',
      '"a-b-c","-",-4,0,1',
      '"a-b-c","-",0',
      '"a-b-c","-",6',
      '"a-b-c","-",5',
      '"a-b-c","x"',
      '"a-b-c","x",1,0,0,"none"',
      '"aXb","x"',
      '"aXb","x",1,1',
      '"abc",""',
      '"abc","",2',
      '"abc","",-1',
      '"abc","",-2',
      '"aaa","aa",2',
      '"aaaa","aa",2',
      '"abc",T!B1:B2',
      '"xabc",T!B3:B4',
      '"xabc",T!B3:B4,-1',
      '"a-b-c","-",1,2',
      '"a-b-c","-",1,0,2',
      '"","x"',
      '"","x",1,0,1',
      '12.5,"."',
      '"a-b-c","-",1.9',
      '"a-b-c",D!A1',
      '"a-b-c","-",1,0,0,D!A8',
      'D!A8,"-"',
      '"a-b-c",D!A7',
      '"ЁЛКА-ёлка","ё",1,1',
      '"Ёлка","ё"',
    ].map((a) => `TEXTBEFORE(${a})`),
  )
  add(
    'TEXTAFTER',
    ...[
      '"a-b-c","-"',
      '"a-b-c","-",2',
      '"a-b-c","-",-1',
      '"a-b-c","-",-3,0,1',
      '"a-b-c","-",3,0,1',
      '"a-b-c","-",4,0,1',
      '"a-b-c","-",-4,0,1',
      '"abc",""',
      '"abc","",2',
      '"abc","",-1',
      '"abc","",-2',
      '"xabc",T!B3:B4',
      '"xabc",T!B3:B4,-1',
      '"aXbXc","x",-1,1',
      '"a-b-c","x",1,0,0,0',
      '"aaaa","aa",-1',
      '"aaa","aa",-1',
      '"a-b-c","-",-2',
    ].map((a) => `TEXTAFTER(${a})`),
  )
  add(
    'VALUETOTEXT',
    ...['1.5', '"abc",1', '"a""b",1', 'TRUE,1', 'NA()', '1/0,1', 'D!A1', 'D!A1,1', 'D!A2,1', '1E+21', 'D!A6', '"abc",2', '"abc",-1', '"abc",TRUE', 'D!A13,1', 'D!A1:A1', '0.1+0.2,1'].map((a) => `VALUETOTEXT(${a})`),
  )
  add(
    'ARRAYTOTEXT',
    ...['D!P1:Q2', 'D!P1:Q2,1', 'D!D1:D8', 'D!D1:D8,1', 'D!E1:E8', 'D!E1:E8,1', 'D!K1:K3', 'D!K1:K3,1', '5', '"a",1', 'D!A1:A2,2', 'D!M1:M3,1', 'NA()', 'D!L8:N9,1'].map((a) => `ARRAYTOTEXT(${a})`),
  )

  // ---------- lookup and reference ----------
  add(
    'VLOOKUP',
    'VLOOKUP(D!A1,L!A1:B8,2)',
    'VLOOKUP(D!A1,L!A1:B8,2,FALSE)',
    'VLOOKUP(2,L!A1:B8,2)',
    'VLOOKUP(2,L!A1:B8,2,FALSE)',
    'VLOOKUP(-1,L!A1:B8,2)',
    'VLOOKUP(100,L!A1:B8,2)',
    'VLOOKUP(2.5,L!A1:B8,2)',
    'VLOOKUP(5,L!A1:B8,3)',
    'VLOOKUP(99,L!A1:B8,3,FALSE)',
    'VLOOKUP(5,L!A1:B8,2.9)',
    'VLOOKUP(5,L!A1:B8,"2")',
    'VLOOKUP(5,L!A1:B8,"abc")',
    'VLOOKUP(5,L!A1:B8,NA())',
    'VLOOKUP(5,L!A1:B8,1/0)',
    'VLOOKUP(NA(),L!A1:B8,1/0)',
    'VLOOKUP(5,L!A1:B8,NA(),1/0)',
    'VLOOKUP(5,L!A1:B8,2,"TRUE")',
    'VLOOKUP(5,L!A1:B8,2,2)',
    'VLOOKUP(8,L!A:B,2)',
    'VLOOKUP(8,L!A:B,2,FALSE)',
    'VLOOKUP("c",L!D1:E6,2)',
    'VLOOKUP("CHERRY",L!D1:E6,2,FALSE)',
    'VLOOKUP("b*",L!D1:E6,2,FALSE)',
    'VLOOKUP("b*",L!D1:E6,2)',
    'VLOOKUP("zzz",L!D1:E6,2)',
    'VLOOKUP("A",L!D1:E6,2)',
    'VLOOKUP("a~*c",L!S1:S4,1,FALSE)',
    'VLOOKUP("a~?c",L!S1:S4,1,FALSE)',
    'VLOOKUP("a~c",L!S1:S4,1,FALSE)',
    'VLOOKUP("a?c",L!S1:S4,1,FALSE)',
    'VLOOKUP(5,L!A1:B8,-1)',
    'VLOOKUP(TRUE,L!G1:G8,1)',
  )
  add('HLOOKUP', 'HLOOKUP(100,D!L8:Q10,2)', 'HLOOKUP(5,D!L8:Q10,4)', 'HLOOKUP(5,D!L8:Q10,1.5)', 'HLOOKUP("FIVE",D!L9:Q10,2,FALSE)', 'HLOOKUP("f*",D!L9:Q10,2,FALSE)', 'HLOOKUP(5,D!L8:Q10,NA())', 'HLOOKUP(5,D!8:10,2)', 'HLOOKUP(6,D!L8:Q10,3,TRUE)', 'HLOOKUP(0.5,D!L8:Q10,3)', 'HLOOKUP(D!A8,D!L8:Q10,3)')
  add(
    'MATCH',
    'MATCH(5,D!L1:L6,0.5)',
    'MATCH(5,D!L1:L6,-0.5)',
    'MATCH(4,D!L1:L6,1.5)',
    'MATCH(5,D!L1:L6,"0")',
    'MATCH(5,D!L1:L6,"abc")',
    'MATCH(5,D!L1:N6,0)',
    'MATCH(5,D!L8:Q8,0)',
    'MATCH("b",L!D1:D6,1)',
    'MATCH(D!A1,L!A1:A8,0)',
    'MATCH(D!A1,L!A1:A8,1)',
    'MATCH("",L!A1:A8,0)',
    'MATCH(2,L!A1:A8,1)',
    'MATCH(2,L!A1:A8,0)',
    'MATCH(13,L!A:A,0)',
    'MATCH(TRUE,L!G1:G8,0)',
    'MATCH(FALSE,L!G1:G8,1)',
    'MATCH("ЁЖ",L!Q1:Q4,0)',
    'MATCH("ЕЖ",L!Q1:Q4,0)',
    'MATCH("ИОД",L!Q1:Q4,0)',
    'MATCH("й*",L!Q1:Q4,0)',
    'MATCH(25,L!O1:O6,-1)',
    'MATCH(60,L!O1:O6,-1)',
    'MATCH(-5,L!O1:O6,-1)',
    'MATCH(30,L!O1:O6,-1)',
    'MATCH(5,D!L1:L6,NA())',
    'MATCH(NA(),D!L1:L6,1/0)',
    // sorted with blanks, errors and mixed types
    ...[0, 1, 2, 4, 6, 8].map((k) => `MATCH(${k},L!I1:I8,1)`),
    ...[2.5, 3, 3.5, 9].map((k) => `MATCH(${k},D!E1:E8,1)`),
    ...['4', '10', '"b"', '"z"', 'TRUE', 'FALSE', '0'].map((k) => `MATCH(${k},L!G1:G8,1)`),
    // unsorted: Excel's binary search
    ...[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map((k) => `MATCH(${k},L!K1:K7,1)`),
    ...[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map((k) => `MATCH(${k},L!K1:K7,-1)`),
    ...['0', '1', '2', '3', '4', '5', '6', '"a"', '"b"', '"x"', '"z"', 'TRUE', 'FALSE'].map((k) => `MATCH(${k},L!M1:M7,1)`),
  )
  add(
    'LOOKUP',
    'LOOKUP(D!A1,L!A1:A8,L!B1:B8)',
    'LOOKUP(8,L!A1:A8,L!B1:B3)',
    'LOOKUP(8,L!A1:A8,L!B1)',
    'LOOKUP(2,L!A1:A8,L!B1:B8)',
    'LOOKUP("c",L!D1:D6,L!E1:E6)',
    'LOOKUP(5,D!L8:Q10)',
    'LOOKUP(5,D!P1:Q2)',
    'LOOKUP(5,D!L1:N6,D!N1:N6)',
    'LOOKUP(5,D!L1:L6,D!L8:Q8)',
    'LOOKUP("b*",L!D1:D6,L!E1:E6)',
    'LOOKUP(TRUE,L!G1:G8)',
    'LOOKUP(3,D!E1:E8)',
    'LOOKUP(4,L!M1:M7)',
    'LOOKUP("b",L!M1:M7)',
    'LOOKUP(100,L!A:A,L!B:B)',
    'LOOKUP(5,D!L1:L6,D!N1:N6)',
  )
  add(
    'XMATCH',
    'XMATCH(5,D!L1:L6)',
    'XMATCH(4,D!L1:L6,-1)',
    'XMATCH(4,D!L1:L6,1)',
    'XMATCH("th*",D!M1:M6,2)',
    'XMATCH("TH*",D!M1:M6,0)',
    'XMATCH("THREE",D!M1:M6)',
    'XMATCH(2,L!A1:A8,0,-1)',
    'XMATCH(5,D!L1:L6,0,2)',
    'XMATCH(5,D!L1:N6)',
    'XMATCH(5,D!L1:L6,3)',
    'XMATCH(5,D!L1:L6,0,0)',
    'XMATCH(D!A8,D!L1:L6)',
    'XMATCH(D!A1,L!A1:A8)',
    'XMATCH(5,D!L8:Q8)',
    'XMATCH(100,D!L1:L6,-1)',
    'XMATCH(-1,D!L1:L6,1)',
    'XMATCH("Ёж",L!Q1:Q4)',
  )
  add(
    'INDEX',
    'INDEX(D!L1:N6,2)',
    'INDEX(D!L8:Q8,3)',
    'INDEX(D!L1:L6,0)',
    'INDEX(D!L1:N1,0,2)',
    'INDEX(D!L1:N6,2,0)',
    'INDEX(D!L1:N6,6,3)',
    'INDEX(D!L1:N6,1,4)',
    'INDEX(D!L1:N6,1,1,1)',
    'INDEX(D!L1:N6,1,1,2)',
    'INDEX(D!A1,1,1)',
    'INDEX(D!L1:L6,7)',
    'INDEX(D!L8:Q8,0)',
    'INDEX(D!L8:Q8,1,3)',
    'INDEX(D!L:L,3)',
    'INDEX(D!8:8,12)',
    'INDEX(D!L1:N6,10000000000)',
    'INDEX(D!L1:N6,-1)',
    'INDEX(D!L1:L1,0,0)',
    'INDEX(D!E1:E8,3)',
    'INDEX(D!L1:N6,1,1,0)',
  )
  add('ROW', 'ROW()', 'ROW()+1', 'ROW(D!A5:C5)', 'ROW(D!A:A)', 'ROW(D!C10)', 'ROW(D!A1048576)', 'ROW(L!B3)')
  add('COLUMN', 'COLUMN()', 'COLUMN(D!B3:B9)', 'COLUMN(D!3:3)', 'COLUMN(D!XFD1)', 'COLUMN(L!AA1)')
  add(
    'ADDRESS',
    ...['2,3', '2,3,2', '2,3,3', '2,3,4', '2,3,1,FALSE', '2,3,4,FALSE', '2,3,2,FALSE', '2,3,3,FALSE', '1,16384', '1048576,1', '0,1', '1,16385', '1048577,1', '1,1,5', '1,1,0', '2.7,3.9', '"2","3"', 'TRUE,1', '1,1,1,0', '1,27', '1,1,D!A1', '1,1,1,TRUE,D!A1'].map((a) => `ADDRESS(${a})`),
    ...['Sheet1', 'My Sheet', '1abc', 'A1', 'Лист1', 'a-b', 'a_b', 'a.b', "O'Brien", '', 'R1C1', 'TRUE', '_x', 'abc1', 'XFD1', 'XFE1', 'AAAA1', 'R1', 'C1', 'RC', 'a1b', 'R', 'C'].map((s) => `ADDRESS(1,1,1,TRUE,${q(s)})`),
  )
  add('HYPERLINK', 'HYPERLINK("http://x","name")', 'HYPERLINK("http://x")', 'HYPERLINK("x",5)', 'HYPERLINK("x",D!A1)', 'HYPERLINK(5)', 'HYPERLINK("x",TRUE)', 'HYPERLINK(NA(),"a")', 'HYPERLINK("x",NA())')

  // ---------- database ----------
  const db = 'B!A1:E9'
  const blocks = ['G1:G2', 'H1:H2', 'I1:I2', 'J1:J2', 'K1:K2', 'L1:L2', 'M1:M3', 'N1:O2', 'P1:Q2', 'R1:R2', 'S1:S2', 'T1:T2', 'U1:U2', 'V1:V2', 'W1:W2', 'X1:X2', 'Y1', 'Z1:Z3', 'AA1:AA2', 'AB1:AB2', 'AC1:AC2', 'AD1:AD2']
  for (const b of blocks) {
    add('DCOUNTA', `DCOUNTA(${db},"Tree",B!${b})`)
    add('DSUM', `DSUM(${db},"Profit",B!${b})`)
  }
  for (const f of ['5', '5.9', '"PROFIT"', 'TRUE', '"5"', 'D!A1', '6', '0', '-1', '1/0', '"Prof"', '" Profit"']) add('DSUM', `DSUM(${db},${f},B!H1:H2)`)
  for (const fn of ['DSUM', 'DCOUNT', 'DCOUNTA', 'DAVERAGE', 'DMAX', 'DMIN', 'DPRODUCT', 'DSTDEV', 'DSTDEVP', 'DVAR', 'DVARP', 'DGET']) add(fn, `${fn}(${db},"Age",B!G1:G2)`)
  for (const fn of ['DSUM', 'DCOUNT', 'DCOUNTA', 'DAVERAGE']) add(fn, `${fn}(${db},"Yield",B!Z1:Z3)`)
  add('DGET', `DGET(${db},"Age",B!AE1:AE2)`, `DGET(${db},"Yield",B!AF1:AF2)`, `DGET(${db},"Tree",B!AF1:AF2)`, `DGET(${db},"Tree",B!I1:I2)`, `DGET(${db},"Tree",B!H1:H2)`)
  for (const fn of ['DSUM', 'DCOUNT', 'DCOUNTA', 'DMAX', 'DGET']) add(fn, `${fn}(D!E1:E8,1,B!AG1:AG2)`, `${fn}(D!E1:E8,1,B!AH1:AH2)`)
  add('DSUM', 'DSUM(B!A:E,"Profit",B!H1:H2)')
  add('DCOUNTA', 'DCOUNTA(B!A:E,"Tree",B!Z1:Z3)', `DCOUNTA(${db},"Tree",B!G:G)`)
  add('DGET', 'DGET(B!A:E,"Tree",B!Z1:Z3)')

  // ---------- logical and information ----------
  add(
    'SWITCH',
    'SWITCH(D!A1,0,"zero","x")',
    'SWITCH(D!A1,"","empty","x")',
    'SWITCH(D!A1,D!A1,"blank","x")',
    'SWITCH("",D!A1,"b","x")',
    'SWITCH(1,TRUE,"t","x")',
    'SWITCH("1",1,"n","x")',
    'SWITCH(1,1,1/0,"x")',
    'SWITCH(1,2,1/0,"x")',
    'SWITCH(1,2,"two",1/0)',
    'SWITCH(NA(),1,"a")',
    'SWITCH(1,NA(),"a",1,"b")',
    'SWITCH(2,1,"a",2)',
    'SWITCH("A","a","lower")',
    'SWITCH("й","и","no","yes")',
    'SWITCH(D!A2,2,"num","text")',
    'SWITCH(0,D!A1,"blank","x")',
  )
  add('IFNA', 'IFNA(1/0,"x")', 'IFNA(NA(),1/0)', 'IFNA(MATCH(99,D!L1:L6,0),"nf")', 'IFNA(D!A13,"x")')
  add('ISEVEN', 'ISEVEN(-0.5)', 'ISEVEN(D!A9)', 'ISEVEN(-3)')
  add('ISODD', 'ISODD(-1.5)', 'ISODD(D!A3)', 'ISODD(7.9)')
  add('TYPE', 'TYPE(1/0)', 'TYPE(D!A1:A1)', 'TYPE(D!K1:K2)', 'TYPE(NA())', 'TYPE("")', 'TYPE(D!A4)')

  // ---------- exploration (not implemented yet) ----------
  for (let n = 1; n <= 255; n++) if (n >= 120 || n % 15 === 0) add('CHAR', `CHAR(${n})`)
  add('CODE', ...['€', 'È', '»', 'Ж', 'é', 'ÿ', 'ß', '😀', 'Ā', 'ł', '‚', '—'].map((c) => `CODE(${q(c)})`))
  add('LENB', 'LENB("日本")', 'LENB("Привет")', 'LENB(T!C5)', 'LENB(T!C6)')
  add('LEFTB', 'LEFTB("日本語",2)', 'LEFTB("Привет",3)')
  add('RIGHTB', 'RIGHTB("日本語",2)')
  add('MIDB', 'MIDB("日本語",2,1)')
  add('FINDB', 'FINDB("本","日本語")')
  add('SEARCHB', 'SEARCHB("本","日本語")')
  add('REPLACEB', 'REPLACEB("日本語",2,1,"x")')
  add('ASC', 'ASC(T!C5)', 'ASC(T!C6)', 'ASC("abc")', 'ASC("ＡＢＣ")')
  add('DBCS', 'DBCS(T!C5)', 'DBCS("abc")', 'DBCS("Привет")')
  add('JIS', 'JIS("abc")', 'JIS(T!C5)')
  add('BAHTTEXT', 'BAHTTEXT(1234)', 'BAHTTEXT(0)', 'BAHTTEXT(1.25)', 'BAHTTEXT(-21)', 'BAHTTEXT(1000000)', 'BAHTTEXT(11)', 'BAHTTEXT(101)', 'BAHTTEXT(0.5)')
  add('PHONETIC', 'PHONETIC(T!C7)', 'PHONETIC(D!B1)', 'PHONETIC(T!C4:C5)', 'PHONETIC(D!A1)')
  return out
}
