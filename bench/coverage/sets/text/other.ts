/** Cases of the database, logical and information functions. */
import type { Add } from './data.js'

export function database(add: Add) {
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
  for (const b of ['AI1:AI2', 'AJ1:AJ2', 'AK1:AK2', 'AL1:AL2', 'AM1:AM2', 'AN1:AN2', 'AO1:AO2', 'AP1:AP3']) add('DCOUNTA', `DCOUNTA(${db},"Tree",B!${b})`)
  for (const fn of ['DAVERAGE', 'DMIN', 'DPRODUCT', 'DSTDEV', 'DSTDEVP', 'DVAR', 'DVARP']) add(fn, `${fn}(D!E1:E8,1,B!AG1:AG2)`, `${fn}(D!E1:E8,1,B!AH1:AH2)`)
}

export function logical(add: Add) {
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
    'SWITCH(D!A1,FALSE,"f","x")',
    'SWITCH(D!A13,D!A1,"b","x")',
  )
  add('IFNA', 'IFNA(1/0,"x")', 'IFNA(NA(),1/0)', 'IFNA(MATCH(99,D!L1:L6,0),"nf")', 'IFNA(D!A13,"x")')
  add('ISEVEN', 'ISEVEN(-0.5)', 'ISEVEN(D!A9)', 'ISEVEN(-3)')
  add('ISODD', 'ISODD(-1.5)', 'ISODD(D!A3)', 'ISODD(7.9)')
  add('TYPE', 'TYPE(1/0)', 'TYPE(D!A1:A1)', 'TYPE(D!K1:K2)', 'TYPE(NA())', 'TYPE("")', 'TYPE(D!A4)')
}
