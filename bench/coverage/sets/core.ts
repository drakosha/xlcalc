/**
 * Excel rules of the engine's core that the random check pointed at, case by case, as dynamic array formulas
 * (Excel 365): the order in which an operator converts its operands and finds errors, arrays of one value under
 * operators and lifted calls, HSTACK and VSTACK over errors given as values and as references, functions that
 * pass a reference on, VLOOKUP and HLOOKUP lifted by the key with an index array, the argument counts Excel
 * loads. Answers: `docs/excel-cases/core.json`.
 */
import type { Probe } from '../probes.js'
import { collector, type Add } from './arrays/common.js'

export const DYNAMIC = true

/** Operands of each binary operator: a text that is not a number and an error on either side, two errors. */
const OPERANDS = [['"abc"', 'NA()'], ['NA()', '"abc"'], ['1/0', 'NA()'], ['NA()', '1/0'], ['D!A3', 'D!A8'], ['D!A8', 'D!A3'], ['""', 'NA()'], ['"abc"', '1/0']]

function order(add: Add) {
  for (const op of ['+', '-', '*', '/', '^', '&', '=', '<', '>=', '<>']) add(`op:${op}`, op, ...OPERANDS.map(([a, b]) => `${a}${op}${b}`))
  add('op:/', '/', '"abc"/0', 'NA()/0', '1/"abc"')
  add('unary', 'unary', '-"abc"', '-NA()', '"abc"%', 'NA()%', '-D!A3', 'D!A3%', 'TYPE(+"abc")', '-1/0')
  // lifted: each element converts its left operand first; a position outside a shorter array
  add('op:lift', 'lift', 'INDEX({"a",1}-NA(),1,1)', 'INDEX({"a",1}-NA(),1,2)', 'INDEX(NA()-{"a",1},1,1)', 'INDEX({"a";1}-{#N/A;#N/A},1)', 'INDEX({#N/A;1}-{"a";"b"},1)')
  add('op:lift', 'lift', 'INDEX({"a","b","c"}+{1,2},1,3)', 'INDEX({1,2}+{"a","b","c"},1,3)', 'INDEX({"a","b","c"}&{1,2},1,3)', 'INDEX({1,2,#DIV/0!}-{#N/A,2},1,3)')
  add('op:lift', 'lift', 'INDEX(D!A3:A4-NA(),1)', 'INDEX(D!A3:A4^NA(),2)', 'INDEX({"a",1}^{#N/A,2},1,1)', 'INDEX(-{"a",#N/A},1,1)', 'INDEX({"a";2}={#N/A;2},1)')
}

/** An array of one value stays an array under an operator or a lifted call (TYPE 64); a one-cell reference is a value. */
function single(add: Add) {
  add('1x1:op', 'TYPE', 'TYPE({1}+0)', 'TYPE(0+{1})', 'TYPE(-{1})', 'TYPE({1}%)', 'TYPE({1}&"")', 'TYPE({1}={1})', 'TYPE({1}>=TRUE)', 'TYPE({1}+D!B1)', 'TYPE(+{1})')
  add('1x1:op', 'TYPE', 'TYPE(D!B1:B1+0)', 'TYPE(D!B1+0)', 'TYPE(INDEX(D!B1:B8,1)+0)', 'TYPE(TAKE(D!B1:B8,1)+0)', 'TYPE(INDEX({1;2},1,0)+0)', 'TYPE(SEQUENCE(1)+0)', 'TYPE(TRANSPOSE({1})*1)')
  // an array skips text and booleans where a direct value is converted
  add('1x1:list', '1x1', 'DEVSQ({1}>=TRUE)', 'COUNT({"2"}&"")', 'COUNT("2"&"")', 'SUM({TRUE}=TRUE)', 'SUM(TRUE=TRUE)', 'AVERAGE({TRUE}=TRUE)', 'MAXA({"x"}&"")', 'SUM({1}+0)')
  add('1x1:call', 'TYPE', 'TYPE(ABS({-1}))', 'TYPE(LEN({"ab"}))', 'TYPE(ROUND({1.5},0))', 'TYPE(DATE({2025},1,1))', 'TYPE(INDEX({10,20},1,{2}))', 'TYPE(MATCH(5,D!L1:L6,{0}))', 'TYPE(COUNTIF(D!B1:B8,{4}))')
  add('1x1:call', 'TYPE', 'TYPE(N({1}))', 'TYPE(T({"a"}))', 'TYPE(ISNUMBER({1}))', 'TYPE(MID({"abc"},1,1))', 'TYPE(SUM({1}))', 'TYPE(SUMPRODUCT({1}+0))', 'TYPE(XLOOKUP(5,D!L1:L6,D!M1:M6)&"")')
  add('1x1:lazy', 'TYPE', 'TYPE(IF({TRUE},1,2))', 'TYPE(IF(TRUE,{1},2))', 'TYPE(IF(TRUE,{1},2)+0)', 'TYPE(IFERROR({1},0))', 'TYPE(IFERROR({1}/0,0))', 'TYPE(CHOOSE({1},1,2))', 'TYPE(CHOOSE(1,{1},2)+0)', 'TYPE(SWITCH({1},1,"a"))', 'TYPE(IFS({TRUE},1))')
  add('1x1:fn', 'TYPE', 'TYPE(TRANSPOSE(5))', 'TYPE(SORT(5))', 'TYPE(UNIQUE(5))', 'TYPE(MMULT(2,3))', 'TYPE(VSTACK(5))', 'TYPE(TAKE({1,2},1,1))', 'TYPE(FILTER({1,2},{1,0}))')
  add('1x1:fn', 'TYPE', 'TYPE(TRANSPOSE(D!B1))', 'TYPE(TRANSPOSE("a"))', 'TYPE(TRANSPOSE({5}))', 'TYPE(TOCOL(5))', 'TYPE(EXPAND(5,1))', 'TYPE(SORT(D!B1))', 'TYPE(UNIQUE(D!B1))', 'TYPE(MMULT(D!B1,D!C1))')
  add('1x1:fn', 'TYPE', 'TYPE(CHOOSECOLS(5,1))', 'TYPE(WRAPROWS(5,1))', 'TYPE(MINVERSE(2))', 'TYPE(MUNIT(1))', 'TYPE(SEQUENCE(1,1))', 'TYPE(TEXTSPLIT("a",","))', 'TYPE(MODE.MULT(1))', 'TYPE(INDEX({5},1,1))', 'TYPE(XLOOKUP({5},D!L1:L6,D!M1:M6))', 'TYPE(TRANSPOSE(5)+0)')
  add('1x1:lookup', 'VLOOKUP', 'TYPE(VLOOKUP({5},D!L1:N6,2,FALSE))', 'VLOOKUP({5},D!L1:N6,{2;4},FALSE)', 'ROWS(VLOOKUP({5},D!L1:N6,{2;3},FALSE))')
}

/** HSTACK and VSTACK: an error given as a value is the answer, a reference to an error cell is stacked. */
function stacks(add: Add) {
  add('stack', 'VSTACK', 'ROWS(VSTACK(1,1/0,2))', 'INDEX(VSTACK(1,1/0,2),1)', 'ROWS(VSTACK(1,D!A7,2))', 'ROWS(VSTACK(1/0,NA()))', 'ROWS(VSTACK(D!A7,NA()))', 'ROWS(VSTACK(1,{1,#N/A}))', 'ROWS(VSTACK(NA()))')
  add('stack', 'VSTACK', 'ROWS(VSTACK(1,INDEX(D!E1:E8,3)))', 'ROWS(VSTACK(1,(D!A7)))', 'ROWS(VSTACK(1,+D!A7))', 'ROWS(VSTACK(1,D!A7+0))', 'ROWS(VSTACK(1,IFERROR(1/0,NA())))', 'ROWS(VSTACK(1,XLOOKUP("abc",D!A3,D!A7)))')
  add('stack', 'VSTACK', 'ROWS(VSTACK(1,IF(TRUE,D!A7)))', 'ROWS(VSTACK(1,CHOOSE(1,D!A7)))', 'ROWS(VSTACK(1,IFS(TRUE,D!A7)))', 'ROWS(VSTACK(1,SWITCH(1,1,D!A7)))')
  add('stack', 'HSTACK', 'COLUMNS(HSTACK(D!B1:B2,D!E3))', 'COLUMNS(HSTACK(D!B1:B2,1/0))', 'COLUMNS(HSTACK(1/0))', 'INDEX(HSTACK(D!B1:B2,D!E3),1,2)', 'COLUMNS(HSTACK(D!B1,NA(),1/0))')
  // a function that answers its argument: a reference (SUM skips the text "2" of D!A2) or a value (SUM reads it)
  add('ref', 'ref', 'SUM(CHOOSE(1,D!A2))', 'SUM(IF(TRUE,D!A2))', 'SUM(IFS(TRUE,D!A2))', 'SUM(SWITCH(1,1,D!A2))', 'SUM(IFERROR(D!A2,0))', 'SUM(IFNA(D!A2,0))', 'SUM(XLOOKUP(1,D!L1:L6,D!A2:A7))', 'SUM(INDEX(D!A2:A3,1))', 'COUNT(CHOOSE(1,D!A2))')
  add('ref', 'ref', 'SUM(IFERROR(1/0,D!A2))', 'SUM(IFNA(NA(),D!A2))', 'SUM(XLOOKUP(99,D!L1:L6,D!A2:A7,D!A2))', 'SUM(FILTER(D!B1:B2,{0;0},D!A2))', 'SUM(SWITCH(9,1,1,D!A2))', 'SUM(IF(FALSE,1,D!A2))', 'SUM(IFS(FALSE,1,TRUE,D!A2))', 'SUM(CHOOSE(2,1,D!A2))')
  add('ref', 'ref', 'COUNTA(IF(TRUE,D!K1))', 'COUNTA(CHOOSE(1,D!K1))', 'COUNTBLANK(IF(TRUE,D!K1))', 'COLUMN(CHOOSE(1,D!C1))', 'COLUMN(IF(TRUE,D!C1))', 'ROW(IFS(TRUE,D!B3))', 'SUM(+D!A2)', 'SUM((D!A2))', 'COUNT(IF(TRUE,D!A2))', 'SUMIF(IF(TRUE,D!L1),">0")')
  add('ref', 'ref', 'ROWS(VSTACK(1,IFERROR(1/0,D!A7)))', 'ROWS(VSTACK(1,IFNA(NA(),D!A7)))', 'ROWS(VSTACK(1,XLOOKUP(99,D!L1:L6,D!A2:A7,D!A7)))', 'ROWS(VSTACK(1,FILTER(D!B1:B2,{0;0},D!A7)))', 'SUM(IF(TRUE,D!A2)+0)', 'SUM(IF({TRUE},D!A2))')
}

/** VLOOKUP and HLOOKUP lifted by the key or the match mode take the index's top-left value in each call. */
function lookups(add: Add) {
  add('lookup', 'VLOOKUP', 'INDEX(VLOOKUP({5;8},D!L1:N6,{"x";2},FALSE),1)', 'INDEX(VLOOKUP({5;8},D!L1:N6,{#N/A;2},FALSE),1)', 'INDEX(VLOOKUP({5;8},D!L1:N6,{"3";2},FALSE),2)', 'INDEX(VLOOKUP({5;8},D!L1:N6,{2.9;9},FALSE),2)')
  add('lookup', 'VLOOKUP', 'INDEX(VLOOKUP({5;8},D!L1:N6,D!A7:A8,FALSE),1)', 'INDEX(VLOOKUP({5;8},D!L1:N6,D!A2:A3,FALSE),1)', 'INDEX(VLOOKUP({5;8},D!L1:N6,D!AH3:AH5,FALSE),2)', 'INDEX(VLOOKUP({5;8},D!L1:N6,{3,9;9,9},FALSE),2)')
  add('lookup', 'VLOOKUP', 'INDEX(VLOOKUP({5;8},D!L1:N6,{TRUE;2},FALSE),1)', 'INDEX(VLOOKUP(D!L4:L5,D!L1:N6,{3;0},FALSE),2)', 'INDEX(VLOOKUP({5;8},D!L1:N6,D!A1:A2,FALSE),1)', 'INDEX(VLOOKUP({5;8},D!L1:N6,{0;2},TRUE),2)')
  add('lookup', 'HLOOKUP', 'INDEX(HLOOKUP({5,8},D!L8:Q10,{2;4},FALSE),1,1)', 'INDEX(HLOOKUP({5,8},D!L8:Q10,{4;2},FALSE),1,2)', 'INDEX(HLOOKUP({5,8},D!L8:Q10,D!AH3:AH4,FALSE),1,2)', 'INDEX(HLOOKUP(5,D!L8:Q10,{3;9},{FALSE,TRUE}),1,2)')
  add('lookup', 'HLOOKUP', 'COLUMNS(HLOOKUP({5,8},D!L8:Q10,{2;4},FALSE))', 'ROWS(HLOOKUP({5,8},D!L8:Q10,{2;4},FALSE))', 'INDEX(HLOOKUP({5,8},D!L8:Q10,{"x";2},FALSE),1,1)')
}

/**
 * Argument counts Excel loads (it refuses the book otherwise, so only those are cases): SORTBY with 2, 3 or 5
 * arguments, not 4 or 6 (each further by-array with its order); SWITCH with 126 pairs and a default, not 127 pairs.
 */
function arity(add: Add) {
  const pairs = Array.from({ length: 126 }, (_, i) => `${i + 1},${i + 1}`).join(',')
  add('arity', 'SORTBY', 'ROWS(SORTBY(D!B1:B8,D!C1:C8))', 'ROWS(SORTBY(D!B1:B8,D!C1:C8,-1))', 'ROWS(SORTBY(D!B1:B8,D!C1:C8,-1,D!E1:E8,1))')
  add('arity', 'SWITCH', `SWITCH(1,${pairs},0)`)
}

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const part of [order, single, stacks, lookups, arity]) part(add)
  return out
}
