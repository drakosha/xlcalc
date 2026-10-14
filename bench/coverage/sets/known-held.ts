/**
 * The rules that closed three known differences, held out: SORTBY lifted over its sort orders (successful sorts too,
 * an order array of one value, blanks or the first value in each element), LOOKUP past the end of a shorter result
 * (one cell, a row, a computed lookup array, the sheet's end, an array result), and IFS answering the #VALUE! of a
 * text condition as an array of one error (TYPE, VSTACK, an error given as a condition or a value, no condition
 * met), with SWITCH, IF and CHOOSE beside it. Answers: `test/excel/cases/known-held.json`; `PREDICTED`, the engine's
 * answers, was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { n, s, e } from '../vals.js'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // SORTBY: an array of orders lifts the call, each element the top-left value of its sort (D!M1:M6 by D!N1:N6, D!B1:B8 by D!C1:C8)
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!M1:M6,D!N1:N6,{1;-1}),1)', s('{"one";"thirteen"}')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,{1,-1}),1)', s('{-3.5,42}')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,{1;-1},D!B1:B8,{1,-1}),1)', s('{-3.5,-3.5;42,42}')],
  ['SORTBY', 'SUM(SORTBY(D!B1:B8,D!C1:C8,{1,-1}))', n(38.5)],
  ['SORTBY', 'INDEX(SORTBY(D!M1:M6,D!N1:N6,{1;1;1}),3)', s('one')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!M1:M6,D!N1:N6,{1;2;-1}),1)', s('{"one";#VALUE!;"thirteen"}')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!M1:M6,D!N1:N6,D!A4:A5),1)', s('{"one";#VALUE!}')],
  ['SORTBY', 'ROWS(SORTBY(D!M1:M6,D!N1:N6,{1}))', n(1)],
  ['SORTBY', 'TYPE(SORTBY(D!M1:M6,D!N1:N6,{1}))', n(64)],
  ['SORTBY', 'SORTBY(D!M1:M6,D!N1:N6,{-1})', s('thirteen')],
  // LOOKUP past the end of a shorter result (4.19 falls in D!L3, 5 in D!O8)
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,D!B2)', e('REF')],
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,D!L9)', e('REF')],
  ['LOOKUP', 'LOOKUP(5,D!L8:Q8,D!B2)', e('REF')],
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,D!L9:M9)', s('three')],
  ['LOOKUP', 'LOOKUP(5,D!L8:Q8,D!B2:B3)', n(23)],
  ['LOOKUP', 'LOOKUP(4.19,{1;2;3;5},D!B2:B3)', n(16)],
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6+0,D!B2:B3)', e('REF')],
  ['LOOKUP', 'LOOKUP(100,D!L1:L6,D!B2:B3)', n(-3.5)],
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,D!B1048575:B1048576)', e('REF')],
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,{1;2})', e('NA')],
  // IFS: the #VALUE! of a text condition is an array of one error; an error given as a condition or a value, or no condition met, is the answer itself
  ['IFS', 'TYPE(IFS("x",1))', n(64)],
  ['IFS', 'COLUMNS(IFS("x",1))', n(1)],
  ['IFS', 'ROWS(IFS(D!A3,1))', n(1)],
  ['IFS', 'ROWS(IFS("",1))', n(1)],
  ['IFS', 'ROWS(IFS(FALSE,1,"x",2))', n(1)],
  ['IFS', 'ROWS(VSTACK(1,IFS("x",1)))', n(2)],
  ['IFS', 'ROWS(IFS(1/0,1))', e('DIV0')],
  ['IFS', 'ROWS(IFS(D!A7,1))', e('DIV0')],
  ['IFS', 'ROWS(IFS(TRUE,1/0))', e('DIV0')],
  ['IFS', 'ROWS(IFS(FALSE,1))', e('NA')],
  ['IFS', 'TYPE(IFS(FALSE,1))', n(16)],
  // beside IFS: SWITCH with no value matched, IF with a text condition, CHOOSE past its values
  ['SWITCH', 'ROWS(SWITCH(3,1,"a"))', e('NA')],
  ['SWITCH', 'ROWS(SWITCH("x",1,"a"))', e('NA')],
  ['SWITCH', 'TYPE(SWITCH(3,1,"a"))', n(16)],
  ['IF', 'ROWS(IF("x",1,2))', e('VALUE')],
  ['IF', 'TYPE(IF("x",1,2))', n(16)],
  ['CHOOSE', 'ROWS(CHOOSE(5,1,2))', e('VALUE')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

/**
 * The predictions Excel did not confirm (the Mac and Windows alike): SORTBY puts the first sorted value in the first
 * place only, the others blank (or an error), with no row or column of an order repeated; LOOKUP reads a result of
 * one cell as a row, reads on with a computed lookup array too, and past the sheet's end reads a blank.
 */
export const MISSED: readonly string[] = ['SORTBY#1', 'SORTBY#2', 'SORTBY#3', 'SORTBY#4', 'SORTBY#5', 'SORTBY#6', 'LOOKUP#1', 'LOOKUP#2', 'LOOKUP#3', 'LOOKUP#7', 'LOOKUP#9']
