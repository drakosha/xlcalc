/**
 * REGEXTEST, REGEXEXTRACT and REGEXREPLACE held out, a set of Excel for Windows alone (Excel for Mac 16.114 has none
 * of them): a test, the first match, all matches and the groups of the first, replacement of every match, the nth and
 * the nth from the end, `$1` in the replacement, case sensitivity, no match. `PREDICTED`, from Microsoft's
 * descriptions of the functions, was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { b, e, s } from '../vals.js'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  ['REGEXTEST', 'REGEXTEST("Excel 2025","\\d+")', b(true)],
  ['REGEXTEST', 'REGEXTEST("Excel","excel")', b(false)],
  ['REGEXTEST', 'REGEXTEST("Excel","excel",1)', b(true)],
  ['REGEXTEST', 'REGEXTEST("abc","[")', e('VALUE')],
  ['REGEXEXTRACT', 'REGEXEXTRACT("tel 555-1234 or 555-9876","\\d{3}-\\d{4}")', s('555-1234')],
  ['REGEXEXTRACT', 'REGEXEXTRACT("abc","\\d")', e('NA')],
  ['REGEXEXTRACT', 'INDEX(REGEXEXTRACT("a1b22c333","\\d+",1),2)', s('22')],
  ['REGEXEXTRACT', 'ROWS(REGEXEXTRACT("a1b22c333","\\d+",1))', { t: 'n', v: 1 }],
  ['REGEXEXTRACT', 'COLUMNS(REGEXEXTRACT("a1b22c333","\\d+",1))', { t: 'n', v: 3 }],
  ['REGEXEXTRACT', 'INDEX(REGEXEXTRACT("John Smith","(\\w+) (\\w+)",2),2)', s('Smith')],
  ['REGEXEXTRACT', 'COLUMNS(REGEXEXTRACT("John Smith","(\\w+) (\\w+)",2))', { t: 'n', v: 2 }],
  ['REGEXEXTRACT', 'REGEXEXTRACT("ABC","b",0,1)', s('B')],
  ['REGEXREPLACE', 'REGEXREPLACE("a1b2c3","\\d","#")', s('a#b#c#')],
  ['REGEXREPLACE', 'REGEXREPLACE("a1b2c3","\\d","#",2)', s('a1b#c3')],
  ['REGEXREPLACE', 'REGEXREPLACE("a1b2c3","\\d","#",-1)', s('a1b2c#')],
  ['REGEXREPLACE', 'REGEXREPLACE("SoniaBrown","([A-Z][a-z]+)([A-Z][a-z]+)","$2, $1")', s('Brown, Sonia')],
  ['REGEXREPLACE', 'REGEXREPLACE("abc","B","x",0,1)', s('axc')],
  ['REGEXREPLACE', 'REGEXREPLACE("abc","\\d","x")', s('abc')],
]

/** Excel for Windows 16.0.20430 answered #NAME? everywhere: it has none of the three functions. */
export const MISSED = ['REGEXTEST#1', 'REGEXTEST#2', 'REGEXTEST#3', 'REGEXTEST#4', 'REGEXEXTRACT#1', 'REGEXEXTRACT#2', 'REGEXEXTRACT#3', 'REGEXEXTRACT#4', 'REGEXEXTRACT#5', 'REGEXEXTRACT#6', 'REGEXEXTRACT#7', 'REGEXEXTRACT#8', 'REGEXREPLACE#1', 'REGEXREPLACE#2', 'REGEXREPLACE#3', 'REGEXREPLACE#4', 'REGEXREPLACE#5', 'REGEXREPLACE#6']

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
