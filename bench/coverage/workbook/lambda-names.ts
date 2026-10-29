/**
 * LAMBDAs in defined names, as Name Manager keeps them: called from cells with their arguments (fewer, more, one left
 * out), calling themselves and each other, a name's LAMBDA using another name, one scoped to a sheet, a name of a value
 * called, a LAMBDA left without a call, a name given to MAP and BYROW as a function, a LAMBDA over a range, and a cell its
 * calculation reads changed after opening (the cells that call it computed again). Answers predicted before Excel ran
 * (`lambda-names.predicted.json`, `workbook/predict.ts lambda-names`).
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { WorkbookCase } from './cases.js'
import { formulaCell as f } from '../stored.js'

const NAMES = [
  { name: 'Twice', ref: 'LAMBDA(x,x*2)' },
  { name: 'Fact', ref: 'LAMBDA(n,IF(n<=1,1,n*Fact(n-1)))' },
  { name: 'AddOpt', ref: 'LAMBDA(a,b,IF(ISOMITTED(b),a,a+b))' },
  { name: 'Pct', ref: '0.1' },
  { name: 'Scaled', ref: 'LAMBDA(x,x*Pct)' },
  { name: 'IsEven', ref: 'LAMBDA(n,IF(n=0,TRUE,IsOdd(n-1)))' },
  { name: 'IsOdd', ref: 'LAMBDA(n,IF(n=0,FALSE,IsEven(n-1)))' },
  { name: 'Cnt', ref: 'LAMBDA(n,IF(n<=0,0,1+Cnt(n-1)))' },
  { name: 'PlusA1', ref: "LAMBDA(x,x+'S'!$A$1)" },
  { name: 'Local', ref: 'LAMBDA(x,x+100)', scope: 'S' },
  { name: 'Rows2', ref: 'LAMBDA(r,ROWS(r)*10+COLUMNS(r))' },
]

const FS = [
  'Twice(3)',
  'Fact(5)',
  'Fact(20)',
  'AddOpt(1)',
  'AddOpt(1,)',
  'AddOpt(1,2)',
  'Twice(1,2)',
  'Pct(2)',
  'Scaled(10)',
  'IsEven(10)',
  'IsOdd(7)',
  'Cnt(100)',
  'Cnt(150)',
  'Cnt(300)',
  'PlusA1(1)',
  'Twice',
  'Local(1)',
  'S!Local',
  'ARRAYTOTEXT(MAP(S!A1:A3,Twice),1)',
  'ARRAYTOTEXT(BYROW(S!A1:B3,Rows2),1)',
  'ARRAYTOTEXT(Twice(S!A1:A3),1)',
  'Rows2(S!A1:B3)',
  'twice(4)',
  'Fact(Twice(2))',
  'LET(f,Twice,f(5))',
  'ISREF(LAMBDA(r,r)(S!A1))',
]

const book = {
  names: NAMES,
  sheets: [
    { name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, B1: { value: 4 }, B2: { value: 5 }, B3: { value: 6 }, C1: f('Local(1)') } },
    { name: 'F', cells: Object.fromEntries(FS.map((x, i) => [`A${i + 1}`, f(x)])) },
  ],
}
const read = [
  { sheet: 'F', range: `A1:A${FS.length}` },
  { sheet: 'S', range: 'C1' },
]

const PREDICTED = JSON.parse(readFileSync(new URL('./lambda-names.predicted.json', import.meta.url), 'utf8')) as Record<string, Record<string, Val>[]>

export function buildWorkbookCases(): WorkbookCase[] {
  const cases: WorkbookCase[] = [
    {
      id: 'names',
      about: 'LAMBDAs in names called from cells, then the cell PlusA1 reads set to 10',
      book,
      read,
      steps: ['full', 'read', { set: { sheet: 'S', cell: 'A1', value: 10 } }, 'calc', 'read'],
    },
  ]
  return cases.map((c) => (PREDICTED[c.id] ? { ...c, predicted: PREDICTED[c.id] } : c))
}
