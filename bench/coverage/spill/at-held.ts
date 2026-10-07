/**
 * Held-out cases of the operator `@` (`run.ts spill at-held`): answers predicted from the rules of the set `at` and
 * committed before Excel computed them (`PREDICTED`).
 */
import { one, sheet, type CellSpec, type SpillCase } from './cases.js'
import { PREDICTED } from './at-held.predicted.js'

const DATA: Record<string, CellSpec> = {
  ...Object.fromEntries(Array.from({ length: 5 }, (_, i) => [`A${i + 1}`, i + 1])),
  ...Object.fromEntries(Array.from({ length: 5 }, (_, i) => [`B${i + 1}`, (i + 1) * 10])),
  ...Object.fromEntries(Array.from({ length: 5 }, (_, i) => [`C${i + 1}`, (i + 1) * 100])),
}

export function buildSpillCases(): SpillCase[] {
  return [
    {
      ...one('mixed', '@ over two columns, a call, a whole column, an array, an error counted', { ...DATA, H3: '=@B1:B5+@C1:C5', I2: '=@INDEX(A1:C5,0,3)', H4: '=SUM(@A:A,1)', H1: '=@{7,8;9,10}', I5: '=@(B1:B5/10)', H9: '=COUNTA(@A1:A5)', I3: '=@T!A1:A5' }, ['H1:I9'], {
        book: { sheets: [sheet('T', { A1: 'a', A2: 'b', A3: 'c', A4: 'd', A5: 'e' })] },
      }),
      predicted: PREDICTED.mixed,
    },
  ]
}
