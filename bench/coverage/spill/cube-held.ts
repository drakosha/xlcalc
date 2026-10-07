/**
 * Held-out references across sheets (`run.ts spill cube-held`): answers predicted from the rules of the set `areas`
 * and committed before Excel computed them (`PREDICTED`). Four sheets Q1…Q4 in order; their names read as cell
 * addresses, so every reference across them is quoted, as the engine writes it.
 */
import { sheet, type SpillCase } from './cases.js'
import { PREDICTED } from './cube-held.predicted.js'

const FORMULAS: Record<string, string> = {
  A1: "=SUM('Q1:Q3'!A1:A2)",
  A2: "=AVERAGE('Q2:Q4'!A1)",
  A3: "=MAX('Q1:Q4'!A:A)",
  A4: "=COUNT('Q1:Q4'!A1:B2)",
  A5: "=SUM('Q3:Q1'!A1)",
  A6: "=ROWS(VSTACK('Q1:Q4'!A1:A2))",
  A7: "=INDEX(VSTACK('Q1:Q4'!A1:A2),5)",
  A8: "=SUM('Q1:Q4'!A1:A2*2)",
  A9: "='Q1:Q2'!A1",
  A10: "=COUNTA('Q1:Q4'!A1:B2)",
  A11: "=STDEV('Q1:Q4'!A1)",
  A12: "=PRODUCT('Q1:Q2'!A1:A2)",
  A13: "=SUMIF('Q1:Q4'!A1:A2,\">1\")",
  A14: "=SUM('Q2:Q2'!A1:A2)",
}

export function buildSpillCases(): SpillCase[] {
  const q = (n: number) => sheet(`Q${n}`, { A1: n, A2: n * 10, B1: n % 2 === 0 ? 'x' : true })
  return [
    {
      id: 'sums',
      about: 'references across four sheets in lists, in reverse order, stacked, under an operator, alone',
      book: { sheets: [q(1), q(2), q(3), q(4), sheet('S', FORMULAS)] },
      read: [{ sheet: 'S', range: 'A1:A14' }],
      predicted: PREDICTED.sums,
    },
  ]
}
