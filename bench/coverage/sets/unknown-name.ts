/**
 * Unknown names and functions where they are not computed, held out: a host's books hold `IF(ISNA(R),"",R)` with `""`
 * stored where the engine gave `#NAME?` for an unknown name inside. A name or a function the book does not define, in
 * the branch of IF that is not taken, after IFERROR's value, in CHOOSE's other choices, and where they are computed (in
 * the condition, inside the lookups, ISNA and ISERROR over them). Computed in Excel en-US on the Mac and Windows
 * (`test/excel/cases/unknown-name.json`, `test/excel/win/cases/unknown-name.json`); `unknown-name.predicted.json`, the
 * engine's answers, was committed before Excel ran the set.
 */
import type { Probe, SheetCells } from '../probes.js'

/** The cells the lookups read: 1 to 3 in A, their doubles in B. */
export const SHEETS: SheetCells = { E: { A1: 1, A2: 2, A3: 3, B1: 2, B2: 4, B3: 6 } }

const LOOKUP = 'MATCH(99,E!A1:A3,0)'

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const add = (group: string, formula: string) => out.push({ id: `${group}#${out.filter((p) => p.group === group).length + 1}`, group, formula })
  for (const u of ['Nope', 'NOPE(1)', 'Nope+1', 'SUM(Nope)']) {
    add('untaken', `IF(ISNA(${LOOKUP}),"",${u})`)
    add('untaken', `IF(TRUE,1,${u})`)
    add('untaken', `IF(FALSE,${u},2)`)
    add('untaken', `IFERROR(1,${u})`)
    add('untaken', `IFNA(1,${u})`)
    add('untaken', `CHOOSE(1,3,${u})`)
    add('untaken', `IF(ISNA(VLOOKUP(99,E!A1:B3,2,0)),"",VLOOKUP(99,E!A1:B3,2,0)+${u})`)
    add('taken', `IF(ISNA(${u}),"",1)`)
    add('taken', `ISERROR(${u})`)
    add('taken', `IFERROR(${u},5)`)
    add('taken', `IF(TRUE,${u},1)`)
  }
  add('lookup', 'IF(ISNA(VLOOKUP(99,E!A1:B3,Nope,0)),"",1)')
  add('lookup', 'IF(ISNA(VLOOKUP(99,E!A1:B3,2,Nope)),"",1)')
  add('lookup', 'IF(ISNA(MATCH(99,E!A1:A3,Nope)),"",1)')
  add('lookup', 'IF(ISNA(MATCH(Nope,E!A1:A3,0)),"",1)')
  add('lookup', 'IF(ISNA(INDEX(E!B1:B3,MATCH(99,E!A1:A3,0),Nope)),"",1)')
  add('lookup', 'IF(ISNA(HLOOKUP(99,E!A1:B3,Nope,0)),"",1)')
  add('lookup', 'IF(ISNA(LOOKUP(99,E!A1:A3,Nope)),"",1)')
  add('lookup', 'IF(ISNA(VLOOKUP(99,Nope,2,0)),"",1)')
  return out
}
