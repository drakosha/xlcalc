/**
 * Excel cases for an omitted argument (`IF(A1,,1)`, `SUM(1,,2)`): what each kind of argument takes for it
 * (`src/args.ts`), an omitted last optional argument against one not given (`LEFT("abc",)` and `LEFT("abc")`),
 * an omitted required argument, the functions whose omitted argument may mean their default (`XLOOKUP`,
 * `SEQUENCE`, `TEXTSPLIT`), and omitted arguments under lifting and inside LET. Answers: `test/excel/cases/omit.json`.
 */
import type { Probe } from '../probes.js'
import { collector } from './arrays/common.js'
import { all } from './omit/all.js'
import { kinds } from './omit/kinds.js'

/** Calls Excel does not load: an omitted argument where Excel takes only a reference. */
export const REFUSED: readonly string[] = ['SUMIF(D!B1:B8,">10",)']

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  kinds(add)
  all(add)
  return out
}
