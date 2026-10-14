/**
 * Held out after `let-grid-held`, where Excel took every cell of the sheet tried as LET's name (`IV1`, `A65536`, `XFD1`)
 * but `R1` and `C5` of `names-held`, predicted by the engine before Excel ran it: cells of one letter and of two, near
 * A1 and far, those whose letter is R or C (R1C1's row and column alone), lower case, and R1C1's shapes with both parts.
 */
import { type EntryCase, entryCases } from './cases.js'

const let1 = (name: string) => `=LET(${name},1,${name})`

const GENERAL: Record<string, string[]> = {
  a1: ['A1', 'B2', 'Q4', 'D1', 'Z1', 'A2', 'A100', 'A10000', 'AA1', 'AB12', 'IU1', 'K3'].map(let1),
  rc: ['R2', 'C2', 'R10', 'C100', 'R1048576', 'C16384', 'r1', 'c5', 'R0', 'C0', 'R5C', 'RC5', 'R2C3'].map(let1),
}

export const buildEntryCases = (): EntryCase[] => entryCases(GENERAL)
