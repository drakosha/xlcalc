/**
 * Held-out cases of names (`run.ts spill names-held`): each answer predicted from the rules the set `names` found and
 * committed before Excel computed it (`PREDICTED`). The rules: a name's formula is computed where it is used, its
 * references without `$` moved by the cell from A1, its names the global ones; a union is a reference of several
 * areas (INDEX takes an area number); INDEX's reference and a spill reference stay references.
 */
import { one, sheet, type CellSpec, type SpillCase } from './cases.js'
import type { NameInput } from '../../../src/contract.js'
import { PREDICTED } from './names-held.predicted.js'

const DATA: Record<string, CellSpec> = {
  ...Object.fromEntries(Array.from({ length: 5 }, (_, i) => [`A${i + 1}`, i + 1])),
  ...Object.fromEntries(Array.from({ length: 5 }, (_, i) => [`B${i + 1}`, (i + 1) * 10])),
  ...Object.fromEntries(Array.from({ length: 5 }, (_, i) => [`C${i + 1}`, (i + 1) * 100])),
  D1: 1, E1: 2, D2: 3, E2: 4,
}

const held = (id: string, about: string, names: NameInput[], cells: Record<string, CellSpec>, ranges: string[], other?: Record<string, CellSpec>): SpillCase => ({
  ...one(id, about, { ...DATA, ...cells }, [], { book: { sheets: other ? [sheet('T', other)] : [], names } }),
  read: ranges.map((r) => (r.includes('!') ? { sheet: r.split('!')[0], range: r.split('!')[1] } : { sheet: 'S', range: r })),
  predicted: PREDICTED[id],
})

export function buildSpillCases(): SpillCase[] {
  return [
    held('relative-other-sheet', 'a relative name used on another sheet', [{ name: 'nxt', ref: "'S'!XFD1" }], {}, ['T!C4'], { C4: '=nxt' }),
    held('union-area', 'a union name: INDEX by area, AREAS, COUNT', [{ name: 'uu', ref: "'S'!$A$1:$B$2,'S'!$D$1:$E$2" }], { G1: '=INDEX(uu,2,2,2)', G2: '=AREAS(uu)', G3: '=SUM(INDEX(uu,0,0,1))', G4: '=COUNT(uu)' }, ['G1:G4']),
    held('array-arith', 'an array constant name under operators', [{ name: 'arr', ref: '{10,20,30}' }], { G1: '=arr/10', G3: '=SUM(arr*arr)' }, ['G1:I3']),
    held('index-criteria', "INDEX's reference as a name in criteria functions", [{ name: 'colb', ref: "INDEX('S'!$A$1:$C$5,0,2)" }], { G1: '=SUMIF(colb,">20")', G2: '=COUNTIF(colb,">=30")', G3: '=MAX(colb)' }, ['G1:G3']),
    held('qualified-local', "a global name over another sheet's local name", [{ name: 'k', ref: '1' }, { name: 'k', ref: '5', scope: 'T' }, { name: 'kk', ref: 'T!k*2' }], { G1: '=kk', G2: '=k' }, ['G1:G2'], { A1: '=k' }),
    held('let-in-name', 'LET in a name over a range', [{ name: 'sqa', ref: "LET(a,'S'!$A$1:$A$3,a*a)" }], { G1: '=SUM(sqa)', H1: '=sqa' }, ['G1:H3']),
    held('relative-chain', 'a name over a relative name', [{ name: 'lf', ref: "'S'!XFD1" }, { name: 'lften', ref: 'lf*10' }], { C3: '=lften' }, ['C3']),
    held('spill-ref', 'a name over a spill reference', [{ name: 'sp', ref: "'S'!$A$10#" }], { A10: '=SEQUENCE(3)', G1: '=SUM(sp)', G2: '=ROWS(sp)' }, ['G1:G2']),
    held('at-name', '@ before a name over a column', [{ name: 'nr', ref: "'S'!$A$1:$A$5" }], { G3: '=@nr', G7: '=@nr' }, ['G3', 'G7']),
    held('formula-twice', 'a name over a cell used twice in one formula', [{ name: 'half', ref: "'S'!$B$1/2" }], { G1: '=half+half' }, ['G1']),
  ]
}
